import win32com.client
import pythoncom

class AboveAverage:

    def __init__(self, aboveaverage=None):
        self.aboveaverage = aboveaverage

    @property
    def AboveBelow(self):
        return XlAboveBelow(self.aboveaverage.AboveBelow)

    @AboveBelow.setter
    def AboveBelow(self, value):
        self.aboveaverage.AboveBelow = value

    @property
    def Application(self):
        return self.aboveaverage.Application

    @property
    def AppliesTo(self):
        return Range(self.aboveaverage.AppliesTo)

    @property
    def Borders(self):
        return Borders(self.aboveaverage.Borders)

    @property
    def CalcFor(self):
        return XlCalcFor(self.aboveaverage.CalcFor)

    @CalcFor.setter
    def CalcFor(self, value):
        self.aboveaverage.CalcFor = value

    @property
    def Creator(self):
        return self.aboveaverage.Creator

    @property
    def Font(self):
        return Font(self.aboveaverage.Font)

    @property
    def Interior(self):
        return Interior(self.aboveaverage.Interior)

    @property
    def NumberFormat(self):
        return self.aboveaverage.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value):
        self.aboveaverage.NumberFormat = value

    @property
    def NumStdDev(self):
        return AboveAverage(self.aboveaverage.NumStdDev)

    @NumStdDev.setter
    def NumStdDev(self, value):
        self.aboveaverage.NumStdDev = value

    @property
    def Parent(self):
        return self.aboveaverage.Parent

    @property
    def Priority(self):
        return self.aboveaverage.Priority

    @Priority.setter
    def Priority(self, value):
        self.aboveaverage.Priority = value

    @property
    def PTCondition(self):
        return self.aboveaverage.PTCondition

    @property
    def ScopeType(self):
        return XlPivotConditionScope(self.aboveaverage.ScopeType)

    @ScopeType.setter
    def ScopeType(self, value):
        self.aboveaverage.ScopeType = value

    @property
    def StopIfTrue(self):
        return self.aboveaverage.StopIfTrue

    @StopIfTrue.setter
    def StopIfTrue(self, value):
        self.aboveaverage.StopIfTrue = value

    @property
    def Type(self):
        return XlFormatConditionType(self.aboveaverage.Type)

    def Delete(self):
        self.aboveaverage.Delete()

    def ModifyAppliesToRange(self, Range=None):
        params = [
            Range if Range is not None else pythoncom.Missing,
        ]
        self.aboveaverage.ModifyAppliesToRange(*params)

    def SetFirstPriority(self):
        self.aboveaverage.SetFirstPriority()

    def SetLastPriority(self):
        self.aboveaverage.SetLastPriority()


class Action:

    def __init__(self, action=None):
        self.action = action

    @property
    def Application(self):
        return self.action.Application

    @property
    def Caption(self):
        return Action(self.action.Caption)

    @property
    def Content(self):
        return Action(self.action.Content)

    @property
    def Coordinate(self):
        return Action(self.action.Coordinate)

    @property
    def Creator(self):
        return self.action.Creator

    @property
    def Name(self):
        return self.action.Name

    @property
    def Parent(self):
        return self.action.Parent

    @property
    def Type(self):
        return XlActionType(self.action.Type)

    def Execute(self):
        self.action.Execute()


class Actions:

    def __init__(self, actions=None):
        self.actions = actions

    def __call__(self, item):
        return Action(self.actions(item))

    @property
    def Application(self):
        return self.actions.Application

    @property
    def Count(self):
        return self.actions.Count

    @property
    def Creator(self):
        return self.actions.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.actions.Item):
            return Actions(self.actions.Item(*params))
        else:
            return Actions(self.actions.GetItem(*params))

    @property
    def Parent(self):
        return self.actions.Parent


class AddIn:

    def __init__(self, addin=None):
        self.addin = addin

    @property
    def Application(self):
        return self.addin.Application

    @property
    def CLSID(self):
        return self.addin.CLSID

    @property
    def Creator(self):
        return self.addin.Creator

    @property
    def FullName(self):
        return self.addin.FullName

    @property
    def Installed(self):
        return self.addin.Installed

    @Installed.setter
    def Installed(self, value):
        self.addin.Installed = value

    @property
    def IsOpen(self):
        return self.addin.IsOpen

    @property
    def Name(self):
        return self.addin.Name

    @property
    def Parent(self):
        return self.addin.Parent

    @property
    def Path(self):
        return self.addin.Path

    @property
    def progID(self):
        return self.addin.progID


class AddIns:

    def __init__(self, addins=None):
        self.addins = addins

    def __call__(self, item):
        return AddIn(self.addins(item))

    @property
    def Application(self):
        return self.addins.Application

    @property
    def Count(self):
        return self.addins.Count

    @property
    def Creator(self):
        return self.addins.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.addins.Item):
            return self.addins.Item(*params)
        else:
            return self.addins.GetItem(*params)

    @property
    def Parent(self):
        return self.addins.Parent

    def Add(self, FileName=None, CopyFile=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            CopyFile if CopyFile is not None else pythoncom.Missing,
        ]
        return AddIn(self.addins.Add(*params))


class AddIns2:

    def __init__(self, addins2=None):
        self.addins2 = addins2

    def __call__(self, item):
        return AddIns(self.addins2(item))

    @property
    def Application(self):
        return self.addins2.Application

    @property
    def Count(self):
        return self.addins2.Count

    @property
    def Creator(self):
        return self.addins2.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.addins2.Item):
            return self.addins2.Item(*params)
        else:
            return self.addins2.GetItem(*params)

    @property
    def Parent(self):
        return self.addins2.Parent

    def Add(self, FileName=None, CopyFile=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            CopyFile if CopyFile is not None else pythoncom.Missing,
        ]
        return AddIns(self.addins2.Add(*params))


class Adjustments:

    def __init__(self, adjustments=None):
        self.adjustments = adjustments

    @property
    def Application(self):
        return self.adjustments.Application

    @property
    def Count(self):
        return self.adjustments.Count

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


class AllowEditRange:

    def __init__(self, alloweditrange=None):
        self.alloweditrange = alloweditrange

    @property
    def Range(self):
        return Range(self.alloweditrange.Range)

    @property
    def Title(self):
        return self.alloweditrange.Title

    @Title.setter
    def Title(self, value):
        self.alloweditrange.Title = value

    @property
    def Users(self):
        return UserAccessList(self.alloweditrange.Users)

    def ChangePassword(self, Password=None):
        params = [
            Password if Password is not None else pythoncom.Missing,
        ]
        self.alloweditrange.ChangePassword(*params)

    def Delete(self):
        self.alloweditrange.Delete()

    def Unprotect(self, Password=None):
        params = [
            Password if Password is not None else pythoncom.Missing,
        ]
        self.alloweditrange.Unprotect(*params)


class AllowEditRanges:

    def __init__(self, alloweditranges=None):
        self.alloweditranges = alloweditranges

    def __call__(self, item):
        return AllowEditRange(self.alloweditranges(item))

    @property
    def Count(self):
        return self.alloweditranges.Count

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.alloweditranges.Item):
            return self.alloweditranges.Item(*params)
        else:
            return self.alloweditranges.GetItem(*params)

    def Add(self, Title=None, Range=None, Password=None):
        params = [
            Title if Title is not None else pythoncom.Missing,
            Range if Range is not None else pythoncom.Missing,
            Password if Password is not None else pythoncom.Missing,
        ]
        return AllowEditRange(self.alloweditranges.Add(*params))


class Application:

    def __init__(self, application=None):
        self.application = application

    def new(self):
        self.application = win32com.client.Dispatch("Excel.Application")
        return self

    @property
    def ActiveCell(self):
        return Range(self.application.ActiveCell)

    @property
    def ActiveChart(self):
        return Chart(self.application.ActiveChart)

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
        return ProtectedViewWindow(self.application.ActiveProtectedViewWindow)

    @property
    def ActiveSheet(self):
        return self.application.ActiveSheet

    @property
    def ActiveWindow(self):
        return Window(self.application.ActiveWindow)

    @property
    def ActiveWorkbook(self):
        return Workbook(self.application.ActiveWorkbook)

    @property
    def AddIns(self):
        return AddIns(self.application.AddIns)

    @property
    def AddIns2(self):
        return AddIns2(self.application.AddIns2)

    @property
    def AlertBeforeOverwriting(self):
        return self.application.AlertBeforeOverwriting

    @AlertBeforeOverwriting.setter
    def AlertBeforeOverwriting(self, value):
        self.application.AlertBeforeOverwriting = value

    @property
    def AltStartupPath(self):
        return self.application.AltStartupPath

    @AltStartupPath.setter
    def AltStartupPath(self, value):
        self.application.AltStartupPath = value

    @property
    def AlwaysUseClearType(self):
        return self.application.AlwaysUseClearType

    @AlwaysUseClearType.setter
    def AlwaysUseClearType(self, value):
        self.application.AlwaysUseClearType = value

    @property
    def Application(self):
        return self.application.Application

    @property
    def ArbitraryXMLSupportAvailable(self):
        return self.application.ArbitraryXMLSupportAvailable

    @property
    def AskToUpdateLinks(self):
        return self.application.AskToUpdateLinks

    @AskToUpdateLinks.setter
    def AskToUpdateLinks(self, value):
        self.application.AskToUpdateLinks = value

    @property
    def Assistance(self):
        return self.application.Assistance

    @property
    def AutoCorrect(self):
        return AutoCorrect(self.application.AutoCorrect)

    @property
    def AutoFormatAsYouTypeReplaceHyperlinks(self):
        return self.application.AutoFormatAsYouTypeReplaceHyperlinks

    @AutoFormatAsYouTypeReplaceHyperlinks.setter
    def AutoFormatAsYouTypeReplaceHyperlinks(self, value):
        self.application.AutoFormatAsYouTypeReplaceHyperlinks = value

    @property
    def AutomationSecurity(self):
        return self.application.AutomationSecurity

    @AutomationSecurity.setter
    def AutomationSecurity(self, value):
        self.application.AutomationSecurity = value

    @property
    def AutoPercentEntry(self):
        return self.application.AutoPercentEntry

    @AutoPercentEntry.setter
    def AutoPercentEntry(self, value):
        self.application.AutoPercentEntry = value

    @property
    def AutoRecover(self):
        return AutoRecover(self.application.AutoRecover)

    @property
    def Build(self):
        return self.application.Build

    @property
    def CalculateBeforeSave(self):
        return self.application.CalculateBeforeSave

    @CalculateBeforeSave.setter
    def CalculateBeforeSave(self, value):
        self.application.CalculateBeforeSave = value

    @property
    def Calculation(self):
        return XlCalculation(self.application.Calculation)

    @Calculation.setter
    def Calculation(self, value):
        self.application.Calculation = value

    @property
    def CalculationInterruptKey(self):
        return self.application.CalculationInterruptKey

    @CalculationInterruptKey.setter
    def CalculationInterruptKey(self, value):
        self.application.CalculationInterruptKey = value

    @property
    def CalculationState(self):
        return XlCalculationState(self.application.CalculationState)

    @property
    def CalculationVersion(self):
        return self.application.CalculationVersion

    def Caller(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.application.Caller):
            return self.application.Caller(*params)
        else:
            return self.application.GetCaller(*params)

    @property
    def CanPlaySounds(self):
        return self.application.CanPlaySounds

    @property
    def CanRecordSounds(self):
        return self.application.CanRecordSounds

    @property
    def Caption(self):
        return self.application.Caption

    @Caption.setter
    def Caption(self, value):
        self.application.Caption = value

    @property
    def CellDragAndDrop(self):
        return self.application.CellDragAndDrop

    @CellDragAndDrop.setter
    def CellDragAndDrop(self, value):
        self.application.CellDragAndDrop = value

    def Cells(self, RowIndex=None, ColumnIndex=None):
        params = [
            RowIndex if RowIndex is not None else pythoncom.Missing,
            ColumnIndex if ColumnIndex is not None else pythoncom.Missing,
        ]
        if callable(self.application.Cells):
            return Range(self.application.Cells(*params))
        else:
            return Range(self.application.GetCells(*params))

    @property
    def Charts(self):
        return Sheets(self.application.Charts)

    def ClipboardFormats(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.application.ClipboardFormats):
            return self.application.ClipboardFormats(*params)
        else:
            return self.application.GetClipboardFormats(*params)

    @property
    def ClusterConnector(self):
        return self.application.ClusterConnector

    @ClusterConnector.setter
    def ClusterConnector(self, value):
        self.application.ClusterConnector = value

    @property
    def Columns(self):
        return Range(self.application.Columns)

    @property
    def COMAddIns(self):
        return self.application.COMAddIns

    @property
    def CommandBars(self):
        return CommandBars(self.application.CommandBars)

    @property
    def CommandUnderlines(self):
        return XlCommandUnderlines(self.application.CommandUnderlines)

    @CommandUnderlines.setter
    def CommandUnderlines(self, value):
        self.application.CommandUnderlines = value

    @property
    def ConstrainNumeric(self):
        return self.application.ConstrainNumeric

    @ConstrainNumeric.setter
    def ConstrainNumeric(self, value):
        self.application.ConstrainNumeric = value

    @property
    def ControlCharacters(self):
        return self.application.ControlCharacters

    @ControlCharacters.setter
    def ControlCharacters(self, value):
        self.application.ControlCharacters = value

    @property
    def CopyObjectsWithCells(self):
        return self.application.CopyObjectsWithCells

    @CopyObjectsWithCells.setter
    def CopyObjectsWithCells(self, value):
        self.application.CopyObjectsWithCells = value

    @property
    def Creator(self):
        return self.application.Creator

    @property
    def Cursor(self):
        return XlMousePointer(self.application.Cursor)

    @Cursor.setter
    def Cursor(self, value):
        self.application.Cursor = value

    @property
    def CursorMovement(self):
        return self.application.CursorMovement

    @CursorMovement.setter
    def CursorMovement(self, value):
        self.application.CursorMovement = value

    @property
    def CustomListCount(self):
        return self.application.CustomListCount

    @property
    def CutCopyMode(self):
        return XLCutCopyMode(self.application.CutCopyMode)

    @CutCopyMode.setter
    def CutCopyMode(self, value):
        self.application.CutCopyMode = value

    @property
    def DataEntryMode(self):
        return self.application.DataEntryMode

    @DataEntryMode.setter
    def DataEntryMode(self, value):
        self.application.DataEntryMode = value

    @property
    def DDEAppReturnCode(self):
        return self.application.DDEAppReturnCode

    @property
    def DecimalSeparator(self):
        return self.application.DecimalSeparator

    @DecimalSeparator.setter
    def DecimalSeparator(self, value):
        self.application.DecimalSeparator = value

    @property
    def DefaultFilePath(self):
        return self.application.DefaultFilePath

    @DefaultFilePath.setter
    def DefaultFilePath(self, value):
        self.application.DefaultFilePath = value

    @property
    def DefaultSaveFormat(self):
        return self.application.DefaultSaveFormat

    @DefaultSaveFormat.setter
    def DefaultSaveFormat(self, value):
        self.application.DefaultSaveFormat = value

    @property
    def DefaultSheetDirection(self):
        return self.application.DefaultSheetDirection

    @DefaultSheetDirection.setter
    def DefaultSheetDirection(self, value):
        self.application.DefaultSheetDirection = value

    @property
    def DefaultWebOptions(self):
        return DefaultWebOptions(self.application.DefaultWebOptions)

    @property
    def DeferAsyncQueries(self):
        return self.application.DeferAsyncQueries

    @DeferAsyncQueries.setter
    def DeferAsyncQueries(self, value):
        self.application.DeferAsyncQueries = value

    @property
    def Dialogs(self):
        return Dialogs(self.application.Dialogs)

    @property
    def DisplayAlerts(self):
        return self.application.DisplayAlerts

    @DisplayAlerts.setter
    def DisplayAlerts(self, value):
        self.application.DisplayAlerts = value

    @property
    def DisplayClipboardWindow(self):
        return self.application.DisplayClipboardWindow

    @DisplayClipboardWindow.setter
    def DisplayClipboardWindow(self, value):
        self.application.DisplayClipboardWindow = value

    @property
    def DisplayCommentIndicator(self):
        return XlCommentDisplayMode(self.application.DisplayCommentIndicator)

    @DisplayCommentIndicator.setter
    def DisplayCommentIndicator(self, value):
        self.application.DisplayCommentIndicator = value

    @property
    def DisplayDocumentActionTaskPane(self):
        return self.application.DisplayDocumentActionTaskPane

    @DisplayDocumentActionTaskPane.setter
    def DisplayDocumentActionTaskPane(self, value):
        self.application.DisplayDocumentActionTaskPane = value

    @property
    def DisplayDocumentInformationPanel(self):
        return self.application.DisplayDocumentInformationPanel

    @DisplayDocumentInformationPanel.setter
    def DisplayDocumentInformationPanel(self, value):
        self.application.DisplayDocumentInformationPanel = value

    @property
    def DisplayExcel4Menus(self):
        return self.application.DisplayExcel4Menus

    @DisplayExcel4Menus.setter
    def DisplayExcel4Menus(self, value):
        self.application.DisplayExcel4Menus = value

    @property
    def DisplayFormulaAutoComplete(self):
        return self.application.DisplayFormulaAutoComplete

    @DisplayFormulaAutoComplete.setter
    def DisplayFormulaAutoComplete(self, value):
        self.application.DisplayFormulaAutoComplete = value

    @property
    def DisplayFormulaBar(self):
        return self.application.DisplayFormulaBar

    @DisplayFormulaBar.setter
    def DisplayFormulaBar(self, value):
        self.application.DisplayFormulaBar = value

    @property
    def DisplayFullScreen(self):
        return self.application.DisplayFullScreen

    @DisplayFullScreen.setter
    def DisplayFullScreen(self, value):
        self.application.DisplayFullScreen = value

    @property
    def DisplayFunctionToolTips(self):
        return self.application.DisplayFunctionToolTips

    @DisplayFunctionToolTips.setter
    def DisplayFunctionToolTips(self, value):
        self.application.DisplayFunctionToolTips = value

    @property
    def DisplayInsertOptions(self):
        return self.application.DisplayInsertOptions

    @DisplayInsertOptions.setter
    def DisplayInsertOptions(self, value):
        self.application.DisplayInsertOptions = value

    @property
    def DisplayNoteIndicator(self):
        return self.application.DisplayNoteIndicator

    @DisplayNoteIndicator.setter
    def DisplayNoteIndicator(self, value):
        self.application.DisplayNoteIndicator = value

    @property
    def DisplayPasteOptions(self):
        return self.application.DisplayPasteOptions

    @DisplayPasteOptions.setter
    def DisplayPasteOptions(self, value):
        self.application.DisplayPasteOptions = value

    @property
    def DisplayRecentFiles(self):
        return self.application.DisplayRecentFiles

    @DisplayRecentFiles.setter
    def DisplayRecentFiles(self, value):
        self.application.DisplayRecentFiles = value

    @property
    def DisplayScrollBars(self):
        return self.application.DisplayScrollBars

    @DisplayScrollBars.setter
    def DisplayScrollBars(self, value):
        self.application.DisplayScrollBars = value

    @property
    def DisplayStatusBar(self):
        return self.application.DisplayStatusBar

    @DisplayStatusBar.setter
    def DisplayStatusBar(self, value):
        self.application.DisplayStatusBar = value

    @property
    def EditDirectlyInCell(self):
        return self.application.EditDirectlyInCell

    @EditDirectlyInCell.setter
    def EditDirectlyInCell(self, value):
        self.application.EditDirectlyInCell = value

    @property
    def EnableAnimations(self):
        return self.application.EnableAnimations

    @property
    def EnableAutoComplete(self):
        return self.application.EnableAutoComplete

    @EnableAutoComplete.setter
    def EnableAutoComplete(self, value):
        self.application.EnableAutoComplete = value

    @property
    def EnableCancelKey(self):
        return self.application.EnableCancelKey

    @EnableCancelKey.setter
    def EnableCancelKey(self, value):
        self.application.EnableCancelKey = value

    @property
    def EnableEvents(self):
        return self.application.EnableEvents

    @EnableEvents.setter
    def EnableEvents(self, value):
        self.application.EnableEvents = value

    @property
    def EnableLargeOperationAlert(self):
        return self.application.EnableLargeOperationAlert

    @EnableLargeOperationAlert.setter
    def EnableLargeOperationAlert(self, value):
        self.application.EnableLargeOperationAlert = value

    @property
    def EnableLivePreview(self):
        return self.application.EnableLivePreview

    @EnableLivePreview.setter
    def EnableLivePreview(self, value):
        self.application.EnableLivePreview = value

    @property
    def EnableSound(self):
        return self.application.EnableSound

    @EnableSound.setter
    def EnableSound(self, value):
        self.application.EnableSound = value

    @property
    def ErrorCheckingOptions(self):
        return ErrorCheckingOptions(self.application.ErrorCheckingOptions)

    @property
    def Excel4IntlMacroSheets(self):
        return Sheets(self.application.Excel4IntlMacroSheets)

    @property
    def Excel4MacroSheets(self):
        return Sheets(self.application.Excel4MacroSheets)

    @property
    def ExtendList(self):
        return self.application.ExtendList

    @ExtendList.setter
    def ExtendList(self, value):
        self.application.ExtendList = value

    @property
    def FeatureInstall(self):
        return self.application.FeatureInstall

    @FeatureInstall.setter
    def FeatureInstall(self, value):
        self.application.FeatureInstall = value

    def FileConverters(self, Index1=None, Index2=None):
        params = [
            Index1 if Index1 is not None else pythoncom.Missing,
            Index2 if Index2 is not None else pythoncom.Missing,
        ]
        if callable(self.application.FileConverters):
            return self.application.FileConverters(*params)
        else:
            return self.application.GetFileConverters(*params)

    def FileDialog(self, fileDialogType=None):
        params = [
            fileDialogType if fileDialogType is not None else pythoncom.Missing,
        ]
        if callable(self.application.FileDialog):
            return self.application.FileDialog(*params)
        else:
            return self.application.GetFileDialog(*params)

    @property
    def FileExportConverters(self):
        return FileExportConverters(self.application.FileExportConverters)

    @property
    def FileValidation(self):
        return self.application.FileValidation

    @FileValidation.setter
    def FileValidation(self, value):
        self.application.FileValidation = value

    @property
    def FileValidationPivot(self):
        return self.application.FileValidationPivot

    @FileValidationPivot.setter
    def FileValidationPivot(self, value):
        self.application.FileValidationPivot = value

    @property
    def FindFormat(self):
        return self.application.FindFormat

    @property
    def FixedDecimal(self):
        return self.application.FixedDecimal

    @FixedDecimal.setter
    def FixedDecimal(self, value):
        self.application.FixedDecimal = value

    @property
    def FixedDecimalPlaces(self):
        return self.application.FixedDecimalPlaces

    @FixedDecimalPlaces.setter
    def FixedDecimalPlaces(self, value):
        self.application.FixedDecimalPlaces = value

    @property
    def FormulaBarHeight(self):
        return self.application.FormulaBarHeight

    @FormulaBarHeight.setter
    def FormulaBarHeight(self, value):
        self.application.FormulaBarHeight = value

    @property
    def GenerateGetPivotData(self):
        return self.application.GenerateGetPivotData

    @GenerateGetPivotData.setter
    def GenerateGetPivotData(self, value):
        self.application.GenerateGetPivotData = value

    @property
    def GenerateTableRefs(self):
        return self.application.GenerateTableRefs

    @GenerateTableRefs.setter
    def GenerateTableRefs(self, value):
        self.application.GenerateTableRefs = value

    @property
    def Height(self):
        return self.application.Height

    @Height.setter
    def Height(self, value):
        self.application.Height = value

    @property
    def HighQualityModeForGraphics(self):
        return self.application.HighQualityModeForGraphics

    @HighQualityModeForGraphics.setter
    def HighQualityModeForGraphics(self, value):
        self.application.HighQualityModeForGraphics = value

    @property
    def Hinstance(self):
        return Application(self.application.Hinstance)

    @property
    def HinstancePtr(self):
        return Application(self.application.HinstancePtr)

    @property
    def Hwnd(self):
        return self.application.Hwnd

    @property
    def IgnoreRemoteRequests(self):
        return self.application.IgnoreRemoteRequests

    @IgnoreRemoteRequests.setter
    def IgnoreRemoteRequests(self, value):
        self.application.IgnoreRemoteRequests = value

    @property
    def Interactive(self):
        return self.application.Interactive

    @Interactive.setter
    def Interactive(self, value):
        self.application.Interactive = value

    @property
    def IsSandboxed(self):
        return self.application.IsSandboxed

    @property
    def Iteration(self):
        return self.application.Iteration

    @Iteration.setter
    def Iteration(self, value):
        self.application.Iteration = value

    @property
    def LanguageSettings(self):
        return self.application.LanguageSettings

    @property
    def LargeOperationCellThousandCount(self):
        return self.application.LargeOperationCellThousandCount

    @LargeOperationCellThousandCount.setter
    def LargeOperationCellThousandCount(self, value):
        self.application.LargeOperationCellThousandCount = value

    @property
    def Left(self):
        return self.application.Left

    @Left.setter
    def Left(self, value):
        self.application.Left = value

    @property
    def LibraryPath(self):
        return self.application.LibraryPath

    @property
    def MailSession(self):
        return self.application.MailSession

    @property
    def MailSystem(self):
        return XlMailSystem(self.application.MailSystem)

    @property
    def MapPaperSize(self):
        return self.application.MapPaperSize

    @MapPaperSize.setter
    def MapPaperSize(self, value):
        self.application.MapPaperSize = value

    @property
    def MathCoprocessorAvailable(self):
        return self.application.MathCoprocessorAvailable

    @property
    def MaxChange(self):
        return self.application.MaxChange

    @MaxChange.setter
    def MaxChange(self, value):
        self.application.MaxChange = value

    @property
    def MaxIterations(self):
        return self.application.MaxIterations

    @MaxIterations.setter
    def MaxIterations(self, value):
        self.application.MaxIterations = value

    @property
    def MeasurementUnit(self):
        return self.application.MeasurementUnit

    @MeasurementUnit.setter
    def MeasurementUnit(self, value):
        self.application.MeasurementUnit = value

    @property
    def MouseAvailable(self):
        return self.application.MouseAvailable

    @property
    def MoveAfterReturn(self):
        return self.application.MoveAfterReturn

    @MoveAfterReturn.setter
    def MoveAfterReturn(self, value):
        self.application.MoveAfterReturn = value

    @property
    def MoveAfterReturnDirection(self):
        return XlDirection(self.application.MoveAfterReturnDirection)

    @MoveAfterReturnDirection.setter
    def MoveAfterReturnDirection(self, value):
        self.application.MoveAfterReturnDirection = value

    @property
    def MultiThreadedCalculation(self):
        return MultiThreadedCalculation(self.application.MultiThreadedCalculation)

    @property
    def Name(self):
        return self.application.Name

    @property
    def Names(self):
        return Names(self.application.Names)

    @property
    def NetworkTemplatesPath(self):
        return self.application.NetworkTemplatesPath

    @property
    def NewWorkbook(self):
        return self.application.NewWorkbook

    @property
    def ODBCErrors(self):
        return ODBCErrors(self.application.ODBCErrors)

    @property
    def ODBCTimeout(self):
        return self.application.ODBCTimeout

    @ODBCTimeout.setter
    def ODBCTimeout(self, value):
        self.application.ODBCTimeout = value

    @property
    def OLEDBErrors(self):
        return OLEDBErrors(self.application.OLEDBErrors)

    @property
    def OnWindow(self):
        return self.application.OnWindow

    @OnWindow.setter
    def OnWindow(self, value):
        self.application.OnWindow = value

    @property
    def OperatingSystem(self):
        return self.application.OperatingSystem

    @property
    def OrganizationName(self):
        return self.application.OrganizationName

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
    def PivotTableSelection(self):
        return self.application.PivotTableSelection

    @PivotTableSelection.setter
    def PivotTableSelection(self, value):
        self.application.PivotTableSelection = value

    def PreviousSelections(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.application.PreviousSelections):
            return Range(self.application.PreviousSelections(*params))
        else:
            return Range(self.application.GetPreviousSelections(*params))

    @property
    def PrintCommunication(self):
        return self.application.PrintCommunication

    @PrintCommunication.setter
    def PrintCommunication(self, value):
        self.application.PrintCommunication = value

    @property
    def ProductCode(self):
        return self.application.ProductCode

    @property
    def PromptForSummaryInfo(self):
        return self.application.PromptForSummaryInfo

    @PromptForSummaryInfo.setter
    def PromptForSummaryInfo(self, value):
        self.application.PromptForSummaryInfo = value

    @property
    def ProtectedViewWindows(self):
        return ProtectedViewWindows(self.application.ProtectedViewWindows)

    def Range(self, Cell1=None, Cell2=None):
        params = [
            Cell1 if Cell1 is not None else pythoncom.Missing,
            Cell2 if Cell2 is not None else pythoncom.Missing,
        ]
        if callable(self.application.Range):
            return Range(self.application.Range(*params))
        else:
            return Range(self.application.GetRange(*params))

    @property
    def Ready(self):
        return self.application.Ready

    @property
    def RecentFiles(self):
        return RecentFiles(self.application.RecentFiles)

    @property
    def RecordRelative(self):
        return self.application.RecordRelative

    @property
    def ReferenceStyle(self):
        return XlReferenceStyle(self.application.ReferenceStyle)

    @ReferenceStyle.setter
    def ReferenceStyle(self, value):
        self.application.ReferenceStyle = value

    def RegisteredFunctions(self, Index1=None, Index2=None):
        params = [
            Index1 if Index1 is not None else pythoncom.Missing,
            Index2 if Index2 is not None else pythoncom.Missing,
        ]
        if callable(self.application.RegisteredFunctions):
            return self.application.RegisteredFunctions(*params)
        else:
            return self.application.GetRegisteredFunctions(*params)

    @property
    def ReplaceFormat(self):
        return self.application.ReplaceFormat

    @property
    def RollZoom(self):
        return self.application.RollZoom

    @RollZoom.setter
    def RollZoom(self, value):
        self.application.RollZoom = value

    @property
    def Rows(self):
        return Range(self.application.Rows)

    @property
    def RTD(self):
        return RTD(self.application.RTD)

    @property
    def ScreenUpdating(self):
        return self.application.ScreenUpdating

    @ScreenUpdating.setter
    def ScreenUpdating(self, value):
        self.application.ScreenUpdating = value

    @property
    def Selection(self):
        return self.application.Selection

    @property
    def SensitivityLabelPolicy(self):
        return Application(self.application.SensitivityLabelPolicy)

    @property
    def Sheets(self):
        return Sheets(self.application.Sheets)

    @property
    def SheetsInNewWorkbook(self):
        return self.application.SheetsInNewWorkbook

    @SheetsInNewWorkbook.setter
    def SheetsInNewWorkbook(self, value):
        self.application.SheetsInNewWorkbook = value

    @property
    def ShowChartTipNames(self):
        return self.application.ShowChartTipNames

    @ShowChartTipNames.setter
    def ShowChartTipNames(self, value):
        self.application.ShowChartTipNames = value

    @property
    def ShowChartTipValues(self):
        return self.application.ShowChartTipValues

    @ShowChartTipValues.setter
    def ShowChartTipValues(self, value):
        self.application.ShowChartTipValues = value

    @property
    def ShowDevTools(self):
        return self.application.ShowDevTools

    @ShowDevTools.setter
    def ShowDevTools(self, value):
        self.application.ShowDevTools = value

    @property
    def ShowMenuFloaties(self):
        return self.application.ShowMenuFloaties

    @ShowMenuFloaties.setter
    def ShowMenuFloaties(self, value):
        self.application.ShowMenuFloaties = value

    @property
    def ShowSelectionFloaties(self):
        return self.application.ShowSelectionFloaties

    @ShowSelectionFloaties.setter
    def ShowSelectionFloaties(self, value):
        self.application.ShowSelectionFloaties = value

    @property
    def ShowStartupDialog(self):
        return self.application.ShowStartupDialog

    @ShowStartupDialog.setter
    def ShowStartupDialog(self, value):
        self.application.ShowStartupDialog = value

    @property
    def ShowToolTips(self):
        return self.application.ShowToolTips

    @ShowToolTips.setter
    def ShowToolTips(self, value):
        self.application.ShowToolTips = value

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
    def Speech(self):
        return Speech(self.application.Speech)

    @property
    def SpellingOptions(self):
        return SpellingOptions(self.application.SpellingOptions)

    @property
    def StandardFont(self):
        return self.application.StandardFont

    @StandardFont.setter
    def StandardFont(self, value):
        self.application.StandardFont = value

    @property
    def StandardFontSize(self):
        return self.application.StandardFontSize

    @StandardFontSize.setter
    def StandardFontSize(self, value):
        self.application.StandardFontSize = value

    @property
    def StartupPath(self):
        return self.application.StartupPath

    @property
    def StatusBar(self):
        return self.application.StatusBar

    @StatusBar.setter
    def StatusBar(self, value):
        self.application.StatusBar = value

    @property
    def TemplatesPath(self):
        return self.application.TemplatesPath

    @property
    def ThisCell(self):
        return Range(self.application.ThisCell)

    @property
    def ThisWorkbook(self):
        return Workbook(self.application.ThisWorkbook)

    @property
    def ThousandsSeparator(self):
        return self.application.ThousandsSeparator

    @ThousandsSeparator.setter
    def ThousandsSeparator(self, value):
        self.application.ThousandsSeparator = value

    @property
    def Top(self):
        return self.application.Top

    @Top.setter
    def Top(self, value):
        self.application.Top = value

    @property
    def TransitionMenuKey(self):
        return self.application.TransitionMenuKey

    @TransitionMenuKey.setter
    def TransitionMenuKey(self, value):
        self.application.TransitionMenuKey = value

    @property
    def TransitionMenuKeyAction(self):
        return self.application.TransitionMenuKeyAction

    @TransitionMenuKeyAction.setter
    def TransitionMenuKeyAction(self, value):
        self.application.TransitionMenuKeyAction = value

    @property
    def TransitionNavigKeys(self):
        return self.application.TransitionNavigKeys

    @TransitionNavigKeys.setter
    def TransitionNavigKeys(self, value):
        self.application.TransitionNavigKeys = value

    @property
    def UsableHeight(self):
        return self.application.UsableHeight

    @property
    def UsableWidth(self):
        return self.application.UsableWidth

    @property
    def UseClusterConnector(self):
        return self.application.UseClusterConnector

    @UseClusterConnector.setter
    def UseClusterConnector(self, value):
        self.application.UseClusterConnector = value

    @property
    def UsedObjects(self):
        return UsedObjects(self.application.UsedObjects)

    @property
    def UserControl(self):
        return self.application.UserControl

    @UserControl.setter
    def UserControl(self, value):
        self.application.UserControl = value

    @property
    def UserLibraryPath(self):
        return self.application.UserLibraryPath

    @property
    def UserName(self):
        return self.application.UserName

    @UserName.setter
    def UserName(self, value):
        self.application.UserName = value

    @property
    def UseSystemSeparators(self):
        return self.application.UseSystemSeparators

    @UseSystemSeparators.setter
    def UseSystemSeparators(self, value):
        self.application.UseSystemSeparators = value

    @property
    def Value(self):
        return self.application.Value

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
    def WarnOnFunctionNameConflict(self):
        return self.application.WarnOnFunctionNameConflict

    @WarnOnFunctionNameConflict.setter
    def WarnOnFunctionNameConflict(self, value):
        self.application.WarnOnFunctionNameConflict = value

    @property
    def Watches(self):
        return Watches(self.application.Watches)

    @property
    def Width(self):
        return self.application.Width

    @Width.setter
    def Width(self, value):
        self.application.Width = value

    @property
    def Windows(self):
        return Windows(self.application.Windows)

    @property
    def WindowsForPens(self):
        return self.application.WindowsForPens

    @property
    def WindowState(self):
        return XlWindowState(self.application.WindowState)

    @WindowState.setter
    def WindowState(self, value):
        self.application.WindowState = value

    @property
    def Workbooks(self):
        return Workbooks(self.application.Workbooks)

    @property
    def WorksheetFunction(self):
        return WorksheetFunction(self.application.WorksheetFunction)

    @property
    def Worksheets(self):
        return self.application.Worksheets

    def ActivateMicrosoftApp(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        self.application.ActivateMicrosoftApp(*params)

    def AddCustomList(self, ListArray=None, ByRow=None):
        params = [
            ListArray if ListArray is not None else pythoncom.Missing,
            ByRow if ByRow is not None else pythoncom.Missing,
        ]
        self.application.AddCustomList(*params)

    def Calculate(self):
        self.application.Calculate()

    def CalculateFull(self):
        self.application.CalculateFull()

    def CalculateFullRebuild(self):
        self.application.CalculateFullRebuild()

    def CalculateUntilAsyncQueriesDone(self):
        self.application.CalculateUntilAsyncQueriesDone()

    def CentimetersToPoints(self, Centimeters=None):
        params = [
            Centimeters if Centimeters is not None else pythoncom.Missing,
        ]
        return self.application.CentimetersToPoints(*params)

    def CheckAbort(self, KeepAbort=None):
        params = [
            KeepAbort if KeepAbort is not None else pythoncom.Missing,
        ]
        self.application.CheckAbort(*params)

    def CheckSpelling(self, Word=None, CustomDictionary=None, IgnoreUppercase=None):
        params = [
            Word if Word is not None else pythoncom.Missing,
            CustomDictionary if CustomDictionary is not None else pythoncom.Missing,
            IgnoreUppercase if IgnoreUppercase is not None else pythoncom.Missing,
        ]
        return self.application.CheckSpelling(*params)

    def ConvertFormula(self, Formula=None, FromReferenceStyle=None, ToReferenceStyle=None, ToAbsolute=None, RelativeTo=None):
        params = [
            Formula if Formula is not None else pythoncom.Missing,
            FromReferenceStyle if FromReferenceStyle is not None else pythoncom.Missing,
            ToReferenceStyle if ToReferenceStyle is not None else pythoncom.Missing,
            ToAbsolute if ToAbsolute is not None else pythoncom.Missing,
            RelativeTo if RelativeTo is not None else pythoncom.Missing,
        ]
        return self.application.ConvertFormula(*params)

    def DDEExecute(self, Channel=None, String=None):
        params = [
            Channel if Channel is not None else pythoncom.Missing,
            String if String is not None else pythoncom.Missing,
        ]
        self.application.DDEExecute(*params)

    def DDEInitiate(self, App=None, Topic=None):
        params = [
            App if App is not None else pythoncom.Missing,
            Topic if Topic is not None else pythoncom.Missing,
        ]
        return self.application.DDEInitiate(*params)

    def DDEPoke(self, Channel=None, Item=None, Data=None):
        params = [
            Channel if Channel is not None else pythoncom.Missing,
            Item if Item is not None else pythoncom.Missing,
            Data if Data is not None else pythoncom.Missing,
        ]
        self.application.DDEPoke(*params)

    def DDERequest(self, Channel=None, Item=None):
        params = [
            Channel if Channel is not None else pythoncom.Missing,
            Item if Item is not None else pythoncom.Missing,
        ]
        return self.application.DDERequest(*params)

    def DDETerminate(self, Channel=None):
        params = [
            Channel if Channel is not None else pythoncom.Missing,
        ]
        self.application.DDETerminate(*params)

    def DeleteCustomList(self, ListNum=None):
        params = [
            ListNum if ListNum is not None else pythoncom.Missing,
        ]
        self.application.DeleteCustomList(*params)

    def DisplayXMLSourcePane(self, XmlMap=None):
        params = [
            XmlMap if XmlMap is not None else pythoncom.Missing,
        ]
        self.application.DisplayXMLSourcePane(*params)

    def DoubleClick(self):
        self.application.DoubleClick()

    def Evaluate(self, Name=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
        ]
        return self.application.Evaluate(*params)

    def ExecuteExcel4Macro(self, String=None):
        params = [
            String if String is not None else pythoncom.Missing,
        ]
        return self.application.ExecuteExcel4Macro(*params)

    def FindFile(self):
        return self.application.FindFile()

    def GetCustomListContents(self, ListNum=None):
        params = [
            ListNum if ListNum is not None else pythoncom.Missing,
        ]
        return self.application.GetCustomListContents(*params)

    def GetCustomListNum(self, ListArray=None):
        params = [
            ListArray if ListArray is not None else pythoncom.Missing,
        ]
        return self.application.GetCustomListNum(*params)

    def GetOpenFilename(self, FileFilter=None, FilterIndex=None, Title=None, ButtonText=None, MultiSelect=None):
        params = [
            FileFilter if FileFilter is not None else pythoncom.Missing,
            FilterIndex if FilterIndex is not None else pythoncom.Missing,
            Title if Title is not None else pythoncom.Missing,
            ButtonText if ButtonText is not None else pythoncom.Missing,
            MultiSelect if MultiSelect is not None else pythoncom.Missing,
        ]
        return self.application.GetOpenFilename(*params)

    def GetPhonetic(self, Text=None):
        params = [
            Text if Text is not None else pythoncom.Missing,
        ]
        return self.application.GetPhonetic(*params)

    def GetSaveAsFilename(self, InitialFilename=None, FileFilter=None, FilterIndex=None, Title=None, ButtonText=None):
        params = [
            InitialFilename if InitialFilename is not None else pythoncom.Missing,
            FileFilter if FileFilter is not None else pythoncom.Missing,
            FilterIndex if FilterIndex is not None else pythoncom.Missing,
            Title if Title is not None else pythoncom.Missing,
            ButtonText if ButtonText is not None else pythoncom.Missing,
        ]
        return self.application.GetSaveAsFilename(*params)

    def Goto(self, Reference=None, Scroll=None):
        params = [
            Reference if Reference is not None else pythoncom.Missing,
            Scroll if Scroll is not None else pythoncom.Missing,
        ]
        self.application.Goto(*params)

    def Help(self, HelpFile=None, HelpContextID=None):
        params = [
            HelpFile if HelpFile is not None else pythoncom.Missing,
            HelpContextID if HelpContextID is not None else pythoncom.Missing,
        ]
        self.application.Help(*params)

    def InchesToPoints(self, Inches=None):
        params = [
            Inches if Inches is not None else pythoncom.Missing,
        ]
        return self.application.InchesToPoints(*params)

    def InputBox(self, Prompt=None, Title=None, Default=None, Left=None, Top=None, HelpFile=None, HelpContextID=None, Type=None):
        params = [
            Prompt if Prompt is not None else pythoncom.Missing,
            Title if Title is not None else pythoncom.Missing,
            Default if Default is not None else pythoncom.Missing,
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            HelpFile if HelpFile is not None else pythoncom.Missing,
            HelpContextID if HelpContextID is not None else pythoncom.Missing,
            Type if Type is not None else pythoncom.Missing,
        ]
        return self.application.InputBox(*params)

    def Intersect(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.application.Intersect(*params)

    def MacroOptions(self, Macro=None, Description=None, HasMenu=None, MenuText=None, HasShortcutKey=None, ShortcutKey=None, Category=None, StatusBar=None, HelpContextID=None, HelpFile=None, ArgumentDescriptions=None):
        params = [
            Macro if Macro is not None else pythoncom.Missing,
            Description if Description is not None else pythoncom.Missing,
            HasMenu if HasMenu is not None else pythoncom.Missing,
            MenuText if MenuText is not None else pythoncom.Missing,
            HasShortcutKey if HasShortcutKey is not None else pythoncom.Missing,
            ShortcutKey if ShortcutKey is not None else pythoncom.Missing,
            Category if Category is not None else pythoncom.Missing,
            StatusBar if StatusBar is not None else pythoncom.Missing,
            HelpContextID if HelpContextID is not None else pythoncom.Missing,
            HelpFile if HelpFile is not None else pythoncom.Missing,
            ArgumentDescriptions if ArgumentDescriptions is not None else pythoncom.Missing,
        ]
        self.application.MacroOptions(*params)

    def MailLogoff(self):
        self.application.MailLogoff()

    def MailLogon(self, Name=None, Password=None, DownloadNewMail=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
            Password if Password is not None else pythoncom.Missing,
            DownloadNewMail if DownloadNewMail is not None else pythoncom.Missing,
        ]
        self.application.MailLogon(*params)

    def NextLetter(self):
        return self.application.NextLetter()

    def OnKey(self, Key=None, Procedure=None):
        params = [
            Key if Key is not None else pythoncom.Missing,
            Procedure if Procedure is not None else pythoncom.Missing,
        ]
        self.application.OnKey(*params)

    def OnRepeat(self, Text=None, Procedure=None):
        params = [
            Text if Text is not None else pythoncom.Missing,
            Procedure if Procedure is not None else pythoncom.Missing,
        ]
        self.application.OnRepeat(*params)

    def OnTime(self, EarliestTime=None, Procedure=None, LatestTime=None, Schedule=None):
        params = [
            EarliestTime if EarliestTime is not None else pythoncom.Missing,
            Procedure if Procedure is not None else pythoncom.Missing,
            LatestTime if LatestTime is not None else pythoncom.Missing,
            Schedule if Schedule is not None else pythoncom.Missing,
        ]
        self.application.OnTime(*params)

    def OnUndo(self, Text=None, Procedure=None):
        params = [
            Text if Text is not None else pythoncom.Missing,
            Procedure if Procedure is not None else pythoncom.Missing,
        ]
        self.application.OnUndo(*params)

    def Quit(self):
        self.application.Quit()

    def RecordMacro(self, BasicCode=None, XlmCode=None):
        params = [
            BasicCode if BasicCode is not None else pythoncom.Missing,
            XlmCode if XlmCode is not None else pythoncom.Missing,
        ]
        self.application.RecordMacro(*params)

    def RegisterXLL(self, FileName=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
        ]
        return self.application.RegisterXLL(*params)

    def Repeat(self):
        self.application.Repeat()

    def Run(self, Macro=None, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Macro if Macro is not None else pythoncom.Missing,
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.application.Run(*params)

    def SaveWorkspace(self, FileName=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
        ]
        self.application.SaveWorkspace(*params)

    def SendKeys(self, Keys=None, Wait=None):
        params = [
            Keys if Keys is not None else pythoncom.Missing,
            Wait if Wait is not None else pythoncom.Missing,
        ]
        self.application.SendKeys(*params)

    def SharePointVersion(self, bstrUrl=None):
        params = [
            bstrUrl if bstrUrl is not None else pythoncom.Missing,
        ]
        return self.application.SharePointVersion(*params)

    def Undo(self):
        self.application.Undo()

    def Union(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.application.Union(*params)

    def Volatile(self, Volatile=None):
        params = [
            Volatile if Volatile is not None else pythoncom.Missing,
        ]
        self.application.Volatile(*params)

    def Wait(self, Time=None):
        params = [
            Time if Time is not None else pythoncom.Missing,
        ]
        return self.application.Wait(*params)


class Areas:

    def __init__(self, areas=None):
        self.areas = areas

    def __call__(self, item):
        return Area(self.areas(item))

    @property
    def Application(self):
        return self.areas.Application

    @property
    def Count(self):
        return self.areas.Count

    @property
    def Creator(self):
        return self.areas.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.areas.Item):
            return self.areas.Item(*params)
        else:
            return self.areas.GetItem(*params)

    @property
    def Parent(self):
        return self.areas.Parent


class Author:

    def __init__(self, author=None):
        self.author = author

    @property
    def Application(self):
        return self.author.Application

    @property
    def Creator(self):
        return self.author.Creator

    @property
    def Name(self):
        return self.author.Name

    @property
    def Parent(self):
        return self.author.Parent

    @property
    def ProviderID(self):
        return self.author.ProviderID

    @property
    def UserID(self):
        return self.author.UserID


class AutoCorrect:

    def __init__(self, autocorrect=None):
        self.autocorrect = autocorrect

    @property
    def Application(self):
        return self.autocorrect.Application

    @property
    def AutoExpandListRange(self):
        return self.autocorrect.AutoExpandListRange

    @AutoExpandListRange.setter
    def AutoExpandListRange(self, value):
        self.autocorrect.AutoExpandListRange = value

    @property
    def AutoFillFormulasInLists(self):
        return self.autocorrect.AutoFillFormulasInLists

    @AutoFillFormulasInLists.setter
    def AutoFillFormulasInLists(self, value):
        self.autocorrect.AutoFillFormulasInLists = value

    @property
    def CapitalizeNamesOfDays(self):
        return self.autocorrect.CapitalizeNamesOfDays

    @CapitalizeNamesOfDays.setter
    def CapitalizeNamesOfDays(self, value):
        self.autocorrect.CapitalizeNamesOfDays = value

    @property
    def CorrectCapsLock(self):
        return self.autocorrect.CorrectCapsLock

    @CorrectCapsLock.setter
    def CorrectCapsLock(self, value):
        self.autocorrect.CorrectCapsLock = value

    @property
    def CorrectSentenceCap(self):
        return self.autocorrect.CorrectSentenceCap

    @CorrectSentenceCap.setter
    def CorrectSentenceCap(self, value):
        self.autocorrect.CorrectSentenceCap = value

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
    def Parent(self):
        return self.autocorrect.Parent

    def ReplacementList(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.autocorrect.ReplacementList):
            return self.autocorrect.ReplacementList(*params)
        else:
            return self.autocorrect.GetReplacementList(*params)

    @property
    def ReplaceText(self):
        return self.autocorrect.ReplaceText

    @ReplaceText.setter
    def ReplaceText(self, value):
        self.autocorrect.ReplaceText = value

    @property
    def TwoInitialCapitals(self):
        return self.autocorrect.TwoInitialCapitals

    @TwoInitialCapitals.setter
    def TwoInitialCapitals(self, value):
        self.autocorrect.TwoInitialCapitals = value

    def AddReplacement(self, What=None, Replacement=None):
        params = [
            What if What is not None else pythoncom.Missing,
            Replacement if Replacement is not None else pythoncom.Missing,
        ]
        return self.autocorrect.AddReplacement(*params)

    def DeleteReplacement(self, What=None):
        params = [
            What if What is not None else pythoncom.Missing,
        ]
        return self.autocorrect.DeleteReplacement(*params)


class AutoFilter:

    def __init__(self, autofilter=None):
        self.autofilter = autofilter

    @property
    def Application(self):
        return self.autofilter.Application

    @property
    def Creator(self):
        return self.autofilter.Creator

    @property
    def FilterMode(self):
        return self.autofilter.FilterMode

    @property
    def Filters(self):
        return Filters(self.autofilter.Filters)

    @property
    def Parent(self):
        return self.autofilter.Parent

    @property
    def Range(self):
        return Range(self.autofilter.Range)

    @property
    def Sort(self):
        return self.autofilter.Sort

    def ApplyFilter(self):
        self.autofilter.ApplyFilter()

    def ShowAllData(self):
        self.autofilter.ShowAllData()


class AutoRecover:

    def __init__(self, autorecover=None):
        self.autorecover = autorecover

    @property
    def Application(self):
        return self.autorecover.Application

    @property
    def Creator(self):
        return self.autorecover.Creator

    @property
    def Enabled(self):
        return self.autorecover.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.autorecover.Enabled = value

    @property
    def Parent(self):
        return self.autorecover.Parent

    @property
    def Path(self):
        return self.autorecover.Path

    @Path.setter
    def Path(self, value):
        self.autorecover.Path = value

    @property
    def Time(self):
        return self.autorecover.Time

    @Time.setter
    def Time(self, value):
        self.autorecover.Time = value


class Axes:

    def __init__(self, axes=None):
        self.axes = axes

    def __call__(self, item):
        return Axe(self.axes(item))

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
        params = [
            Type if Type is not None else pythoncom.Missing,
            AxisGroup if AxisGroup is not None else pythoncom.Missing,
        ]
        return self.axes.Item(*params)


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
        return XlTimeUnit(self.axis.BaseUnit)

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
        return Border(self.axis.Border)

    @property
    def CategoryNames(self):
        return Range(self.axis.CategoryNames)

    @CategoryNames.setter
    def CategoryNames(self, value):
        self.axis.CategoryNames = value

    @property
    def CategoryType(self):
        return XlCategoryType(self.axis.CategoryType)

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
        return XlDisplayUnit(self.axis.DisplayUnit)

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
        return XlTickMark(self.axis.MajorTickMark)

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
        return CategoryType(self.axis.MajorUnitScale)

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
        return XlTickMark(self.axis.MinorTickMark)

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
        return CategoryType(self.axis.MinorUnitScale)

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
        return XlScaleType(self.axis.ScaleType)

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
        return XlAxisType(self.axis.Type)

    @property
    def Width(self):
        return self.axis.Width

    def Delete(self):
        return self.axis.Delete()

    def Select(self):
        return self.axis.Select()


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
        params = [
            Start if Start is not None else pythoncom.Missing,
            Length if Length is not None else pythoncom.Missing,
        ]
        if callable(self.axistitle.Characters):
            return Characters(self.axistitle.Characters(*params))
        else:
            return Characters(self.axistitle.GetCharacters(*params))

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
        return XlChartElementPosition(self.axistitle.Position)

    @Position.setter
    def Position(self, value):
        self.axistitle.Position = value

    @property
    def ReadingOrder(self):
        return self.axistitle.ReadingOrder

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
        return self.axistitle.Delete()

    def Select(self):
        return self.axistitle.Select()


class Border:

    def __init__(self, border=None):
        self.border = border

    @property
    def Application(self):
        return self.border.Application

    @property
    def Color(self):
        return RGB(self.border.Color)

    @Color.setter
    def Color(self, value):
        self.border.Color = value

    @property
    def ColorIndex(self):
        return self.border.ColorIndex

    @ColorIndex.setter
    def ColorIndex(self, value):
        self.border.ColorIndex = value

    @property
    def Creator(self):
        return self.border.Creator

    @property
    def LineStyle(self):
        return XlLineStyle(self.border.LineStyle)

    @LineStyle.setter
    def LineStyle(self, value):
        self.border.LineStyle = value

    @property
    def Parent(self):
        return self.border.Parent

    @property
    def ThemeColor(self):
        return self.border.ThemeColor

    @ThemeColor.setter
    def ThemeColor(self, value):
        self.border.ThemeColor = value

    @property
    def TintAndShade(self):
        return self.border.TintAndShade

    @TintAndShade.setter
    def TintAndShade(self, value):
        self.border.TintAndShade = value

    @property
    def Weight(self):
        return XlBorderWeight(self.border.Weight)

    @Weight.setter
    def Weight(self, value):
        self.border.Weight = value


class Borders:

    def __init__(self, borders=None):
        self.borders = borders

    def __call__(self, item):
        return Border(self.borders(item))

    @property
    def Application(self):
        return self.borders.Application

    @property
    def Color(self):
        return RGB(self.borders.Color)

    @Color.setter
    def Color(self, value):
        self.borders.Color = value

    @property
    def ColorIndex(self):
        return self.borders.ColorIndex

    @ColorIndex.setter
    def ColorIndex(self, value):
        self.borders.ColorIndex = value

    @property
    def Count(self):
        return self.borders.Count

    @property
    def Creator(self):
        return self.borders.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.borders.Item):
            return Border(self.borders.Item(*params))
        else:
            return Border(self.borders.GetItem(*params))

    @property
    def LineStyle(self):
        return XlLineStyle(self.borders.LineStyle)

    @LineStyle.setter
    def LineStyle(self, value):
        self.borders.LineStyle = value

    @property
    def Parent(self):
        return self.borders.Parent

    @property
    def ThemeColor(self):
        return self.borders.ThemeColor

    @ThemeColor.setter
    def ThemeColor(self, value):
        self.borders.ThemeColor = value

    @property
    def TintAndShade(self):
        return self.borders.TintAndShade

    @TintAndShade.setter
    def TintAndShade(self, value):
        self.borders.TintAndShade = value

    @property
    def Value(self):
        return self.borders.Value

    @property
    def Weight(self):
        return XlBorderWeight(self.borders.Weight)

    @Weight.setter
    def Weight(self, value):
        self.borders.Weight = value


class CalculatedFields:

    def __init__(self, calculatedfields=None):
        self.calculatedfields = calculatedfields

    def __call__(self, item):
        return CalculatedField(self.calculatedfields(item))

    @property
    def Application(self):
        return self.calculatedfields.Application

    @property
    def Count(self):
        return self.calculatedfields.Count

    @property
    def Creator(self):
        return self.calculatedfields.Creator

    @property
    def Parent(self):
        return self.calculatedfields.Parent

    def Add(self, Name=None, Formula=None, UseStandardFormula=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
            Formula if Formula is not None else pythoncom.Missing,
            UseStandardFormula if UseStandardFormula is not None else pythoncom.Missing,
        ]
        return CalculatedField(self.calculatedfields.Add(*params))

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return PivotField(self.calculatedfields.Item(*params))


class CalculatedItems:

    def __init__(self, calculateditems=None):
        self.calculateditems = calculateditems

    def __call__(self, item):
        return CalculatedItem(self.calculateditems(item))

    @property
    def Application(self):
        return self.calculateditems.Application

    @property
    def Count(self):
        return self.calculateditems.Count

    @property
    def Creator(self):
        return self.calculateditems.Creator

    @property
    def Parent(self):
        return self.calculateditems.Parent

    def Add(self, Name=None, Formula=None, UseStandardFormula=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
            Formula if Formula is not None else pythoncom.Missing,
            UseStandardFormula if UseStandardFormula is not None else pythoncom.Missing,
        ]
        return CalculatedItem(self.calculateditems.Add(*params))

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return PivotItem(self.calculateditems.Item(*params))


class CalculatedMember:

    def __init__(self, calculatedmember=None):
        self.calculatedmember = calculatedmember

    @property
    def Application(self):
        return self.calculatedmember.Application

    @property
    def Creator(self):
        return self.calculatedmember.Creator

    @property
    def DisplayFolder(self):
        return self.calculatedmember.DisplayFolder

    @property
    def Dynamic(self):
        return self.calculatedmember.Dynamic

    @property
    def FlattenHierarchies(self):
        return self.calculatedmember.FlattenHierarchies

    @FlattenHierarchies.setter
    def FlattenHierarchies(self, value):
        self.calculatedmember.FlattenHierarchies = value

    @property
    def Formula(self):
        return self.calculatedmember.Formula

    @property
    def HierarchizeDistinct(self):
        return self.calculatedmember.HierarchizeDistinct

    @HierarchizeDistinct.setter
    def HierarchizeDistinct(self, value):
        self.calculatedmember.HierarchizeDistinct = value

    @property
    def IsValid(self):
        return self.calculatedmember.IsValid

    @property
    def Name(self):
        return self.calculatedmember.Name

    @property
    def Parent(self):
        return self.calculatedmember.Parent

    @property
    def SolveOrder(self):
        return self.calculatedmember.SolveOrder

    @property
    def SourceName(self):
        return self.calculatedmember.SourceName

    @property
    def Type(self):
        return XlCalculatedMemberType(self.calculatedmember.Type)

    def Delete(self):
        self.calculatedmember.Delete()


class CalculatedMembers:

    def __init__(self, calculatedmembers=None):
        self.calculatedmembers = calculatedmembers

    def __call__(self, item):
        return CalculatedMember(self.calculatedmembers(item))

    @property
    def Application(self):
        return self.calculatedmembers.Application

    @property
    def Count(self):
        return self.calculatedmembers.Count

    @property
    def Creator(self):
        return self.calculatedmembers.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.calculatedmembers.Item):
            return self.calculatedmembers.Item(*params)
        else:
            return self.calculatedmembers.GetItem(*params)

    @property
    def Parent(self):
        return self.calculatedmembers.Parent

    def Add(self, Name=None, Formula=None, SolveOrder=None, Type=None, Dynamic=None, DisplayFolder=None, HierarchizeDistinct=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
            Formula if Formula is not None else pythoncom.Missing,
            SolveOrder if SolveOrder is not None else pythoncom.Missing,
            Type if Type is not None else pythoncom.Missing,
            Dynamic if Dynamic is not None else pythoncom.Missing,
            DisplayFolder if DisplayFolder is not None else pythoncom.Missing,
            HierarchizeDistinct if HierarchizeDistinct is not None else pythoncom.Missing,
        ]
        return CalculatedMember(self.calculatedmembers.Add(*params))


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
        return self.calloutformat.Application

    @property
    def AutoAttach(self):
        return self.calloutformat.AutoAttach

    @AutoAttach.setter
    def AutoAttach(self, value):
        self.calloutformat.AutoAttach = value

    @property
    def AutoLength(self):
        return self.calloutformat.AutoLength

    @AutoLength.setter
    def AutoLength(self, value):
        self.calloutformat.AutoLength = value

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

    def AutomaticLength(self):
        self.calloutformat.AutomaticLength()

    def CustomDrop(self, Drop=None):
        params = [
            Drop if Drop is not None else pythoncom.Missing,
        ]
        self.calloutformat.CustomDrop(*params)

    def CustomLength(self, Length=None):
        params = [
            Length if Length is not None else pythoncom.Missing,
        ]
        self.calloutformat.CustomLength(*params)

    def PresetDrop(self, DropType=None):
        params = [
            DropType if DropType is not None else pythoncom.Missing,
        ]
        self.calloutformat.PresetDrop(*params)


class CategoryCollection:

    def __init__(self, categorycollection=None):
        self.categorycollection = categorycollection

    @property
    def Application(self):
        return Application(self.categorycollection.Application)

    @property
    def Count(self):
        return self.categorycollection.Count

    @property
    def Creator(self):
        return self.categorycollection.Creator

    @property
    def Parent(self):
        return self.categorycollection.Parent

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return ChartCategory(self.categorycollection.Item(*params))


class CellFormat:

    def __init__(self, cellformat=None):
        self.cellformat = cellformat

    @property
    def AddIndent(self):
        return self.cellformat.AddIndent

    @AddIndent.setter
    def AddIndent(self, value):
        self.cellformat.AddIndent = value

    @property
    def Application(self):
        return self.cellformat.Application

    @property
    def Borders(self):
        return Borders(self.cellformat.Borders)

    @Borders.setter
    def Borders(self, value):
        self.cellformat.Borders = value

    @property
    def Creator(self):
        return self.cellformat.Creator

    @property
    def Font(self):
        return Font(self.cellformat.Font)

    @property
    def FormulaHidden(self):
        return self.cellformat.FormulaHidden

    @FormulaHidden.setter
    def FormulaHidden(self, value):
        self.cellformat.FormulaHidden = value

    @property
    def HorizontalAlignment(self):
        return self.cellformat.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.cellformat.HorizontalAlignment = value

    @property
    def IndentLevel(self):
        return self.cellformat.IndentLevel

    @IndentLevel.setter
    def IndentLevel(self, value):
        self.cellformat.IndentLevel = value

    @property
    def Interior(self):
        return Interior(self.cellformat.Interior)

    @property
    def Locked(self):
        return self.cellformat.Locked

    @Locked.setter
    def Locked(self, value):
        self.cellformat.Locked = value

    @property
    def MergeCells(self):
        return self.cellformat.MergeCells

    @MergeCells.setter
    def MergeCells(self, value):
        self.cellformat.MergeCells = value

    @property
    def NumberFormat(self):
        return self.cellformat.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value):
        self.cellformat.NumberFormat = value

    @property
    def NumberFormatLocal(self):
        return self.cellformat.NumberFormatLocal

    @NumberFormatLocal.setter
    def NumberFormatLocal(self, value):
        self.cellformat.NumberFormatLocal = value

    @property
    def Orientation(self):
        return self.cellformat.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.cellformat.Orientation = value

    @property
    def Parent(self):
        return self.cellformat.Parent

    @property
    def ShrinkToFit(self):
        return self.cellformat.ShrinkToFit

    @ShrinkToFit.setter
    def ShrinkToFit(self, value):
        self.cellformat.ShrinkToFit = value

    @property
    def VerticalAlignment(self):
        return self.cellformat.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.cellformat.VerticalAlignment = value

    @property
    def WrapText(self):
        return self.cellformat.WrapText

    @WrapText.setter
    def WrapText(self, value):
        self.cellformat.WrapText = value

    def Clear(self):
        self.cellformat.Clear()


class Characters:

    def __init__(self, characters=None):
        self.characters = characters

    @property
    def Application(self):
        return self.characters.Application

    @property
    def Caption(self):
        return self.characters.Caption

    @property
    def Count(self):
        return self.characters.Count

    @property
    def Creator(self):
        return self.characters.Creator

    @property
    def Font(self):
        return Font(self.characters.Font)

    @property
    def Parent(self):
        return self.characters.Parent

    @property
    def PhoneticCharacters(self):
        return Characters(self.characters.PhoneticCharacters)

    @PhoneticCharacters.setter
    def PhoneticCharacters(self, value):
        self.characters.PhoneticCharacters = value

    @property
    def Text(self):
        return self.characters.Text

    @Text.setter
    def Text(self, value):
        self.characters.Text = value

    def Delete(self):
        return self.characters.Delete()

    def Insert(self, String=None):
        params = [
            String if String is not None else pythoncom.Missing,
        ]
        return self.characters.Insert(*params)


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
        return XlBarShape(self.chart.BarShape)

    @BarShape.setter
    def BarShape(self, value):
        self.chart.BarShape = value

    @property
    def ChartArea(self):
        return ChartArea(self.chart.ChartArea)

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
        return XlChartType(self.chart.ChartType)

    @ChartType.setter
    def ChartType(self, value):
        self.chart.ChartType = value

    @property
    def CodeName(self):
        return self.chart.CodeName

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
        return XlDisplayBlanksAs(self.chart.DisplayBlanksAs)

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
    def Hyperlinks(self):
        return Hyperlinks(self.chart.Hyperlinks)

    @property
    def Index(self):
        return self.chart.Index

    @property
    def Legend(self):
        return Legend(self.chart.Legend)

    @property
    def MailEnvelope(self):
        return self.chart.MailEnvelope

    @property
    def Name(self):
        return self.chart.Name

    @Name.setter
    def Name(self, value):
        self.chart.Name = value

    @property
    def Next(self):
        return Worksheet(self.chart.Next)

    @property
    def PageSetup(self):
        return PageSetup(self.chart.PageSetup)

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
        return PivotLayout(self.chart.PivotLayout)

    @property
    def PlotArea(self):
        return PlotArea(self.chart.PlotArea)

    @property
    def PlotBy(self):
        return XlRowCol(self.chart.PlotBy)

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
    def Previous(self):
        return Worksheet(self.chart.Previous)

    @property
    def PrintedCommentPages(self):
        return self.chart.PrintedCommentPages

    @property
    def ProtectContents(self):
        return self.chart.ProtectContents

    @property
    def ProtectData(self):
        return self.chart.ProtectData

    @ProtectData.setter
    def ProtectData(self, value):
        self.chart.ProtectData = value

    @property
    def ProtectDrawingObjects(self):
        return self.chart.ProtectDrawingObjects

    @property
    def ProtectFormatting(self):
        return self.chart.ProtectFormatting

    @ProtectFormatting.setter
    def ProtectFormatting(self, value):
        self.chart.ProtectFormatting = value

    @property
    def ProtectionMode(self):
        return self.chart.ProtectionMode

    @property
    def ProtectSelection(self):
        return self.chart.ProtectSelection

    @ProtectSelection.setter
    def ProtectSelection(self, value):
        self.chart.ProtectSelection = value

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
        return Shapes(self.chart.Shapes)

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
    def Tab(self):
        return Tab(self.chart.Tab)

    @property
    def Visible(self):
        return XlSheetVisibility(self.chart.Visible)

    @Visible.setter
    def Visible(self, value):
        self.chart.Visible = value

    @property
    def Walls(self):
        return Walls(self.chart.Walls)

    def Activate(self):
        self.chart.Activate()

    def ApplyChartTemplate(self, FileName=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
        ]
        self.chart.ApplyChartTemplate(*params)

    def ApplyDataLabels(self, Type=None, LegendKey=None, AutoText=None, HasLeaderLines=None, ShowSeriesName=None, ShowCategoryName=None, ShowValue=None, ShowPercentage=None, ShowBubbleSize=None, Separator=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            LegendKey if LegendKey is not None else pythoncom.Missing,
            AutoText if AutoText is not None else pythoncom.Missing,
            HasLeaderLines if HasLeaderLines is not None else pythoncom.Missing,
            ShowSeriesName if ShowSeriesName is not None else pythoncom.Missing,
            ShowCategoryName if ShowCategoryName is not None else pythoncom.Missing,
            ShowValue if ShowValue is not None else pythoncom.Missing,
            ShowPercentage if ShowPercentage is not None else pythoncom.Missing,
            ShowBubbleSize if ShowBubbleSize is not None else pythoncom.Missing,
            Separator if Separator is not None else pythoncom.Missing,
        ]
        self.chart.ApplyDataLabels(*params)

    def ApplyLayout(self, Layout=None, ChartType=None):
        params = [
            Layout if Layout is not None else pythoncom.Missing,
            ChartType if ChartType is not None else pythoncom.Missing,
        ]
        self.chart.ApplyLayout(*params)

    def Axes(self, Type=None, AxisGroup=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            AxisGroup if AxisGroup is not None else pythoncom.Missing,
        ]
        return self.chart.Axes(*params)

    def ChartGroups(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.chart.ChartGroups(*params)

    def ChartObjects(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.chart.ChartObjects(*params)

    def ChartWizard(self, Source=None, Gallery=None, Format=None, PlotBy=None, CategoryLabels=None, SeriesLabels=None, HasLegend=None, Title=None, CategoryTitle=None, ValueTitle=None, ExtraTitle=None):
        params = [
            Source if Source is not None else pythoncom.Missing,
            Gallery if Gallery is not None else pythoncom.Missing,
            Format if Format is not None else pythoncom.Missing,
            PlotBy if PlotBy is not None else pythoncom.Missing,
            CategoryLabels if CategoryLabels is not None else pythoncom.Missing,
            SeriesLabels if SeriesLabels is not None else pythoncom.Missing,
            HasLegend if HasLegend is not None else pythoncom.Missing,
            Title if Title is not None else pythoncom.Missing,
            CategoryTitle if CategoryTitle is not None else pythoncom.Missing,
            ValueTitle if ValueTitle is not None else pythoncom.Missing,
            ExtraTitle if ExtraTitle is not None else pythoncom.Missing,
        ]
        self.chart.ChartWizard(*params)

    def CheckSpelling(self, CustomDictionary=None, IgnoreUppercase=None, AlwaysSuggest=None, SpellLang=None):
        params = [
            CustomDictionary if CustomDictionary is not None else pythoncom.Missing,
            IgnoreUppercase if IgnoreUppercase is not None else pythoncom.Missing,
            AlwaysSuggest if AlwaysSuggest is not None else pythoncom.Missing,
            SpellLang if SpellLang is not None else pythoncom.Missing,
        ]
        self.chart.CheckSpelling(*params)

    def ClearToMatchStyle(self):
        self.chart.ClearToMatchStyle()

    def Copy(self, Before=None, After=None):
        params = [
            Before if Before is not None else pythoncom.Missing,
            After if After is not None else pythoncom.Missing,
        ]
        self.chart.Copy(*params)

    def CopyPicture(self, Appearance=None, Format=None, Size=None):
        params = [
            Appearance if Appearance is not None else pythoncom.Missing,
            Format if Format is not None else pythoncom.Missing,
            Size if Size is not None else pythoncom.Missing,
        ]
        self.chart.CopyPicture(*params)

    def Delete(self):
        self.chart.Delete()

    def Evaluate(self, Name=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
        ]
        return self.chart.Evaluate(*params)

    def Export(self, FileName=None, FilterName=None, Interactive=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            FilterName if FilterName is not None else pythoncom.Missing,
            Interactive if Interactive is not None else pythoncom.Missing,
        ]
        return self.chart.Export(*params)

    def ExportAsFixedFormat(self, Type=None, FileName=None, Quality=None, IncludeDocProperties=None, IgnorePrintAreas=None, From=None, To=None, OpenAfterPublish=None, FixedFormatExtClassPtr=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            FileName if FileName is not None else pythoncom.Missing,
            Quality if Quality is not None else pythoncom.Missing,
            IncludeDocProperties if IncludeDocProperties is not None else pythoncom.Missing,
            IgnorePrintAreas if IgnorePrintAreas is not None else pythoncom.Missing,
            From if From is not None else pythoncom.Missing,
            To if To is not None else pythoncom.Missing,
            OpenAfterPublish if OpenAfterPublish is not None else pythoncom.Missing,
            FixedFormatExtClassPtr if FixedFormatExtClassPtr is not None else pythoncom.Missing,
        ]
        self.chart.ExportAsFixedFormat(*params)

    def GetChartElement(self, x=None, y=None, ElementID=None, Arg1=None, Arg2=None):
        params = [
            x if x is not None else pythoncom.Missing,
            y if y is not None else pythoncom.Missing,
            ElementID if ElementID is not None else pythoncom.Missing,
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        self.chart.GetChartElement(*params)

    def Location(self, Where=None, Name=None):
        params = [
            Where if Where is not None else pythoncom.Missing,
            Name if Name is not None else pythoncom.Missing,
        ]
        return self.chart.Location(*params)

    def Move(self, Before=None, After=None):
        params = [
            Before if Before is not None else pythoncom.Missing,
            After if After is not None else pythoncom.Missing,
        ]
        self.chart.Move(*params)

    def OLEObjects(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.chart.OLEObjects(*params)

    def Paste(self, Type=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
        ]
        self.chart.Paste(*params)

    def PrintOut(self, From=None, To=None, Copies=None, Preview=None, ActivePrinter=None, PrintToFile=None, Collate=None, PrToFileName=None, IgnorePrintAreas=None):
        params = [
            From if From is not None else pythoncom.Missing,
            To if To is not None else pythoncom.Missing,
            Copies if Copies is not None else pythoncom.Missing,
            Preview if Preview is not None else pythoncom.Missing,
            ActivePrinter if ActivePrinter is not None else pythoncom.Missing,
            PrintToFile if PrintToFile is not None else pythoncom.Missing,
            Collate if Collate is not None else pythoncom.Missing,
            PrToFileName if PrToFileName is not None else pythoncom.Missing,
            IgnorePrintAreas if IgnorePrintAreas is not None else pythoncom.Missing,
        ]
        return self.chart.PrintOut(*params)

    def PrintPreview(self, EnableChanges=None):
        params = [
            EnableChanges if EnableChanges is not None else pythoncom.Missing,
        ]
        self.chart.PrintPreview(*params)

    def Protect(self, Password=None, DrawingObjects=None, Contents=None, Scenarios=None, UserInterfaceOnly=None):
        params = [
            Password if Password is not None else pythoncom.Missing,
            DrawingObjects if DrawingObjects is not None else pythoncom.Missing,
            Contents if Contents is not None else pythoncom.Missing,
            Scenarios if Scenarios is not None else pythoncom.Missing,
            UserInterfaceOnly if UserInterfaceOnly is not None else pythoncom.Missing,
        ]
        self.chart.Protect(*params)

    def Refresh(self):
        self.chart.Refresh()

    def SaveAs(self, FileName=None, FileFormat=None, Password=None, WriteResPassword=None, ReadOnlyRecommended=None, CreateBackup=None, AddToMru=None, TextCodepage=None, TextVisualLayout=None, Local=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            FileFormat if FileFormat is not None else pythoncom.Missing,
            Password if Password is not None else pythoncom.Missing,
            WriteResPassword if WriteResPassword is not None else pythoncom.Missing,
            ReadOnlyRecommended if ReadOnlyRecommended is not None else pythoncom.Missing,
            CreateBackup if CreateBackup is not None else pythoncom.Missing,
            AddToMru if AddToMru is not None else pythoncom.Missing,
            TextCodepage if TextCodepage is not None else pythoncom.Missing,
            TextVisualLayout if TextVisualLayout is not None else pythoncom.Missing,
            Local if Local is not None else pythoncom.Missing,
        ]
        self.chart.SaveAs(*params)

    def SaveChartTemplate(self, FileName=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
        ]
        self.chart.SaveChartTemplate(*params)

    def Select(self, Replace=None):
        params = [
            Replace if Replace is not None else pythoncom.Missing,
        ]
        self.chart.Select(*params)

    def SeriesCollection(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.chart.SeriesCollection(*params)

    def SetBackgroundPicture(self, FileName=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
        ]
        self.chart.SetBackgroundPicture(*params)

    def SetDefaultChart(self, Name=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
        ]
        self.chart.SetDefaultChart(*params)

    def SetElement(self, Element=None):
        params = [
            Element if Element is not None else pythoncom.Missing,
        ]
        return self.chart.SetElement(*params)

    def SetSourceData(self, Source=None, PlotBy=None):
        params = [
            Source if Source is not None else pythoncom.Missing,
            PlotBy if PlotBy is not None else pythoncom.Missing,
        ]
        self.chart.SetSourceData(*params)

    def Unprotect(self, Password=None):
        params = [
            Password if Password is not None else pythoncom.Missing,
        ]
        self.chart.Unprotect(*params)


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
    def RoundedCorners(self):
        return self.chartarea.RoundedCorners

    @RoundedCorners.setter
    def RoundedCorners(self, value):
        self.chartarea.RoundedCorners = value

    @property
    def Shadow(self):
        return self.chartarea.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.chartarea.Shadow = value

    @property
    def Top(self):
        return self.chartarea.Top

    @property
    def Width(self):
        return self.chartarea.Width

    @Width.setter
    def Width(self, value):
        self.chartarea.Width = value

    def Clear(self):
        return self.chartarea.Clear()

    def ClearContents(self):
        return self.chartarea.ClearContents()

    def ClearFormats(self):
        return self.chartarea.ClearFormats()

    def Copy(self):
        return self.chartarea.Copy()

    def Select(self):
        return self.chartarea.Select()


class ChartColorFormat:

    def __init__(self, chartcolorformat=None):
        self.chartcolorformat = chartcolorformat


class ChartFillFormat:

    def __init__(self, chartfillformat=None):
        self.chartfillformat = chartfillformat


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
        return self.chartformat.Glow

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
        return self.chartformat.SoftEdge

    @property
    def TextFrame2(self):
        return TextFrame2(self.chartformat.TextFrame2)

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
        return ChartGroup(self.chartgroup.Has3DShading)

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
        return XlSizeRepresents(self.chartgroup.SizeRepresents)

    @SizeRepresents.setter
    def SizeRepresents(self, value):
        self.chartgroup.SizeRepresents = value

    @property
    def SplitType(self):
        return XlChartSplitType(self.chartgroup.SplitType)

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
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.chartgroup.SeriesCollection(*params)


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
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return ChartGroup(self.chartgroups.Item(*params))


class ChartObject:

    def __init__(self, chartobject=None):
        self.chartobject = chartobject

    @property
    def Application(self):
        return self.chartobject.Application

    @property
    def BottomRightCell(self):
        return Range(self.chartobject.BottomRightCell)

    @property
    def Chart(self):
        return Chart(self.chartobject.Chart)

    @property
    def Creator(self):
        return self.chartobject.Creator

    @property
    def Height(self):
        return self.chartobject.Height

    @Height.setter
    def Height(self, value):
        self.chartobject.Height = value

    @property
    def Index(self):
        return self.chartobject.Index

    @property
    def Left(self):
        return self.chartobject.Left

    @Left.setter
    def Left(self, value):
        self.chartobject.Left = value

    @property
    def Locked(self):
        return self.chartobject.Locked

    @Locked.setter
    def Locked(self, value):
        self.chartobject.Locked = value

    @property
    def Name(self):
        return self.chartobject.Name

    @property
    def Parent(self):
        return self.chartobject.Parent

    @property
    def Placement(self):
        return XlPlacement(self.chartobject.Placement)

    @Placement.setter
    def Placement(self, value):
        self.chartobject.Placement = value

    @property
    def PrintObject(self):
        return self.chartobject.PrintObject

    @PrintObject.setter
    def PrintObject(self, value):
        self.chartobject.PrintObject = value

    @property
    def ProtectChartObject(self):
        return self.chartobject.ProtectChartObject

    @ProtectChartObject.setter
    def ProtectChartObject(self, value):
        self.chartobject.ProtectChartObject = value

    @property
    def RoundedCorners(self):
        return self.chartobject.RoundedCorners

    @RoundedCorners.setter
    def RoundedCorners(self, value):
        self.chartobject.RoundedCorners = value

    @property
    def Shadow(self):
        return self.chartobject.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.chartobject.Shadow = value

    @property
    def ShapeRange(self):
        return ShapeRange(self.chartobject.ShapeRange)

    @property
    def Top(self):
        return self.chartobject.Top

    @Top.setter
    def Top(self, value):
        self.chartobject.Top = value

    @property
    def TopLeftCell(self):
        return Range(self.chartobject.TopLeftCell)

    @property
    def Visible(self):
        return self.chartobject.Visible

    @Visible.setter
    def Visible(self, value):
        self.chartobject.Visible = value

    @property
    def Width(self):
        return self.chartobject.Width

    @Width.setter
    def Width(self, value):
        self.chartobject.Width = value

    @property
    def ZOrder(self):
        return self.chartobject.ZOrder

    def Activate(self):
        return self.chartobject.Activate()

    def BringToFront(self):
        return self.chartobject.BringToFront()

    def Copy(self):
        return self.chartobject.Copy()

    def CopyPicture(self, Appearance=None, Format=None):
        params = [
            Appearance if Appearance is not None else pythoncom.Missing,
            Format if Format is not None else pythoncom.Missing,
        ]
        return self.chartobject.CopyPicture(*params)

    def Cut(self):
        return self.chartobject.Cut()

    def Delete(self):
        return self.chartobject.Delete()

    def Duplicate(self):
        return self.chartobject.Duplicate()

    def Select(self, Replace=None):
        params = [
            Replace if Replace is not None else pythoncom.Missing,
        ]
        return self.chartobject.Select(*params)

    def SendToBack(self):
        return self.chartobject.SendToBack()


class ChartObjects:

    def __init__(self, chartobjects=None):
        self.chartobjects = chartobjects

    def __call__(self, item):
        return ChartObject(self.chartobjects(item))

    @property
    def Application(self):
        return self.chartobjects.Application

    @property
    def Count(self):
        return self.chartobjects.Count

    @property
    def Creator(self):
        return self.chartobjects.Creator

    @property
    def Height(self):
        return self.chartobjects.Height

    @Height.setter
    def Height(self, value):
        self.chartobjects.Height = value

    @property
    def Left(self):
        return self.chartobjects.Left

    @Left.setter
    def Left(self, value):
        self.chartobjects.Left = value

    @property
    def Locked(self):
        return self.chartobjects.Locked

    @Locked.setter
    def Locked(self, value):
        self.chartobjects.Locked = value

    @property
    def Parent(self):
        return self.chartobjects.Parent

    @property
    def Placement(self):
        return XlPlacement(self.chartobjects.Placement)

    @Placement.setter
    def Placement(self, value):
        self.chartobjects.Placement = value

    @property
    def PrintObject(self):
        return self.chartobjects.PrintObject

    @PrintObject.setter
    def PrintObject(self, value):
        self.chartobjects.PrintObject = value

    @property
    def ProtectChartObject(self):
        return self.chartobjects.ProtectChartObject

    @ProtectChartObject.setter
    def ProtectChartObject(self, value):
        self.chartobjects.ProtectChartObject = value

    @property
    def ShapeRange(self):
        return ShapeRange(self.chartobjects.ShapeRange)

    @property
    def Top(self):
        return self.chartobjects.Top

    @Top.setter
    def Top(self, value):
        self.chartobjects.Top = value

    @property
    def Visible(self):
        return self.chartobjects.Visible

    @Visible.setter
    def Visible(self, value):
        self.chartobjects.Visible = value

    @property
    def Width(self):
        return self.chartobjects.Width

    @Width.setter
    def Width(self, value):
        self.chartobjects.Width = value

    def Add(self, Left=None, Top=None, Width=None, Height=None):
        params = [
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            Width if Width is not None else pythoncom.Missing,
            Height if Height is not None else pythoncom.Missing,
        ]
        return ChartObject(self.chartobjects.Add(*params))

    def Copy(self):
        return self.chartobjects.Copy()

    def CopyPicture(self, Appearance=None, Format=None):
        params = [
            Appearance if Appearance is not None else pythoncom.Missing,
            Format if Format is not None else pythoncom.Missing,
        ]
        return self.chartobjects.CopyPicture(*params)

    def Cut(self):
        return self.chartobjects.Cut()

    def Delete(self):
        return self.chartobjects.Delete()

    def Duplicate(self):
        return self.chartobjects.Duplicate()

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.chartobjects.Item(*params)

    def Select(self, Replace=None):
        params = [
            Replace if Replace is not None else pythoncom.Missing,
        ]
        return self.chartobjects.Select(*params)


class Charts:

    def __init__(self, charts=None):
        self.charts = charts

    def __call__(self, item):
        return Chart(self.charts(item))

    @property
    def Application(self):
        return self.charts.Application

    @property
    def Count(self):
        return self.charts.Count

    @property
    def Creator(self):
        return self.charts.Creator

    @property
    def HPageBreaks(self):
        return HPageBreaks(self.charts.HPageBreaks)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.charts.Item):
            return self.charts.Item(*params)
        else:
            return self.charts.GetItem(*params)

    @property
    def Parent(self):
        return self.charts.Parent

    @property
    def Visible(self):
        return self.charts.Visible

    @Visible.setter
    def Visible(self, value):
        self.charts.Visible = value

    @property
    def VPageBreaks(self):
        return VPageBreaks(self.charts.VPageBreaks)

    def Copy(self, Before=None, After=None):
        params = [
            Before if Before is not None else pythoncom.Missing,
            After if After is not None else pythoncom.Missing,
        ]
        self.charts.Copy(*params)

    def Delete(self):
        self.charts.Delete()

    def Move(self, Before=None, After=None):
        params = [
            Before if Before is not None else pythoncom.Missing,
            After if After is not None else pythoncom.Missing,
        ]
        self.charts.Move(*params)

    def PrintOut(self, From=None, To=None, Copies=None, Preview=None, ActivePrinter=None, PrintToFile=None, Collate=None, PrToFileName=None, IgnorePrintAreas=None):
        params = [
            From if From is not None else pythoncom.Missing,
            To if To is not None else pythoncom.Missing,
            Copies if Copies is not None else pythoncom.Missing,
            Preview if Preview is not None else pythoncom.Missing,
            ActivePrinter if ActivePrinter is not None else pythoncom.Missing,
            PrintToFile if PrintToFile is not None else pythoncom.Missing,
            Collate if Collate is not None else pythoncom.Missing,
            PrToFileName if PrToFileName is not None else pythoncom.Missing,
            IgnorePrintAreas if IgnorePrintAreas is not None else pythoncom.Missing,
        ]
        return self.charts.PrintOut(*params)

    def PrintPreview(self, EnableChanges=None):
        params = [
            EnableChanges if EnableChanges is not None else pythoncom.Missing,
        ]
        self.charts.PrintPreview(*params)

    def Select(self, Replace=None):
        params = [
            Replace if Replace is not None else pythoncom.Missing,
        ]
        self.charts.Select(*params)


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
        params = [
            Start if Start is not None else pythoncom.Missing,
            Length if Length is not None else pythoncom.Missing,
        ]
        if callable(self.charttitle.Characters):
            return Characters(self.charttitle.Characters(*params))
        else:
            return Characters(self.charttitle.GetCharacters(*params))

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
        return XlChartElementPosition(self.charttitle.Position)

    @Position.setter
    def Position(self, value):
        self.charttitle.Position = value

    @property
    def ReadingOrder(self):
        return self.charttitle.ReadingOrder

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
        return self.charttitle.Delete()

    def Select(self):
        return self.charttitle.Select()


class ChartView:

    def __init__(self, chartview=None):
        self.chartview = chartview

    @property
    def Application(self):
        return self.chartview.Application

    @property
    def Creator(self):
        return self.chartview.Creator

    @property
    def Parent(self):
        return self.chartview.Parent

    @property
    def Sheet(self):
        return ChartView(self.chartview.Sheet)


class ColorFormat:

    def __init__(self, colorformat=None):
        self.colorformat = colorformat

    @property
    def Application(self):
        return self.colorformat.Application

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
        return self.colorformat.ObjectThemeColor

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
    def SchemeColor(self):
        return self.colorformat.SchemeColor

    @SchemeColor.setter
    def SchemeColor(self, value):
        self.colorformat.SchemeColor = value

    @property
    def TintAndShade(self):
        return self.colorformat.TintAndShade

    @TintAndShade.setter
    def TintAndShade(self, value):
        self.colorformat.TintAndShade = value

    @property
    def Type(self):
        return self.colorformat.Type


class ColorScale:

    def __init__(self, colorscale=None):
        self.colorscale = colorscale

    @property
    def Application(self):
        return self.colorscale.Application

    @property
    def AppliesTo(self):
        return Range(self.colorscale.AppliesTo)

    @property
    def ColorScaleCriteria(self):
        return ColorScaleCriteria(self.colorscale.ColorScaleCriteria)

    @property
    def Creator(self):
        return self.colorscale.Creator

    @property
    def Formula(self):
        return self.colorscale.Formula

    @Formula.setter
    def Formula(self, value):
        self.colorscale.Formula = value

    @property
    def Parent(self):
        return self.colorscale.Parent

    @property
    def Priority(self):
        return self.colorscale.Priority

    @Priority.setter
    def Priority(self, value):
        self.colorscale.Priority = value

    @property
    def PTCondition(self):
        return self.colorscale.PTCondition

    @property
    def ScopeType(self):
        return XlPivotConditionScope(self.colorscale.ScopeType)

    @ScopeType.setter
    def ScopeType(self, value):
        self.colorscale.ScopeType = value

    @property
    def StopIfTrue(self):
        return self.colorscale.StopIfTrue

    @StopIfTrue.setter
    def StopIfTrue(self, value):
        self.colorscale.StopIfTrue = value

    @property
    def Type(self):
        return XlFormatConditionType(self.colorscale.Type)

    def Delete(self):
        self.colorscale.Delete()

    def ModifyAppliesToRange(self, Range=None):
        params = [
            Range if Range is not None else pythoncom.Missing,
        ]
        self.colorscale.ModifyAppliesToRange(*params)

    def SetFirstPriority(self):
        self.colorscale.SetFirstPriority()

    def SetLastPriority(self):
        self.colorscale.SetLastPriority()


class ColorScaleCriteria:

    def __init__(self, colorscalecriteria=None):
        self.colorscalecriteria = colorscalecriteria

    def __call__(self, item):
        return ColorScaleCriteri(self.colorscalecriteria(item))

    @property
    def Count(self):
        return self.colorscalecriteria.Count

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.colorscalecriteria.Item):
            return ColorScaleCriterion(self.colorscalecriteria.Item(*params))
        else:
            return ColorScaleCriterion(self.colorscalecriteria.GetItem(*params))


class ColorScaleCriterion:

    def __init__(self, colorscalecriterion=None):
        self.colorscalecriterion = colorscalecriterion

    @property
    def FormatColor(self):
        return FormatColor(self.colorscalecriterion.FormatColor)

    @property
    def Index(self):
        return self.colorscalecriterion.Index

    @property
    def Type(self):
        return XlConditionValueTypes(self.colorscalecriterion.Type)

    @property
    def Value(self):
        return self.colorscalecriterion.Value

    @Value.setter
    def Value(self, value):
        self.colorscalecriterion.Value = value


class ColorStop:

    def __init__(self, colorstop=None):
        self.colorstop = colorstop

    @property
    def Application(self):
        return self.colorstop.Application

    @property
    def Color(self):
        return self.colorstop.Color

    @Color.setter
    def Color(self, value):
        self.colorstop.Color = value

    @property
    def Creator(self):
        return self.colorstop.Creator

    @property
    def Parent(self):
        return self.colorstop.Parent

    @property
    def Position(self):
        return ColorStop(self.colorstop.Position)

    @Position.setter
    def Position(self, value):
        self.colorstop.Position = value

    @property
    def ThemeColor(self):
        return self.colorstop.ThemeColor

    @ThemeColor.setter
    def ThemeColor(self, value):
        self.colorstop.ThemeColor = value

    @property
    def TintAndShade(self):
        return self.colorstop.TintAndShade

    @TintAndShade.setter
    def TintAndShade(self, value):
        self.colorstop.TintAndShade = value

    def Delete(self):
        return self.colorstop.Delete()


class ColorStops:

    def __init__(self, colorstops=None):
        self.colorstops = colorstops

    def __call__(self, item):
        return ColorStop(self.colorstops(item))

    @property
    def Application(self):
        return self.colorstops.Application

    @property
    def Count(self):
        return self.colorstops.Count

    @Count.setter
    def Count(self, value):
        self.colorstops.Count = value

    @property
    def Creator(self):
        return self.colorstops.Creator

    @property
    def Parent(self):
        return self.colorstops.Parent

    def Add(self, Position=None):
        params = [
            Position if Position is not None else pythoncom.Missing,
        ]
        return ColorStop(self.colorstops.Add(*params))

    def Clear(self):
        return self.colorstops.Clear()

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.colorstops.Item(*params)


class Comment:

    def __init__(self, comment=None):
        self.comment = comment

    @property
    def Application(self):
        return self.comment.Application

    @property
    def Author(self):
        return self.comment.Author

    @property
    def Creator(self):
        return self.comment.Creator

    @property
    def Parent(self):
        return self.comment.Parent

    @property
    def Shape(self):
        return Shape(self.comment.Shape)

    @property
    def Visible(self):
        return self.comment.Visible

    @Visible.setter
    def Visible(self, value):
        self.comment.Visible = value

    @property
    def Application(self):
        return self.comment.Application

    def Delete(self):
        self.comment.Delete()

    def Next(self):
        return self.comment.Next()

    def Previous(self):
        return self.comment.Previous()

    def Text(self, Text=None, Start=None, Overwrite=None):
        params = [
            Text if Text is not None else pythoncom.Missing,
            Start if Start is not None else pythoncom.Missing,
            Overwrite if Overwrite is not None else pythoncom.Missing,
        ]
        return self.comment.Text(*params)


class Comments:

    def __init__(self, comments=None):
        self.comments = comments

    def __call__(self, item):
        return Comment(self.comments(item))

    @property
    def Application(self):
        return self.comments.Application

    @property
    def Count(self):
        return self.comments.Count

    @property
    def Creator(self):
        return self.comments.Creator

    @property
    def Parent(self):
        return self.comments.Parent

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return Comment(self.comments.Item(*params))


class CommentsThreaded:

    def __init__(self, commentsthreaded=None):
        self.commentsthreaded = commentsthreaded

    def __call__(self, item):
        return CommentsThreade(self.commentsthreaded(item))

    @property
    def Application(self):
        return self.commentsthreaded.Application

    @property
    def Count(self):
        return self.commentsthreaded.Count

    @property
    def Creator(self):
        return self.commentsthreaded.Creator

    @property
    def Parent(self):
        return self.commentsthreaded.Parent

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return CommentThreaded(self.commentsthreaded.Item(*params))


class CommentThreaded:

    def __init__(self, commentthreaded=None):
        self.commentthreaded = commentthreaded

    @property
    def Creator(self):
        return self.commentthreaded.Creator

    @property
    def Parent(self):
        return self.commentthreaded.Parent

    @property
    def Parent(self):
        return self.commentthreaded.Parent

    @property
    def Replies(self):
        return self.commentthreaded.Replies

    def AddReply(self, Text=None):
        params = [
            Text if Text is not None else pythoncom.Missing,
        ]
        return self.commentthreaded.AddReply(*params)

    def Delete(self):
        self.commentthreaded.Delete()

    def Next(self):
        return self.commentthreaded.Next()

    def Previous(self):
        return self.commentthreaded.Previous()

    def Text(self, Text=None, Start=None, Overwrite=None):
        params = [
            Text if Text is not None else pythoncom.Missing,
            Start if Start is not None else pythoncom.Missing,
            Overwrite if Overwrite is not None else pythoncom.Missing,
        ]
        return self.commentthreaded.Text(*params)


class ConditionValue:

    def __init__(self, conditionvalue=None):
        self.conditionvalue = conditionvalue

    @property
    def Application(self):
        return self.conditionvalue.Application

    @property
    def Creator(self):
        return self.conditionvalue.Creator

    @property
    def Parent(self):
        return self.conditionvalue.Parent

    @property
    def Type(self):
        return XlConditionValueTypes(self.conditionvalue.Type)

    @property
    def Value(self):
        return self.conditionvalue.Value

    @Value.setter
    def Value(self, value):
        self.conditionvalue.Value = value

    def Modify(self, NewType=None, NewValue=None):
        params = [
            NewType if NewType is not None else pythoncom.Missing,
            NewValue if NewValue is not None else pythoncom.Missing,
        ]
        self.conditionvalue.Modify(*params)


class Connections:

    def __init__(self, connections=None):
        self.connections = connections

    def __call__(self, item):
        return Connection(self.connections(item))

    @property
    def Application(self):
        return self.connections.Application

    @property
    def Count(self):
        return self.connections.Count

    @property
    def Creator(self):
        return self.connections.Creator

    @property
    def Parent(self):
        return self.connections.Parent

    def Add(self, Name=None, Description=None, ConnectionString=None, CommandText=None, lCmdtype=None, CreateModelConnection=None, ImportRelationships=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
            Description if Description is not None else pythoncom.Missing,
            ConnectionString if ConnectionString is not None else pythoncom.Missing,
            CommandText if CommandText is not None else pythoncom.Missing,
            lCmdtype if lCmdtype is not None else pythoncom.Missing,
            CreateModelConnection if CreateModelConnection is not None else pythoncom.Missing,
            ImportRelationships if ImportRelationships is not None else pythoncom.Missing,
        ]
        return Connection(self.connections.Add(*params))

    def AddFromFile(self, FileName=None, CreateModelConnection=None, ImportRelationships=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            CreateModelConnection if CreateModelConnection is not None else pythoncom.Missing,
            ImportRelationships if ImportRelationships is not None else pythoncom.Missing,
        ]
        return self.connections.AddFromFile(*params)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.connections.Item(*params)


class ConnectorFormat:

    def __init__(self, connectorformat=None):
        self.connectorformat = connectorformat

    @property
    def Application(self):
        return self.connectorformat.Application

    @property
    def BeginConnected(self):
        return self.connectorformat.BeginConnected

    @property
    def BeginConnectedShape(self):
        return Shape(self.connectorformat.BeginConnectedShape)

    @property
    def BeginConnectionSite(self):
        return self.connectorformat.BeginConnectionSite

    @property
    def Creator(self):
        return self.connectorformat.Creator

    @property
    def EndConnected(self):
        return self.connectorformat.EndConnected

    @property
    def EndConnectedShape(self):
        return Shape(self.connectorformat.EndConnectedShape)

    @property
    def EndConnectionSite(self):
        return self.connectorformat.EndConnectionSite

    @property
    def Parent(self):
        return self.connectorformat.Parent

    @property
    def Type(self):
        return self.connectorformat.Type

    @Type.setter
    def Type(self, value):
        self.connectorformat.Type = value

    def BeginConnect(self, ConnectedShape=None, ConnectionSite=None):
        params = [
            ConnectedShape if ConnectedShape is not None else pythoncom.Missing,
            ConnectionSite if ConnectionSite is not None else pythoncom.Missing,
        ]
        self.connectorformat.BeginConnect(*params)

    def BeginDisconnect(self):
        self.connectorformat.BeginDisconnect()

    def EndConnect(self, ConnectedShape=None, ConnectionSite=None):
        params = [
            ConnectedShape if ConnectedShape is not None else pythoncom.Missing,
            ConnectionSite if ConnectionSite is not None else pythoncom.Missing,
        ]
        self.connectorformat.EndConnect(*params)

    def EndDisconnect(self):
        self.connectorformat.EndDisconnect()


# Constants enumeration
xl3DBar = -4099
xl3DEffects1 = 13
xl3DEffects2 = 14
xl3DSurface = -4103
xlAbove = 0
xlAccounting1 = 4
xlAccounting2 = 5
xlAccounting4 = 17
xlAdd = 2
xlAll = -4104
xlAccounting3 = 6
xlAllExceptBorders = 7
xlAutomatic = -4105
xlBar = 2
xlBelow = 1
xlBidi = -5000
xlBidiCalendar = 3
xlBoth = 1
xlBottom = -4107
xlCascade = 7
xlCenter = -4108
xlCenterAcrossSelection = 7
xlChart4 = 2
xlChartSeries = 17
xlChartShort = 6
xlChartTitles = 18
xlChecker = 9
xlCircle = 8
xlClassic1 = 1
xlClassic2 = 2
xlClassic3 = 3
xlClosed = 3
xlColor1 = 7
xlColor2 = 8
xlColor3 = 9
xlColumn = 3
xlCombination = -4111
xlComplete = 4
xlConstants = 2
xlContents = 2
xlContext = -5002
xlCorner = 2
xlCrissCross = 16
xlCross = 4
xlCustom = -4114
xlDebugCodePane = 13
xlDefaultAutoFormat = -1
xlDesktop = 9
xlDiamond = 2
xlDirect = 1
xlDistributed = -4117
xlDivide = 5
xlDoubleAccounting = 5
xlDoubleClosed = 5
xlDoubleOpen = 4
xlDoubleQuote = 1
xlDrawingObject = 14
xlEntireChart = 20
xlExcelMenus = 1
xlExtended = 3
xlFill = 5
xlFirst = 0
xlFixedValue = 1
xlFloating = 5
xlFormats = -4122
xlFormula = 5
xlFullScript = 1
xlGeneral = 1
xlGray16 = 17
xlGray25 = -4124
xlGray50 = -4125
xlGray75 = -4126
xlGray8 = 18
xlGregorian = 2
xlGrid = 15
xlGridline = 22
xlHigh = -4127
xlHindiNumerals = 3
xlIcons = 1
xlImmediatePane = 12
xlInside = 2
xlInteger = 2
xlJustify = -4130
xlLast = 1
xlLastCell = 11
xlLatin = -5001
xlLeft = -4131
xlLeftToRight = 2
xlLightDown = 13
xlLightHorizontal = 11
xlLightUp = 14
xlLightVertical = 12
xlList1 = 10
xlList2 = 11
xlList3 = 12
xlLocalFormat1 = 15
xlLocalFormat2 = 16
xlLogicalCursor = 1
xlLong = 3
xlLotusHelp = 2
xlLow = -4134
xlLTR = -5003
xlMacrosheetCell = 7
xlManual = -4135
xlMaximum = 2
xlMinimum = 4
xlMinusValues = 3
xlMixed = 2
xlMixedAuthorizedScript = 4
xlMixedScript = 3
xlModule = -4141
xlMultiply = 4
xlNarrow = 1
xlNextToAxis = 4
xlNoDocuments = 3
xlNone = -4142
xlNotes = -4144
xlOff = -4146
xlOn = 1
xlOpaque = 3
xlOpen = 2
xlOutside = 3
xlPartial = 3
xlPartialScript = 2
xlPercent = 2
xlPlus = 9
xlPlusValues = 2
xlReference = 4
xlRight = -4152
xlRTL = -5004
xlScale = 3
xlSemiautomatic = 2
xlSemiGray75 = 10
xlShort = 1
xlShowLabel = 4
xlShowLabelAndPercent = 5
xlShowPercent = 3
xlShowValue = 2
xlSimple = -4154
xlSingle = 2
xlSingleAccounting = 4
xlSingleQuote = 2
xlSolid = 1
xlSquare = 1
xlStar = 5
xlStError = 4
xlStrict = 2
xlSubtract = 3
xlSystem = 1
xlTextBox = 16
xlTiled = 1
xlTitleBar = 8
xlToolbar = 1
xlToolbarButton = 2
xlTop = -4160
xlTopToBottom = 1
xlTransparent = 2
xlTriangle = 3
xlVeryHidden = 2
xlVisible = 12
xlVisualCursor = 2
xlWatchPane = 11
xlWide = 3
xlWorkbookTab = 6
xlWorksheet4 = 1
xlWorksheetCell = 3
xlWorksheetShort = 5

class ControlFormat:

    def __init__(self, controlformat=None):
        self.controlformat = controlformat

    @property
    def Application(self):
        return self.controlformat.Application

    @property
    def Creator(self):
        return self.controlformat.Creator

    @property
    def DropDownLines(self):
        return self.controlformat.DropDownLines

    @DropDownLines.setter
    def DropDownLines(self, value):
        self.controlformat.DropDownLines = value

    @property
    def Enabled(self):
        return self.controlformat.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.controlformat.Enabled = value

    @property
    def LargeChange(self):
        return self.controlformat.LargeChange

    @LargeChange.setter
    def LargeChange(self, value):
        self.controlformat.LargeChange = value

    @property
    def LinkedCell(self):
        return self.controlformat.LinkedCell

    @LinkedCell.setter
    def LinkedCell(self, value):
        self.controlformat.LinkedCell = value

    @property
    def ListCount(self):
        return self.controlformat.ListCount

    @property
    def ListFillRange(self):
        return self.controlformat.ListFillRange

    @ListFillRange.setter
    def ListFillRange(self, value):
        self.controlformat.ListFillRange = value

    @property
    def ListIndex(self):
        return self.controlformat.ListIndex

    @ListIndex.setter
    def ListIndex(self, value):
        self.controlformat.ListIndex = value

    @property
    def LockedText(self):
        return self.controlformat.LockedText

    @LockedText.setter
    def LockedText(self, value):
        self.controlformat.LockedText = value

    @property
    def Max(self):
        return self.controlformat.Max

    @Max.setter
    def Max(self, value):
        self.controlformat.Max = value

    @property
    def Min(self):
        return self.controlformat.Min

    @Min.setter
    def Min(self, value):
        self.controlformat.Min = value

    @property
    def MultiSelect(self):
        return self.controlformat.MultiSelect

    @MultiSelect.setter
    def MultiSelect(self, value):
        self.controlformat.MultiSelect = value

    @property
    def Parent(self):
        return self.controlformat.Parent

    @property
    def PrintObject(self):
        return self.controlformat.PrintObject

    @PrintObject.setter
    def PrintObject(self, value):
        self.controlformat.PrintObject = value

    @property
    def SmallChange(self):
        return self.controlformat.SmallChange

    @SmallChange.setter
    def SmallChange(self, value):
        self.controlformat.SmallChange = value

    @property
    def Value(self):
        return self.controlformat.Value

    @Value.setter
    def Value(self, value):
        self.controlformat.Value = value

    def AddItem(self, Text=None, Index=None):
        params = [
            Text if Text is not None else pythoncom.Missing,
            Index if Index is not None else pythoncom.Missing,
        ]
        self.controlformat.AddItem(*params)

    def List(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.controlformat.List(*params)

    def RemoveAllItems(self):
        self.controlformat.RemoveAllItems()

    def RemoveItem(self, Index=None, Count=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
            Count if Count is not None else pythoncom.Missing,
        ]
        self.controlformat.RemoveItem(*params)


class CubeField:

    def __init__(self, cubefield=None):
        self.cubefield = cubefield

    @property
    def AllItemsVisible(self):
        return self.cubefield.AllItemsVisible

    @property
    def Application(self):
        return self.cubefield.Application

    @property
    def Caption(self):
        return self.cubefield.Caption

    @property
    def Creator(self):
        return self.cubefield.Creator

    @property
    def CubeFieldSubType(self):
        return self.cubefield.CubeFieldSubType

    @property
    def CubeFieldType(self):
        return self.cubefield.CubeFieldType

    @property
    def CurrentPageName(self):
        return self.cubefield.CurrentPageName

    @CurrentPageName.setter
    def CurrentPageName(self, value):
        self.cubefield.CurrentPageName = value

    @property
    def DragToColumn(self):
        return self.cubefield.DragToColumn

    @DragToColumn.setter
    def DragToColumn(self, value):
        self.cubefield.DragToColumn = value

    @property
    def DragToData(self):
        return self.cubefield.DragToData

    @DragToData.setter
    def DragToData(self, value):
        self.cubefield.DragToData = value

    @property
    def DragToHide(self):
        return self.cubefield.DragToHide

    @DragToHide.setter
    def DragToHide(self, value):
        self.cubefield.DragToHide = value

    @property
    def DragToPage(self):
        return self.cubefield.DragToPage

    @DragToPage.setter
    def DragToPage(self, value):
        self.cubefield.DragToPage = value

    @property
    def DragToRow(self):
        return self.cubefield.DragToRow

    @DragToRow.setter
    def DragToRow(self, value):
        self.cubefield.DragToRow = value

    @property
    def EnableMultiplePageItems(self):
        return self.cubefield.EnableMultiplePageItems

    @EnableMultiplePageItems.setter
    def EnableMultiplePageItems(self, value):
        self.cubefield.EnableMultiplePageItems = value

    @property
    def FlattenHierarchies(self):
        return self.cubefield.FlattenHierarchies

    @FlattenHierarchies.setter
    def FlattenHierarchies(self, value):
        self.cubefield.FlattenHierarchies = value

    @property
    def HasMemberProperties(self):
        return self.cubefield.HasMemberProperties

    @property
    def HierarchizeDistinct(self):
        return self.cubefield.HierarchizeDistinct

    @HierarchizeDistinct.setter
    def HierarchizeDistinct(self, value):
        self.cubefield.HierarchizeDistinct = value

    @property
    def IncludeNewItemsInFilter(self):
        return self.cubefield.IncludeNewItemsInFilter

    @IncludeNewItemsInFilter.setter
    def IncludeNewItemsInFilter(self, value):
        self.cubefield.IncludeNewItemsInFilter = value

    @property
    def IsDate(self):
        return self.cubefield.IsDate

    @property
    def LayoutForm(self):
        return XlLayoutFormType(self.cubefield.LayoutForm)

    @LayoutForm.setter
    def LayoutForm(self, value):
        self.cubefield.LayoutForm = value

    @property
    def LayoutSubtotalLocation(self):
        return XlSubtotalLocationType(self.cubefield.LayoutSubtotalLocation)

    @LayoutSubtotalLocation.setter
    def LayoutSubtotalLocation(self, value):
        self.cubefield.LayoutSubtotalLocation = value

    @property
    def Name(self):
        return self.cubefield.Name

    @property
    def Orientation(self):
        return XlPivotFieldOrientation(self.cubefield.Orientation)

    @Orientation.setter
    def Orientation(self, value):
        self.cubefield.Orientation = value

    @property
    def Parent(self):
        return self.cubefield.Parent

    @property
    def PivotFields(self):
        return PivotFields(self.cubefield.PivotFields)

    @property
    def Position(self):
        return self.cubefield.Position

    @Position.setter
    def Position(self, value):
        self.cubefield.Position = value

    @property
    def ShowInFieldList(self):
        return self.cubefield.ShowInFieldList

    @ShowInFieldList.setter
    def ShowInFieldList(self, value):
        self.cubefield.ShowInFieldList = value

    @property
    def TreeviewControl(self):
        return TreeviewControl(self.cubefield.TreeviewControl)

    @property
    def Value(self):
        return self.cubefield.Value

    def AddMemberPropertyField(self, Property=None, PropertyOrder=None, PropertyDisplayedIn=None):
        params = [
            Property if Property is not None else pythoncom.Missing,
            PropertyOrder if PropertyOrder is not None else pythoncom.Missing,
            PropertyDisplayedIn if PropertyDisplayedIn is not None else pythoncom.Missing,
        ]
        self.cubefield.AddMemberPropertyField(*params)

    def ClearManualFilter(self):
        self.cubefield.ClearManualFilter()

    def CreatePivotFields(self):
        self.cubefield.CreatePivotFields()

    def Delete(self):
        self.cubefield.Delete()


class CubeFields:

    def __init__(self, cubefields=None):
        self.cubefields = cubefields

    def __call__(self, item):
        return CubeField(self.cubefields(item))

    @property
    def Application(self):
        return self.cubefields.Application

    @property
    def Count(self):
        return self.cubefields.Count

    @property
    def Creator(self):
        return self.cubefields.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.cubefields.Item):
            return self.cubefields.Item(*params)
        else:
            return self.cubefields.GetItem(*params)

    @property
    def Parent(self):
        return self.cubefields.Parent

    def AddSet(self, Name=None, Caption=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
            Caption if Caption is not None else pythoncom.Missing,
        ]
        return self.cubefields.AddSet(*params)


class CustomProperties:

    def __init__(self, customproperties=None):
        self.customproperties = customproperties

    def __call__(self, item):
        return CustomPropertie(self.customproperties(item))

    @property
    def Application(self):
        return self.customproperties.Application

    @property
    def Count(self):
        return self.customproperties.Count

    @property
    def Creator(self):
        return self.customproperties.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.customproperties.Item):
            return self.customproperties.Item(*params)
        else:
            return self.customproperties.GetItem(*params)

    @property
    def Parent(self):
        return self.customproperties.Parent

    def Add(self, Name=None, Value=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
            Value if Value is not None else pythoncom.Missing,
        ]
        return CustomProperty(self.customproperties.Add(*params))


class CustomProperty:

    def __init__(self, customproperty=None):
        self.customproperty = customproperty

    @property
    def Application(self):
        return self.customproperty.Application

    @property
    def Creator(self):
        return self.customproperty.Creator

    @property
    def Name(self):
        return self.customproperty.Name

    @Name.setter
    def Name(self, value):
        self.customproperty.Name = value

    @property
    def Parent(self):
        return self.customproperty.Parent

    @property
    def Value(self):
        return self.customproperty.Value

    def Delete(self):
        self.customproperty.Delete()


class CustomView:

    def __init__(self, customview=None):
        self.customview = customview

    @property
    def Application(self):
        return self.customview.Application

    @property
    def Creator(self):
        return self.customview.Creator

    @property
    def Name(self):
        return self.customview.Name

    @property
    def Parent(self):
        return self.customview.Parent

    @property
    def PrintSettings(self):
        return self.customview.PrintSettings

    @property
    def RowColSettings(self):
        return self.customview.RowColSettings

    def Delete(self):
        self.customview.Delete()

    def Show(self):
        self.customview.Show()


class CustomViews:

    def __init__(self, customviews=None):
        self.customviews = customviews

    def __call__(self, item):
        return CustomView(self.customviews(item))

    @property
    def Application(self):
        return self.customviews.Application

    @property
    def Count(self):
        return self.customviews.Count

    @property
    def Creator(self):
        return self.customviews.Creator

    @property
    def Parent(self):
        return self.customviews.Parent

    def Add(self, ViewName=None, PrintSettings=None, RowColSettings=None):
        params = [
            ViewName if ViewName is not None else pythoncom.Missing,
            PrintSettings if PrintSettings is not None else pythoncom.Missing,
            RowColSettings if RowColSettings is not None else pythoncom.Missing,
        ]
        return CustomView(self.customviews.Add(*params))

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return CustomView(self.customviews.Item(*params))


class Databar:

    def __init__(self, databar=None):
        self.databar = databar

    @property
    def Application(self):
        return self.databar.Application

    @property
    def AppliesTo(self):
        return Range(self.databar.AppliesTo)

    @property
    def AxisColor(self):
        return self.databar.AxisColor

    @property
    def AxisPosition(self):
        return self.databar.AxisPosition

    @AxisPosition.setter
    def AxisPosition(self, value):
        self.databar.AxisPosition = value

    @property
    def BarBorder(self):
        return self.databar.BarBorder

    @property
    def BarColor(self):
        return FormatColor(self.databar.BarColor)

    @property
    def BarFillType(self):
        return self.databar.BarFillType

    @BarFillType.setter
    def BarFillType(self, value):
        self.databar.BarFillType = value

    @property
    def Creator(self):
        return self.databar.Creator

    @property
    def Direction(self):
        return self.databar.Direction

    @Direction.setter
    def Direction(self, value):
        self.databar.Direction = value

    @property
    def Formula(self):
        return self.databar.Formula

    @Formula.setter
    def Formula(self, value):
        self.databar.Formula = value

    @property
    def MaxPoint(self):
        return ConditionValue(self.databar.MaxPoint)

    @property
    def MinPoint(self):
        return ConditionValue(self.databar.MinPoint)

    @property
    def NegativeBarFormat(self):
        return NegativeBarFormat(self.databar.NegativeBarFormat)

    @property
    def Parent(self):
        return self.databar.Parent

    @property
    def PercentMax(self):
        return self.databar.PercentMax

    @PercentMax.setter
    def PercentMax(self, value):
        self.databar.PercentMax = value

    @property
    def PercentMin(self):
        return self.databar.PercentMin

    @PercentMin.setter
    def PercentMin(self, value):
        self.databar.PercentMin = value

    @property
    def Priority(self):
        return self.databar.Priority

    @Priority.setter
    def Priority(self, value):
        self.databar.Priority = value

    @property
    def PTCondition(self):
        return self.databar.PTCondition

    @property
    def ScopeType(self):
        return XlPivotConditionScope(self.databar.ScopeType)

    @ScopeType.setter
    def ScopeType(self, value):
        self.databar.ScopeType = value

    @property
    def ShowValue(self):
        return self.databar.ShowValue

    @ShowValue.setter
    def ShowValue(self, value):
        self.databar.ShowValue = value

    @property
    def StopIfTrue(self):
        return self.databar.StopIfTrue

    @StopIfTrue.setter
    def StopIfTrue(self, value):
        self.databar.StopIfTrue = value

    @property
    def Type(self):
        return XlFormatConditionType(self.databar.Type)

    def Delete(self):
        self.databar.Delete()

    def ModifyAppliesToRange(self, Range=None):
        params = [
            Range if Range is not None else pythoncom.Missing,
        ]
        self.databar.ModifyAppliesToRange(*params)

    def SetFirstPriority(self):
        self.databar.SetFirstPriority()

    def SetLastPriority(self):
        self.databar.SetLastPriority()


class DataBarBorder:

    def __init__(self, databarborder=None):
        self.databarborder = databarborder

    @property
    def Application(self):
        return self.databarborder.Application

    @property
    def Color(self):
        return self.databarborder.Color

    @property
    def Creator(self):
        return self.databarborder.Creator

    @property
    def Parent(self):
        return self.databarborder.Parent

    @property
    def Type(self):
        return self.databarborder.Type

    @Type.setter
    def Type(self, value):
        self.databarborder.Type = value


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
        params = [
            Start if Start is not None else pythoncom.Missing,
            Length if Length is not None else pythoncom.Missing,
        ]
        if callable(self.datalabel.Characters):
            return Characters(self.datalabel.Characters(*params))
        else:
            return Characters(self.datalabel.GetCharacters(*params))

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
        return XlDataLabelPosition(self.datalabel.Position)

    @Position.setter
    def Position(self, value):
        self.datalabel.Position = value

    @property
    def ReadingOrder(self):
        return self.datalabel.ReadingOrder

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

    def Delete(self):
        return self.datalabel.Delete()

    def Select(self):
        return self.datalabel.Select()


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
        return XlDataLabelPosition(self.datalabels.Position)

    @Position.setter
    def Position(self, value):
        self.datalabels.Position = value

    @property
    def ReadingOrder(self):
        return self.datalabels.ReadingOrder

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
        return self.datalabels.Delete()

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return DataLabel(self.datalabels.Item(*params))

    def Select(self):
        return self.datalabels.Select()


class DataTable:

    def __init__(self, datatable=None):
        self.datatable = datatable

    @property
    def Application(self):
        return self.datatable.Application

    @property
    def Border(self):
        return Border(self.datatable.Border)

    @property
    def Creator(self):
        return self.datatable.Creator

    @property
    def Font(self):
        return Font(self.datatable.Font)

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
        return self.defaultweboptions.Application

    @property
    def CheckIfOfficeIsHTMLEditor(self):
        return self.defaultweboptions.CheckIfOfficeIsHTMLEditor

    @CheckIfOfficeIsHTMLEditor.setter
    def CheckIfOfficeIsHTMLEditor(self, value):
        self.defaultweboptions.CheckIfOfficeIsHTMLEditor = value

    @property
    def Creator(self):
        return self.defaultweboptions.Creator

    @property
    def DownloadComponents(self):
        return self.defaultweboptions.DownloadComponents

    @DownloadComponents.setter
    def DownloadComponents(self, value):
        self.defaultweboptions.DownloadComponents = value

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
    def LoadPictures(self):
        return self.defaultweboptions.LoadPictures

    @LoadPictures.setter
    def LoadPictures(self, value):
        self.defaultweboptions.LoadPictures = value

    @property
    def LocationOfComponents(self):
        return self.defaultweboptions.LocationOfComponents

    @LocationOfComponents.setter
    def LocationOfComponents(self, value):
        self.defaultweboptions.LocationOfComponents = value

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
    def SaveHiddenData(self):
        return self.defaultweboptions.SaveHiddenData

    @SaveHiddenData.setter
    def SaveHiddenData(self, value):
        self.defaultweboptions.SaveHiddenData = value

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
        return self.dialog.Application

    @property
    def Creator(self):
        return self.dialog.Creator

    @property
    def Parent(self):
        return self.dialog.Parent

    def Show(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.dialog.Show(*params)


class Dialogs:

    def __init__(self, dialogs=None):
        self.dialogs = dialogs

    def __call__(self, item):
        return Dialog(self.dialogs(item))

    @property
    def Application(self):
        return self.dialogs.Application

    @property
    def Count(self):
        return self.dialogs.Count

    @property
    def Creator(self):
        return self.dialogs.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.dialogs.Item):
            return self.dialogs.Item(*params)
        else:
            return self.dialogs.GetItem(*params)

    @property
    def Parent(self):
        return self.dialogs.Parent


class DialogSheetView:

    def __init__(self, dialogsheetview=None):
        self.dialogsheetview = dialogsheetview

    @property
    def Application(self):
        return self.dialogsheetview.Application

    @property
    def Creator(self):
        return self.dialogsheetview.Creator

    @property
    def Parent(self):
        return self.dialogsheetview.Parent

    @property
    def Sheet(self):
        return DialogSheetView(self.dialogsheetview.Sheet)


class DisplayFormat:

    def __init__(self, displayformat=None):
        self.displayformat = displayformat

    @property
    def AddIndent(self):
        return Range(self.displayformat.AddIndent)

    @property
    def Application(self):
        return self.displayformat.Application

    @property
    def Borders(self):
        return Borders(self.displayformat.Borders)

    def Characters(self, Start=None, Length=None):
        params = [
            Start if Start is not None else pythoncom.Missing,
            Length if Length is not None else pythoncom.Missing,
        ]
        if callable(self.displayformat.Characters):
            return Characters(self.displayformat.Characters(*params))
        else:
            return Characters(self.displayformat.GetCharacters(*params))

    @property
    def Creator(self):
        return self.displayformat.Creator

    @property
    def Font(self):
        return Font(self.displayformat.Font)

    @property
    def FormulaHidden(self):
        return Range(self.displayformat.FormulaHidden)

    @property
    def HorizontalAlignment(self):
        return Range(self.displayformat.HorizontalAlignment)

    @property
    def IndentLevel(self):
        return Range(self.displayformat.IndentLevel)

    @property
    def Interior(self):
        return Interior(self.displayformat.Interior)

    @property
    def Locked(self):
        return Range(self.displayformat.Locked)

    @property
    def MergeCells(self):
        return Range(self.displayformat.MergeCells)

    @property
    def NumberFormat(self):
        return Range(self.displayformat.NumberFormat)

    @property
    def NumberFormatLocal(self):
        return Range(self.displayformat.NumberFormatLocal)

    @property
    def Orientation(self):
        return Range(self.displayformat.Orientation)

    @property
    def Parent(self):
        return self.displayformat.Parent

    @property
    def ReadingOrder(self):
        return Range(self.displayformat.ReadingOrder)

    @property
    def ShrinkToFit(self):
        return Range(self.displayformat.ShrinkToFit)

    @property
    def Style(self):
        return Style(self.displayformat.Style)

    @property
    def VerticalAlignment(self):
        return Range(self.displayformat.VerticalAlignment)

    @property
    def WrapText(self):
        return Range(self.displayformat.WrapText)


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
        params = [
            Start if Start is not None else pythoncom.Missing,
            Length if Length is not None else pythoncom.Missing,
        ]
        if callable(self.displayunitlabel.Characters):
            return Characters(self.displayunitlabel.Characters(*params))
        else:
            return Characters(self.displayunitlabel.GetCharacters(*params))

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
        return XlChartElementPosition(self.displayunitlabel.Position)

    @Position.setter
    def Position(self, value):
        self.displayunitlabel.Position = value

    @property
    def ReadingOrder(self):
        return self.displayunitlabel.ReadingOrder

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
        return self.displayunitlabel.Delete()

    def Select(self):
        return self.displayunitlabel.Select()


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
        return self.downbars.Delete()

    def Select(self):
        return self.downbars.Select()


class DropLines:

    def __init__(self, droplines=None):
        self.droplines = droplines

    @property
    def Application(self):
        return self.droplines.Application

    @property
    def Border(self):
        return Border(self.droplines.Border)

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
        return self.droplines.Delete()

    def Select(self):
        return self.droplines.Select()


class Error:

    def __init__(self, error=None):
        self.error = error

    @property
    def Application(self):
        return self.error.Application

    @property
    def Creator(self):
        return self.error.Creator

    @property
    def Ignore(self):
        return self.error.Ignore

    @Ignore.setter
    def Ignore(self, value):
        self.error.Ignore = value

    @property
    def Parent(self):
        return self.error.Parent

    @property
    def Value(self):
        return self.error.Value


class ErrorBars:

    def __init__(self, errorbars=None):
        self.errorbars = errorbars

    @property
    def Application(self):
        return self.errorbars.Application

    @property
    def Border(self):
        return Border(self.errorbars.Border)

    @property
    def Creator(self):
        return self.errorbars.Creator

    @property
    def EndStyle(self):
        return XlEndStyleCap(self.errorbars.EndStyle)

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
        return self.errorbars.ClearFormats()

    def Delete(self):
        return self.errorbars.Delete()

    def Select(self):
        return self.errorbars.Select()


class ErrorCheckingOptions:

    def __init__(self, errorcheckingoptions=None):
        self.errorcheckingoptions = errorcheckingoptions

    @property
    def Application(self):
        return self.errorcheckingoptions.Application

    @property
    def BackgroundChecking(self):
        return self.errorcheckingoptions.BackgroundChecking

    @BackgroundChecking.setter
    def BackgroundChecking(self, value):
        self.errorcheckingoptions.BackgroundChecking = value

    @property
    def Creator(self):
        return self.errorcheckingoptions.Creator

    @property
    def EmptyCellReferences(self):
        return self.errorcheckingoptions.EmptyCellReferences

    @EmptyCellReferences.setter
    def EmptyCellReferences(self, value):
        self.errorcheckingoptions.EmptyCellReferences = value

    @property
    def EvaluateToError(self):
        return self.errorcheckingoptions.EvaluateToError

    @EvaluateToError.setter
    def EvaluateToError(self, value):
        self.errorcheckingoptions.EvaluateToError = value

    @property
    def InconsistentFormula(self):
        return self.errorcheckingoptions.InconsistentFormula

    @InconsistentFormula.setter
    def InconsistentFormula(self, value):
        self.errorcheckingoptions.InconsistentFormula = value

    @property
    def InconsistentTableFormula(self):
        return self.errorcheckingoptions.InconsistentTableFormula

    @InconsistentTableFormula.setter
    def InconsistentTableFormula(self, value):
        self.errorcheckingoptions.InconsistentTableFormula = value

    @property
    def IndicatorColorIndex(self):
        return XlColorIndex(self.errorcheckingoptions.IndicatorColorIndex)

    @IndicatorColorIndex.setter
    def IndicatorColorIndex(self, value):
        self.errorcheckingoptions.IndicatorColorIndex = value

    @property
    def ListDataValidation(self):
        return self.errorcheckingoptions.ListDataValidation

    @ListDataValidation.setter
    def ListDataValidation(self, value):
        self.errorcheckingoptions.ListDataValidation = value

    @property
    def NumberAsText(self):
        return self.errorcheckingoptions.NumberAsText

    @NumberAsText.setter
    def NumberAsText(self, value):
        self.errorcheckingoptions.NumberAsText = value

    @property
    def OmittedCells(self):
        return self.errorcheckingoptions.OmittedCells

    @OmittedCells.setter
    def OmittedCells(self, value):
        self.errorcheckingoptions.OmittedCells = value

    @property
    def Parent(self):
        return self.errorcheckingoptions.Parent

    @property
    def TextDate(self):
        return self.errorcheckingoptions.TextDate

    @TextDate.setter
    def TextDate(self, value):
        self.errorcheckingoptions.TextDate = value

    @property
    def UnlockedFormulaCells(self):
        return self.errorcheckingoptions.UnlockedFormulaCells

    @UnlockedFormulaCells.setter
    def UnlockedFormulaCells(self, value):
        self.errorcheckingoptions.UnlockedFormulaCells = value


class Errors:

    def __init__(self, errors=None):
        self.errors = errors

    @property
    def Application(self):
        return self.errors.Application

    @property
    def Creator(self):
        return self.errors.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.errors.Item):
            return Error(self.errors.Item(*params))
        else:
            return Error(self.errors.GetItem(*params))

    @property
    def Parent(self):
        return self.errors.Parent


class FileExportConverter:

    def __init__(self, fileexportconverter=None):
        self.fileexportconverter = fileexportconverter

    @property
    def Application(self):
        return Application(self.fileexportconverter.Application)

    @property
    def Creator(self):
        return self.fileexportconverter.Creator

    @property
    def Description(self):
        return self.fileexportconverter.Description

    @property
    def Extensions(self):
        return FileExportConverter(self.fileexportconverter.Extensions)

    @property
    def FileFormat(self):
        return FileExportConverter(self.fileexportconverter.FileFormat)

    @property
    def Parent(self):
        return self.fileexportconverter.Parent


class FileExportConverters:

    def __init__(self, fileexportconverters=None):
        self.fileexportconverters = fileexportconverters

    def __call__(self, item):
        return FileExportConverter(self.fileexportconverters(item))

    @property
    def Application(self):
        return Application(self.fileexportconverters.Application)

    @property
    def Count(self):
        return self.fileexportconverters.Count

    @property
    def Creator(self):
        return self.fileexportconverters.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.fileexportconverters.Item):
            return FileExportConverter(self.fileexportconverters.Item(*params))
        else:
            return FileExportConverter(self.fileexportconverters.GetItem(*params))

    @property
    def Parent(self):
        return self.fileexportconverters.Parent


class FillFormat:

    def __init__(self, fillformat=None):
        self.fillformat = fillformat

    @property
    def Application(self):
        return self.fillformat.Application

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
        params = [
            Style if Style is not None else pythoncom.Missing,
            Variant if Variant is not None else pythoncom.Missing,
            Degree if Degree is not None else pythoncom.Missing,
        ]
        self.fillformat.OneColorGradient(*params)

    def Patterned(self, Pattern=None):
        params = [
            Pattern if Pattern is not None else pythoncom.Missing,
        ]
        self.fillformat.Patterned(*params)

    def PresetGradient(self, Style=None, Variant=None, PresetGradientType=None):
        params = [
            Style if Style is not None else pythoncom.Missing,
            Variant if Variant is not None else pythoncom.Missing,
            PresetGradientType if PresetGradientType is not None else pythoncom.Missing,
        ]
        self.fillformat.PresetGradient(*params)

    def PresetTextured(self, PresetTexture=None):
        params = [
            PresetTexture if PresetTexture is not None else pythoncom.Missing,
        ]
        self.fillformat.PresetTextured(*params)

    def Solid(self):
        self.fillformat.Solid()

    def TwoColorGradient(self, Style=None, Variant=None):
        params = [
            Style if Style is not None else pythoncom.Missing,
            Variant if Variant is not None else pythoncom.Missing,
        ]
        self.fillformat.TwoColorGradient(*params)

    def UserPicture(self, PictureFile=None):
        params = [
            PictureFile if PictureFile is not None else pythoncom.Missing,
        ]
        self.fillformat.UserPicture(*params)

    def UserTextured(self, TextureFile=None):
        params = [
            TextureFile if TextureFile is not None else pythoncom.Missing,
        ]
        self.fillformat.UserTextured(*params)


class Filter:

    def __init__(self, filter=None):
        self.filter = filter

    @property
    def Application(self):
        return self.filter.Application

    @property
    def Count(self):
        return self.filter.Count

    @property
    def Creator(self):
        return self.filter.Creator

    @property
    def Criteria1(self):
        return self.filter.Criteria1

    @property
    def Criteria2(self):
        return self.filter.Criteria2

    @property
    def On(self):
        return self.filter.On

    @property
    def Operator(self):
        return XlAutoFilterOperator(self.filter.Operator)

    @property
    def Parent(self):
        return self.filter.Parent


class Filters:

    def __init__(self, filters=None):
        self.filters = filters

    def __call__(self, item):
        return Filter(self.filters(item))

    @property
    def Application(self):
        return self.filters.Application

    @property
    def Count(self):
        return self.filters.Count

    @property
    def Creator(self):
        return self.filters.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.filters.Item):
            return self.filters.Item(*params)
        else:
            return self.filters.GetItem(*params)

    @property
    def Parent(self):
        return self.filters.Parent


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
        return self.floor.ClearFormats()

    def Paste(self):
        self.floor.Paste()

    def Select(self):
        return self.floor.Select()


class Font:

    def __init__(self, font=None):
        self.font = font

    @property
    def Application(self):
        return self.font.Application

    @property
    def Background(self):
        return XlBackground(self.font.Background)

    @Background.setter
    def Background(self, value):
        self.font.Background = value

    @property
    def Bold(self):
        return self.font.Bold

    @Bold.setter
    def Bold(self, value):
        self.font.Bold = value

    @property
    def Color(self):
        return RGB(self.font.Color)

    @Color.setter
    def Color(self, value):
        self.font.Color = value

    @property
    def ColorIndex(self):
        return self.font.ColorIndex

    @ColorIndex.setter
    def ColorIndex(self, value):
        self.font.ColorIndex = value

    @property
    def Creator(self):
        return self.font.Creator

    @property
    def FontStyle(self):
        return self.font.FontStyle

    @FontStyle.setter
    def FontStyle(self, value):
        self.font.FontStyle = value

    @property
    def Italic(self):
        return self.font.Italic

    @Italic.setter
    def Italic(self, value):
        self.font.Italic = value

    @property
    def Name(self):
        return self.font.Name

    @Name.setter
    def Name(self, value):
        self.font.Name = value

    @property
    def Parent(self):
        return self.font.Parent

    @property
    def Size(self):
        return self.font.Size

    @Size.setter
    def Size(self, value):
        self.font.Size = value

    @property
    def Strikethrough(self):
        return self.font.Strikethrough

    @Strikethrough.setter
    def Strikethrough(self, value):
        self.font.Strikethrough = value

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
    def ThemeColor(self):
        return self.font.ThemeColor

    @ThemeColor.setter
    def ThemeColor(self, value):
        self.font.ThemeColor = value

    @property
    def ThemeFont(self):
        return XlThemeFont(self.font.ThemeFont)

    @ThemeFont.setter
    def ThemeFont(self, value):
        self.font.ThemeFont = value

    @property
    def TintAndShade(self):
        return self.font.TintAndShade

    @TintAndShade.setter
    def TintAndShade(self, value):
        self.font.TintAndShade = value

    @property
    def Underline(self):
        return self.font.Underline

    @Underline.setter
    def Underline(self, value):
        self.font.Underline = value


class FormatColor:

    def __init__(self, formatcolor=None):
        self.formatcolor = formatcolor

    @property
    def Application(self):
        return self.formatcolor.Application

    @property
    def Color(self):
        return self.formatcolor.Color

    @Color.setter
    def Color(self, value):
        self.formatcolor.Color = value

    @property
    def ColorIndex(self):
        return XlColorIndex(self.formatcolor.ColorIndex)

    @ColorIndex.setter
    def ColorIndex(self, value):
        self.formatcolor.ColorIndex = value

    @property
    def Creator(self):
        return self.formatcolor.Creator

    @property
    def Parent(self):
        return self.formatcolor.Parent

    @property
    def ThemeColor(self):
        return XlThemeColor(self.formatcolor.ThemeColor)

    @ThemeColor.setter
    def ThemeColor(self, value):
        self.formatcolor.ThemeColor = value

    @property
    def TintAndShade(self):
        return self.formatcolor.TintAndShade

    @TintAndShade.setter
    def TintAndShade(self, value):
        self.formatcolor.TintAndShade = value


class FormatCondition:

    def __init__(self, formatcondition=None):
        self.formatcondition = formatcondition

    @property
    def Application(self):
        return self.formatcondition.Application

    @property
    def AppliesTo(self):
        return Range(self.formatcondition.AppliesTo)

    @property
    def Borders(self):
        return Borders(self.formatcondition.Borders)

    @property
    def Creator(self):
        return self.formatcondition.Creator

    @property
    def DateOperator(self):
        return self.formatcondition.DateOperator

    @DateOperator.setter
    def DateOperator(self, value):
        self.formatcondition.DateOperator = value

    @property
    def Font(self):
        return Font(self.formatcondition.Font)

    @property
    def Formula1(self):
        return self.formatcondition.Formula1

    @property
    def Formula2(self):
        return self.formatcondition.Formula2

    @property
    def Interior(self):
        return Interior(self.formatcondition.Interior)

    @property
    def NumberFormat(self):
        return self.formatcondition.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value):
        self.formatcondition.NumberFormat = value

    @property
    def Operator(self):
        return self.formatcondition.Operator

    @property
    def Parent(self):
        return self.formatcondition.Parent

    @property
    def Priority(self):
        return self.formatcondition.Priority

    @Priority.setter
    def Priority(self, value):
        self.formatcondition.Priority = value

    @property
    def PTCondition(self):
        return self.formatcondition.PTCondition

    @property
    def ScopeType(self):
        return XlPivotConditionScope(self.formatcondition.ScopeType)

    @ScopeType.setter
    def ScopeType(self, value):
        self.formatcondition.ScopeType = value

    @property
    def StopIfTrue(self):
        return self.formatcondition.StopIfTrue

    @StopIfTrue.setter
    def StopIfTrue(self, value):
        self.formatcondition.StopIfTrue = value

    @property
    def Text(self):
        return self.formatcondition.Text

    @Text.setter
    def Text(self, value):
        self.formatcondition.Text = value

    @property
    def TextOperator(self):
        return XlContainsOperator(self.formatcondition.TextOperator)

    @TextOperator.setter
    def TextOperator(self, value):
        self.formatcondition.TextOperator = value

    @property
    def Type(self):
        return XlFormatConditionType(self.formatcondition.Type)

    def Delete(self):
        self.formatcondition.Delete()

    def Modify(self, Type=None, Operator=None, Formula1=None, Formula2=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            Operator if Operator is not None else pythoncom.Missing,
            Formula1 if Formula1 is not None else pythoncom.Missing,
            Formula2 if Formula2 is not None else pythoncom.Missing,
        ]
        self.formatcondition.Modify(*params)

    def ModifyAppliesToRange(self, Range=None):
        params = [
            Range if Range is not None else pythoncom.Missing,
        ]
        self.formatcondition.ModifyAppliesToRange(*params)

    def SetFirstPriority(self):
        self.formatcondition.SetFirstPriority()

    def SetLastPriority(self):
        self.formatcondition.SetLastPriority()


class FormatConditions:

    def __init__(self, formatconditions=None):
        self.formatconditions = formatconditions

    @property
    def Application(self):
        return self.formatconditions.Application

    @property
    def Count(self):
        return self.formatconditions.Count

    @property
    def Creator(self):
        return self.formatconditions.Creator

    @property
    def Parent(self):
        return self.formatconditions.Parent

    def Add(self, Type=None, Operator=None, Formula1=None, Formula2=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            Operator if Operator is not None else pythoncom.Missing,
            Formula1 if Formula1 is not None else pythoncom.Missing,
            Formula2 if Formula2 is not None else pythoncom.Missing,
        ]
        return FormatCondition(self.formatconditions.Add(*params))

    def AddAboveAverage(self):
        return self.formatconditions.AddAboveAverage()

    def AddColorScale(self, ColorScaleType=None):
        params = [
            ColorScaleType if ColorScaleType is not None else pythoncom.Missing,
        ]
        return self.formatconditions.AddColorScale(*params)

    def AddDatabar(self):
        return self.formatconditions.AddDatabar()

    def AddIconSetCondition(self):
        return self.formatconditions.AddIconSetCondition()

    def AddTop10(self):
        return self.formatconditions.AddTop10()

    def AddUniqueValues(self):
        return self.formatconditions.AddUniqueValues()

    def Delete(self):
        self.formatconditions.Delete()

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.formatconditions.Item(*params)


class FreeformBuilder:

    def __init__(self, freeformbuilder=None):
        self.freeformbuilder = freeformbuilder

    @property
    def Application(self):
        return self.freeformbuilder.Application

    @property
    def Creator(self):
        return self.freeformbuilder.Creator

    @property
    def Parent(self):
        return self.freeformbuilder.Parent

    def AddNodes(self, SegmentType=None, EditingType=None, X1=None, Y1=None, X2=None, Y2=None, X3=None, Y3=None):
        params = [
            SegmentType if SegmentType is not None else pythoncom.Missing,
            EditingType if EditingType is not None else pythoncom.Missing,
            X1 if X1 is not None else pythoncom.Missing,
            Y1 if Y1 is not None else pythoncom.Missing,
            X2 if X2 is not None else pythoncom.Missing,
            Y2 if Y2 is not None else pythoncom.Missing,
            X3 if X3 is not None else pythoncom.Missing,
            Y3 if Y3 is not None else pythoncom.Missing,
        ]
        self.freeformbuilder.AddNodes(*params)

    def ConvertToShape(self):
        return self.freeformbuilder.ConvertToShape()


class Graphic:

    def __init__(self, graphic=None):
        self.graphic = graphic

    @property
    def Application(self):
        return self.graphic.Application

    @property
    def Brightness(self):
        return self.graphic.Brightness

    @Brightness.setter
    def Brightness(self, value):
        self.graphic.Brightness = value

    @property
    def ColorType(self):
        return self.graphic.ColorType

    @ColorType.setter
    def ColorType(self, value):
        self.graphic.ColorType = value

    @property
    def Contrast(self):
        return self.graphic.Contrast

    @Contrast.setter
    def Contrast(self, value):
        self.graphic.Contrast = value

    @property
    def Creator(self):
        return self.graphic.Creator

    @property
    def CropBottom(self):
        return self.graphic.CropBottom

    @CropBottom.setter
    def CropBottom(self, value):
        self.graphic.CropBottom = value

    @property
    def CropLeft(self):
        return self.graphic.CropLeft

    @CropLeft.setter
    def CropLeft(self, value):
        self.graphic.CropLeft = value

    @property
    def CropRight(self):
        return self.graphic.CropRight

    @CropRight.setter
    def CropRight(self, value):
        self.graphic.CropRight = value

    @property
    def CropTop(self):
        return self.graphic.CropTop

    @CropTop.setter
    def CropTop(self, value):
        self.graphic.CropTop = value

    @property
    def Filename(self):
        return self.graphic.Filename

    @Filename.setter
    def Filename(self, value):
        self.graphic.Filename = value

    @property
    def Height(self):
        return self.graphic.Height

    @Height.setter
    def Height(self, value):
        self.graphic.Height = value

    @property
    def LockAspectRatio(self):
        return self.graphic.LockAspectRatio

    @LockAspectRatio.setter
    def LockAspectRatio(self, value):
        self.graphic.LockAspectRatio = value

    @property
    def Parent(self):
        return self.graphic.Parent

    @property
    def Width(self):
        return self.graphic.Width

    @Width.setter
    def Width(self, value):
        self.graphic.Width = value


class Gridlines:

    def __init__(self, gridlines=None):
        self.gridlines = gridlines

    @property
    def Application(self):
        return self.gridlines.Application

    @property
    def Border(self):
        return Border(self.gridlines.Border)

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
        return self.gridlines.Delete()

    def Select(self):
        return self.gridlines.Select()


class GroupShapes:

    def __init__(self, groupshapes=None):
        self.groupshapes = groupshapes

    @property
    def Application(self):
        return self.groupshapes.Application

    @property
    def Count(self):
        return self.groupshapes.Count

    @property
    def Creator(self):
        return self.groupshapes.Creator

    @property
    def Parent(self):
        return self.groupshapes.Parent

    def Range(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.groupshapes.Range):
            return ShapeRange(self.groupshapes.Range(*params))
        else:
            return ShapeRange(self.groupshapes.GetRange(*params))

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return Shape(self.groupshapes.Item(*params))


class HeaderFooter:

    def __init__(self, headerfooter=None):
        self.headerfooter = headerfooter

    @property
    def Picture(self):
        return self.headerfooter.Picture

    @property
    def Text(self):
        return Text(self.headerfooter.Text)

    @Text.setter
    def Text(self, value):
        self.headerfooter.Text = value


class HiLoLines:

    def __init__(self, hilolines=None):
        self.hilolines = hilolines

    @property
    def Application(self):
        return self.hilolines.Application

    @property
    def Border(self):
        return Border(self.hilolines.Border)

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
        return self.hilolines.Delete()

    def Select(self):
        return self.hilolines.Select()


class HPageBreak:

    def __init__(self, hpagebreak=None):
        self.hpagebreak = hpagebreak

    @property
    def Application(self):
        return self.hpagebreak.Application

    @property
    def Creator(self):
        return self.hpagebreak.Creator

    @property
    def Extent(self):
        return XlPageBreakExtent(self.hpagebreak.Extent)

    @property
    def Location(self):
        return Range(self.hpagebreak.Location)

    @Location.setter
    def Location(self, value):
        self.hpagebreak.Location = value

    @property
    def Parent(self):
        return self.hpagebreak.Parent

    @property
    def Type(self):
        return XlPageBreak(self.hpagebreak.Type)

    @Type.setter
    def Type(self, value):
        self.hpagebreak.Type = value

    def Delete(self):
        self.hpagebreak.Delete()

    def DragOff(self, Direction=None, RegionIndex=None):
        params = [
            Direction if Direction is not None else pythoncom.Missing,
            RegionIndex if RegionIndex is not None else pythoncom.Missing,
        ]
        self.hpagebreak.DragOff(*params)


class HPageBreaks:

    def __init__(self, hpagebreaks=None):
        self.hpagebreaks = hpagebreaks

    @property
    def Application(self):
        return self.hpagebreaks.Application

    @property
    def Count(self):
        return self.hpagebreaks.Count

    @property
    def Creator(self):
        return self.hpagebreaks.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.hpagebreaks.Item):
            return self.hpagebreaks.Item(*params)
        else:
            return self.hpagebreaks.GetItem(*params)

    @property
    def Parent(self):
        return self.hpagebreaks.Parent

    def Add(self, Before=None):
        params = [
            Before if Before is not None else pythoncom.Missing,
        ]
        return HPageBreak(self.hpagebreaks.Add(*params))


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
        return self.hyperlink.Application

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
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            EditNow if EditNow is not None else pythoncom.Missing,
            Overwrite if Overwrite is not None else pythoncom.Missing,
        ]
        self.hyperlink.CreateNewDocument(*params)

    def Delete(self):
        self.hyperlink.Delete()

    def Follow(self, NewWindow=None, AddHistory=None, ExtraInfo=None, Method=None, HeaderInfo=None):
        params = [
            NewWindow if NewWindow is not None else pythoncom.Missing,
            AddHistory if AddHistory is not None else pythoncom.Missing,
            ExtraInfo if ExtraInfo is not None else pythoncom.Missing,
            Method if Method is not None else pythoncom.Missing,
            HeaderInfo if HeaderInfo is not None else pythoncom.Missing,
        ]
        self.hyperlink.Follow(*params)


class Hyperlinks:

    def __init__(self, hyperlinks=None):
        self.hyperlinks = hyperlinks

    @property
    def Application(self):
        return self.hyperlinks.Application

    @property
    def Count(self):
        return self.hyperlinks.Count

    @property
    def Creator(self):
        return self.hyperlinks.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.hyperlinks.Item):
            return self.hyperlinks.Item(*params)
        else:
            return self.hyperlinks.GetItem(*params)

    @property
    def Parent(self):
        return self.hyperlinks.Parent

    def Add(self, Anchor=None, Address=None, SubAddress=None, ScreenTip=None, TextToDisplay=None):
        params = [
            Anchor if Anchor is not None else pythoncom.Missing,
            Address if Address is not None else pythoncom.Missing,
            SubAddress if SubAddress is not None else pythoncom.Missing,
            ScreenTip if ScreenTip is not None else pythoncom.Missing,
            TextToDisplay if TextToDisplay is not None else pythoncom.Missing,
        ]
        return Hyperlink(self.hyperlinks.Add(*params))

    def Delete(self):
        self.hyperlinks.Delete()


class Icon:

    def __init__(self, icon=None):
        self.icon = icon

    @property
    def Application(self):
        return self.icon.Application

    @property
    def Creator(self):
        return self.icon.Creator

    @property
    def Index(self):
        return IconSet(self.icon.Index)

    @property
    def Parent(self):
        return self.icon.Parent


class IconCriteria:

    def __init__(self, iconcriteria=None):
        self.iconcriteria = iconcriteria

    @property
    def Count(self):
        return self.iconcriteria.Count

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.iconcriteria.Item):
            return IconCriterion(self.iconcriteria.Item(*params))
        else:
            return IconCriterion(self.iconcriteria.GetItem(*params))


class IconCriterion:

    def __init__(self, iconcriterion=None):
        self.iconcriterion = iconcriterion

    @property
    def Icon(self):
        return self.iconcriterion.Icon

    @Icon.setter
    def Icon(self, value):
        self.iconcriterion.Icon = value

    @property
    def Index(self):
        return self.iconcriterion.Index

    @property
    def Operator(self):
        return XlFormatConditionOperator(self.iconcriterion.Operator)

    @Operator.setter
    def Operator(self, value):
        self.iconcriterion.Operator = value

    @property
    def Type(self):
        return XlConditionValueTypes(self.iconcriterion.Type)

    @property
    def Value(self):
        return self.iconcriterion.Value

    @Value.setter
    def Value(self, value):
        self.iconcriterion.Value = value


class IconSet:

    def __init__(self, iconset=None):
        self.iconset = iconset

    @property
    def Application(self):
        return self.iconset.Application

    @property
    def Count(self):
        return self.iconset.Count

    @property
    def Creator(self):
        return self.iconset.Creator

    @property
    def ID(self):
        return XlIconSet(self.iconset.ID)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.iconset.Item):
            return Icon(self.iconset.Item(*params))
        else:
            return Icon(self.iconset.GetItem(*params))

    @property
    def Parent(self):
        return self.iconset.Parent


class IconSetCondition:

    def __init__(self, iconsetcondition=None):
        self.iconsetcondition = iconsetcondition

    @property
    def Application(self):
        return self.iconsetcondition.Application

    @property
    def AppliesTo(self):
        return Range(self.iconsetcondition.AppliesTo)

    @property
    def Creator(self):
        return self.iconsetcondition.Creator

    @property
    def Formula(self):
        return self.iconsetcondition.Formula

    @Formula.setter
    def Formula(self, value):
        self.iconsetcondition.Formula = value

    @property
    def IconCriteria(self):
        return IconCriteria(self.iconsetcondition.IconCriteria)

    @property
    def IconSet(self):
        return IconSets(self.iconsetcondition.IconSet)

    @IconSet.setter
    def IconSet(self, value):
        self.iconsetcondition.IconSet = value

    @property
    def Parent(self):
        return self.iconsetcondition.Parent

    @property
    def PercentileValues(self):
        return self.iconsetcondition.PercentileValues

    @PercentileValues.setter
    def PercentileValues(self, value):
        self.iconsetcondition.PercentileValues = value

    @property
    def Priority(self):
        return self.iconsetcondition.Priority

    @Priority.setter
    def Priority(self, value):
        self.iconsetcondition.Priority = value

    @property
    def PTCondition(self):
        return self.iconsetcondition.PTCondition

    @property
    def ReverseOrder(self):
        return self.iconsetcondition.ReverseOrder

    @ReverseOrder.setter
    def ReverseOrder(self, value):
        self.iconsetcondition.ReverseOrder = value

    @property
    def ScopeType(self):
        return XlPivotConditionScope(self.iconsetcondition.ScopeType)

    @ScopeType.setter
    def ScopeType(self, value):
        self.iconsetcondition.ScopeType = value

    @property
    def ShowIconOnly(self):
        return self.iconsetcondition.ShowIconOnly

    @ShowIconOnly.setter
    def ShowIconOnly(self, value):
        self.iconsetcondition.ShowIconOnly = value

    @property
    def StopIfTrue(self):
        return self.iconsetcondition.StopIfTrue

    @StopIfTrue.setter
    def StopIfTrue(self, value):
        self.iconsetcondition.StopIfTrue = value

    @property
    def Type(self):
        return XlFormatConditionType(self.iconsetcondition.Type)

    def Delete(self):
        self.iconsetcondition.Delete()

    def ModifyAppliesToRange(self, Range=None):
        params = [
            Range if Range is not None else pythoncom.Missing,
        ]
        self.iconsetcondition.ModifyAppliesToRange(*params)

    def SetFirstPriority(self):
        self.iconsetcondition.SetFirstPriority()

    def SetLastPriority(self):
        self.iconsetcondition.SetLastPriority()


class IconSets:

    def __init__(self, iconsets=None):
        self.iconsets = iconsets

    @property
    def Application(self):
        return self.iconsets.Application

    @property
    def Count(self):
        return self.iconsets.Count

    @property
    def Creator(self):
        return self.iconsets.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.iconsets.Item):
            return IconSet(self.iconsets.Item(*params))
        else:
            return IconSet(self.iconsets.GetItem(*params))

    @property
    def Parent(self):
        return self.iconsets.Parent


class Interior:

    def __init__(self, interior=None):
        self.interior = interior

    @property
    def Application(self):
        return self.interior.Application

    @property
    def Color(self):
        return RGB(self.interior.Color)

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
    def Gradient(self):
        return self.interior.Gradient

    @Gradient.setter
    def Gradient(self, value):
        self.interior.Gradient = value

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
        return XlPattern(self.interior.Pattern)

    @Pattern.setter
    def Pattern(self, value):
        self.interior.Pattern = value

    @property
    def PatternColor(self):
        return RGB(self.interior.PatternColor)

    @PatternColor.setter
    def PatternColor(self, value):
        self.interior.PatternColor = value

    @property
    def PatternColorIndex(self):
        return XlColorIndex(self.interior.PatternColorIndex)

    @PatternColorIndex.setter
    def PatternColorIndex(self, value):
        self.interior.PatternColorIndex = value

    @property
    def PatternThemeColor(self):
        return Interior(self.interior.PatternThemeColor)

    @PatternThemeColor.setter
    def PatternThemeColor(self, value):
        self.interior.PatternThemeColor = value

    @property
    def PatternTintAndShade(self):
        return Interior(self.interior.PatternTintAndShade)

    @PatternTintAndShade.setter
    def PatternTintAndShade(self, value):
        self.interior.PatternTintAndShade = value

    @property
    def ThemeColor(self):
        return XlThemeColor(self.interior.ThemeColor)

    @ThemeColor.setter
    def ThemeColor(self, value):
        self.interior.ThemeColor = value

    @property
    def TintAndShade(self):
        return self.interior.TintAndShade

    @TintAndShade.setter
    def TintAndShade(self, value):
        self.interior.TintAndShade = value


class IRtdServer:

    def __init__(self, irtdserver=None):
        self.irtdserver = irtdserver

    def ConnectData(self, TopicID=None, Strings=None, GetNewValues=None):
        params = [
            TopicID if TopicID is not None else pythoncom.Missing,
            Strings if Strings is not None else pythoncom.Missing,
            GetNewValues if GetNewValues is not None else pythoncom.Missing,
        ]
        return self.irtdserver.ConnectData(*params)

    def DisconnectData(self, TopicID=None):
        params = [
            TopicID if TopicID is not None else pythoncom.Missing,
        ]
        self.irtdserver.DisconnectData(*params)

    def Heartbeat(self):
        return self.irtdserver.Heartbeat()

    def RefreshData(self, TopicCount=None):
        params = [
            TopicCount if TopicCount is not None else pythoncom.Missing,
        ]
        return self.irtdserver.RefreshData(*params)

    def ServerStart(self, CallbackObject=None):
        params = [
            CallbackObject if CallbackObject is not None else pythoncom.Missing,
        ]
        return self.irtdserver.ServerStart(*params)

    def ServerTerminate(self):
        self.irtdserver.ServerTerminate()


class IRTDUpdateEvent:

    def __init__(self, irtdupdateevent=None):
        self.irtdupdateevent = irtdupdateevent

    @property
    def HeartbeatInterval(self):
        return self.irtdupdateevent.HeartbeatInterval

    @HeartbeatInterval.setter
    def HeartbeatInterval(self, value):
        self.irtdupdateevent.HeartbeatInterval = value

    def Disconnect(self):
        self.irtdupdateevent.Disconnect()

    def UpdateNotify(self):
        self.irtdupdateevent.UpdateNotify()


class LeaderLines:

    def __init__(self, leaderlines=None):
        self.leaderlines = leaderlines

    @property
    def Application(self):
        return self.leaderlines.Application

    @property
    def Border(self):
        return Border(self.leaderlines.Border)

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

    @Left.setter
    def Left(self, value):
        self.legend.Left = value

    @property
    def Name(self):
        return self.legend.Name

    @property
    def Parent(self):
        return self.legend.Parent

    @property
    def Position(self):
        return XlLegendPosition(self.legend.Position)

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
        return self.legend.Clear()

    def Delete(self):
        return self.legend.Delete()

    def LegendEntries(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.legend.LegendEntries(*params)

    def Select(self):
        return self.legend.Select()


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
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return LegendEntry(self.legendentries.Item(*params))


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
        return Font(self.legendentry.Font)

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
        return self.legendentry.Delete()

    def Select(self):
        return self.legendentry.Select()


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
        return XlColorIndex(self.legendkey.MarkerBackgroundColorIndex)

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
        return XlColorIndex(self.legendkey.MarkerForegroundColorIndex)

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
        return XlMarkerStyle(self.legendkey.MarkerStyle)

    @MarkerStyle.setter
    def MarkerStyle(self, value):
        self.legendkey.MarkerStyle = value

    @property
    def Parent(self):
        return self.legendkey.Parent

    @property
    def PictureType(self):
        return XlChartPictureType(self.legendkey.PictureType)

    @PictureType.setter
    def PictureType(self, value):
        self.legendkey.PictureType = value

    @property
    def PictureUnit2(self):
        return PictureType(self.legendkey.PictureUnit2)

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
        return self.legendkey.ClearFormats()

    def Delete(self):
        return self.legendkey.Delete()


class LinearGradient:

    def __init__(self, lineargradient=None):
        self.lineargradient = lineargradient

    @property
    def Application(self):
        return self.lineargradient.Application

    @property
    def ColorStops(self):
        return ColorStops(self.lineargradient.ColorStops)

    @property
    def Creator(self):
        return self.lineargradient.Creator

    @property
    def Degree(self):
        return self.lineargradient.Degree

    @Degree.setter
    def Degree(self, value):
        self.lineargradient.Degree = value

    @property
    def Parent(self):
        return self.lineargradient.Parent


class LineFormat:

    def __init__(self, lineformat=None):
        self.lineformat = lineformat

    @property
    def Application(self):
        return self.lineformat.Application

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


class LinkFormat:

    def __init__(self, linkformat=None):
        self.linkformat = linkformat

    @property
    def Application(self):
        return self.linkformat.Application

    @property
    def AutoUpdate(self):
        return self.linkformat.AutoUpdate

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

    def Update(self):
        self.linkformat.Update()


class ListColumn:

    def __init__(self, listcolumn=None):
        self.listcolumn = listcolumn

    @property
    def Application(self):
        return self.listcolumn.Application

    @property
    def Creator(self):
        return self.listcolumn.Creator

    @property
    def DataBodyRange(self):
        return Range(self.listcolumn.DataBodyRange)

    @property
    def Index(self):
        return ListColumns(self.listcolumn.Index)

    @property
    def Name(self):
        return self.listcolumn.Name

    @Name.setter
    def Name(self, value):
        self.listcolumn.Name = value

    @property
    def Parent(self):
        return self.listcolumn.Parent

    @property
    def Range(self):
        return Range(self.listcolumn.Range)

    @property
    def Total(self):
        return ListColumn(self.listcolumn.Total)

    @property
    def TotalsCalculation(self):
        return self.listcolumn.TotalsCalculation

    @TotalsCalculation.setter
    def TotalsCalculation(self, value):
        self.listcolumn.TotalsCalculation = value

    @property
    def XPath(self):
        return XPath(self.listcolumn.XPath)

    def Delete(self):
        self.listcolumn.Delete()


class ListColumns:

    def __init__(self, listcolumns=None):
        self.listcolumns = listcolumns

    def __call__(self, item):
        return ListColumn(self.listcolumns(item))

    @property
    def Application(self):
        return self.listcolumns.Application

    @property
    def Count(self):
        return self.listcolumns.Count

    @property
    def Creator(self):
        return self.listcolumns.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.listcolumns.Item):
            return self.listcolumns.Item(*params)
        else:
            return self.listcolumns.GetItem(*params)

    @property
    def Parent(self):
        return self.listcolumns.Parent

    def Add(self, Position=None):
        params = [
            Position if Position is not None else pythoncom.Missing,
        ]
        return ListColumn(self.listcolumns.Add(*params))


class ListDataFormat:

    def __init__(self, listdataformat=None):
        self.listdataformat = listdataformat

    @property
    def AllowFillIn(self):
        return self.listdataformat.AllowFillIn

    @property
    def Application(self):
        return self.listdataformat.Application

    @property
    def Choices(self):
        return self.listdataformat.Choices

    @property
    def Creator(self):
        return self.listdataformat.Creator

    @property
    def DecimalPlaces(self):
        return ListColumn(self.listdataformat.DecimalPlaces)

    @property
    def DefaultValue(self):
        return self.listdataformat.DefaultValue

    @property
    def IsPercent(self):
        return ListColumn(self.listdataformat.IsPercent)

    @property
    def lcid(self):
        return ListColumn(self.listdataformat.lcid)

    @property
    def MaxCharacters(self):
        return ListColumn(self.listdataformat.MaxCharacters)

    @property
    def MaxNumber(self):
        return self.listdataformat.MaxNumber

    @property
    def MinNumber(self):
        return self.listdataformat.MinNumber

    @property
    def Parent(self):
        return self.listdataformat.Parent

    @property
    def ReadOnly(self):
        return self.listdataformat.ReadOnly

    @property
    def Required(self):
        return self.listdataformat.Required

    @property
    def Type(self):
        return XlListDataType(self.listdataformat.Type)


class ListObject:

    def __init__(self, listobject=None):
        self.listobject = listobject

    @property
    def Active(self):
        return self.listobject.Active

    @property
    def AlternativeText(self):
        return self.listobject.AlternativeText

    @AlternativeText.setter
    def AlternativeText(self, value):
        self.listobject.AlternativeText = value

    @property
    def Application(self):
        return self.listobject.Application

    @property
    def AutoFilter(self):
        return self.listobject.AutoFilter

    @property
    def Comment(self):
        return self.listobject.Comment

    @Comment.setter
    def Comment(self, value):
        self.listobject.Comment = value

    @property
    def Creator(self):
        return self.listobject.Creator

    @property
    def DataBodyRange(self):
        return Range(self.listobject.DataBodyRange)

    @property
    def DisplayName(self):
        return ListObject(self.listobject.DisplayName)

    @DisplayName.setter
    def DisplayName(self, value):
        self.listobject.DisplayName = value

    @property
    def DisplayRightToLeft(self):
        return self.listobject.DisplayRightToLeft

    @property
    def HeaderRowRange(self):
        return Range(self.listobject.HeaderRowRange)

    @property
    def InsertRowRange(self):
        return Range(self.listobject.InsertRowRange)

    @property
    def ListColumns(self):
        return ListColumns(self.listobject.ListColumns)

    @property
    def ListRows(self):
        return ListRows(self.listobject.ListRows)

    @property
    def Name(self):
        return self.listobject.Name

    @Name.setter
    def Name(self, value):
        self.listobject.Name = value

    @property
    def Parent(self):
        return self.listobject.Parent

    @property
    def QueryTable(self):
        return QueryTable(self.listobject.QueryTable)

    @property
    def Range(self):
        return Range(self.listobject.Range)

    @property
    def SharePointURL(self):
        return self.listobject.SharePointURL

    @property
    def ShowAutoFilter(self):
        return self.listobject.ShowAutoFilter

    @ShowAutoFilter.setter
    def ShowAutoFilter(self, value):
        self.listobject.ShowAutoFilter = value

    @property
    def ShowHeaders(self):
        return ListObject(self.listobject.ShowHeaders)

    @ShowHeaders.setter
    def ShowHeaders(self, value):
        self.listobject.ShowHeaders = value

    @property
    def ShowTableStyleColumnStripes(self):
        return ListObject(self.listobject.ShowTableStyleColumnStripes)

    @ShowTableStyleColumnStripes.setter
    def ShowTableStyleColumnStripes(self, value):
        self.listobject.ShowTableStyleColumnStripes = value

    @property
    def ShowTableStyleFirstColumn(self):
        return ListObject(self.listobject.ShowTableStyleFirstColumn)

    @ShowTableStyleFirstColumn.setter
    def ShowTableStyleFirstColumn(self, value):
        self.listobject.ShowTableStyleFirstColumn = value

    @property
    def ShowTableStyleLastColumn(self):
        return ListObject(self.listobject.ShowTableStyleLastColumn)

    @ShowTableStyleLastColumn.setter
    def ShowTableStyleLastColumn(self, value):
        self.listobject.ShowTableStyleLastColumn = value

    @property
    def ShowTableStyleRowStripes(self):
        return ListObject(self.listobject.ShowTableStyleRowStripes)

    @ShowTableStyleRowStripes.setter
    def ShowTableStyleRowStripes(self, value):
        self.listobject.ShowTableStyleRowStripes = value

    @property
    def ShowTotals(self):
        return self.listobject.ShowTotals

    @ShowTotals.setter
    def ShowTotals(self, value):
        self.listobject.ShowTotals = value

    @property
    def Sort(self):
        return self.listobject.Sort

    @property
    def SourceType(self):
        return XlListObjectSourceType(self.listobject.SourceType)

    @property
    def Summary(self):
        return self.listobject.Summary

    @Summary.setter
    def Summary(self, value):
        self.listobject.Summary = value

    @property
    def TableStyle(self):
        return self.listobject.TableStyle

    @TableStyle.setter
    def TableStyle(self, value):
        self.listobject.TableStyle = value

    @property
    def TotalsRowRange(self):
        return Range(self.listobject.TotalsRowRange)

    @property
    def XmlMap(self):
        return XmlMap(self.listobject.XmlMap)

    def Delete(self):
        self.listobject.Delete()

    def ExportToVisio(self):
        self.listobject.ExportToVisio()

    def Publish(self, Target=None, LinkSource=None):
        params = [
            Target if Target is not None else pythoncom.Missing,
            LinkSource if LinkSource is not None else pythoncom.Missing,
        ]
        return self.listobject.Publish(*params)

    def Refresh(self):
        self.listobject.Refresh()

    def Resize(self, Range=None):
        params = [
            Range if Range is not None else pythoncom.Missing,
        ]
        self.listobject.Resize(*params)

    def Unlink(self):
        self.listobject.Unlink()

    def Unlist(self):
        self.listobject.Unlist()


class ListObjects:

    def __init__(self, listobjects=None):
        self.listobjects = listobjects

    def __call__(self, item):
        return ListObject(self.listobjects(item))

    @property
    def Application(self):
        return self.listobjects.Application

    @property
    def Count(self):
        return self.listobjects.Count

    @property
    def Creator(self):
        return self.listobjects.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.listobjects.Item):
            return self.listobjects.Item(*params)
        else:
            return self.listobjects.GetItem(*params)

    @property
    def Parent(self):
        return self.listobjects.Parent

    def Add(self, SourceType=None, Source=None, LinkSource=None, XlListObjectHasHeaders=None, Destination=None, TableStyleName=None):
        params = [
            SourceType if SourceType is not None else pythoncom.Missing,
            Source if Source is not None else pythoncom.Missing,
            LinkSource if LinkSource is not None else pythoncom.Missing,
            XlListObjectHasHeaders if XlListObjectHasHeaders is not None else pythoncom.Missing,
            Destination if Destination is not None else pythoncom.Missing,
            TableStyleName if TableStyleName is not None else pythoncom.Missing,
        ]
        return ListObject(self.listobjects.Add(*params))


class ListRow:

    def __init__(self, listrow=None):
        self.listrow = listrow

    @property
    def Application(self):
        return self.listrow.Application

    @property
    def Creator(self):
        return self.listrow.Creator

    @property
    def Index(self):
        return ListRows(self.listrow.Index)

    @property
    def Parent(self):
        return self.listrow.Parent

    @property
    def Range(self):
        return Range(self.listrow.Range)

    def Delete(self):
        self.listrow.Delete()


class ListRows:

    def __init__(self, listrows=None):
        self.listrows = listrows

    def __call__(self, item):
        return ListRow(self.listrows(item))

    @property
    def Application(self):
        return self.listrows.Application

    @property
    def Count(self):
        return self.listrows.Count

    @property
    def Creator(self):
        return self.listrows.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.listrows.Item):
            return self.listrows.Item(*params)
        else:
            return self.listrows.GetItem(*params)

    @property
    def Parent(self):
        return self.listrows.Parent

    def Add(self, Position=None, AlwaysInsert=None):
        params = [
            Position if Position is not None else pythoncom.Missing,
            AlwaysInsert if AlwaysInsert is not None else pythoncom.Missing,
        ]
        return ListRow(self.listrows.Add(*params))


class Mailer:

    def __init__(self, mailer=None):
        self.mailer = mailer

    @property
    def Application(self):
        return self.mailer.Application

    @property
    def BCCRecipients(self):
        return self.mailer.BCCRecipients

    @property
    def CCRecipients(self):
        return self.mailer.CCRecipients

    @property
    def Creator(self):
        return self.mailer.Creator

    @property
    def Enclosures(self):
        return self.mailer.Enclosures

    @property
    def Parent(self):
        return self.mailer.Parent

    @property
    def Received(self):
        return self.mailer.Received

    @property
    def SendDateTime(self):
        return self.mailer.SendDateTime

    @property
    def Sender(self):
        return self.mailer.Sender

    @property
    def Subject(self):
        return self.mailer.Subject

    @Subject.setter
    def Subject(self, value):
        self.mailer.Subject = value

    @property
    def ToRecipients(self):
        return self.mailer.ToRecipients

    @property
    def WhichAddress(self):
        return self.mailer.WhichAddress


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
        params = [
            Increment if Increment is not None else pythoncom.Missing,
        ]
        self.model3dformat.IncrementRotationX(*params)

    def IncrementRotationY(self, Increment=None):
        params = [
            Increment if Increment is not None else pythoncom.Missing,
        ]
        self.model3dformat.IncrementRotationY(*params)

    def IncrementRotationZ(self, Increment=None):
        params = [
            Increment if Increment is not None else pythoncom.Missing,
        ]
        self.model3dformat.IncrementRotationZ(*params)

    def ResetModel(self, ResetSize=None):
        params = [
            ResetSize if ResetSize is not None else pythoncom.Missing,
        ]
        self.model3dformat.ResetModel(*params)


class ModuleView:

    def __init__(self, moduleview=None):
        self.moduleview = moduleview

    @property
    def Application(self):
        return self.moduleview.Application

    @property
    def Creator(self):
        return self.moduleview.Creator

    @property
    def Parent(self):
        return self.moduleview.Parent

    @property
    def Sheet(self):
        return self.moduleview.Sheet


class MultiThreadedCalculation:

    def __init__(self, multithreadedcalculation=None):
        self.multithreadedcalculation = multithreadedcalculation

    @property
    def Application(self):
        return self.multithreadedcalculation.Application

    @property
    def Creator(self):
        return self.multithreadedcalculation.Creator

    @property
    def Enabled(self):
        return self.multithreadedcalculation.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.multithreadedcalculation.Enabled = value

    @property
    def Parent(self):
        return self.multithreadedcalculation.Parent

    @property
    def ThreadCount(self):
        return self.multithreadedcalculation.ThreadCount

    @property
    def ThreadMode(self):
        return XlThreadMode(self.multithreadedcalculation.ThreadMode)

    @ThreadMode.setter
    def ThreadMode(self, value):
        self.multithreadedcalculation.ThreadMode = value


class Names:

    def __init__(self, names=None):
        self.names = names

    def __call__(self, item):
        return Name(self.names(item))

    @property
    def Application(self):
        return self.names.Application

    @property
    def Count(self):
        return self.names.Count

    @property
    def Creator(self):
        return self.names.Creator

    @property
    def Parent(self):
        return self.names.Parent

    def Add(self, Name=None, RefersTo=None, Visible=None, MacroType=None, ShortcutKey=None, Category=None, NameLocal=None, RefersToLocal=None, CategoryLocal=None, RefersToR1C1=None, RefersToR1C1Local=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
            RefersTo if RefersTo is not None else pythoncom.Missing,
            Visible if Visible is not None else pythoncom.Missing,
            MacroType if MacroType is not None else pythoncom.Missing,
            ShortcutKey if ShortcutKey is not None else pythoncom.Missing,
            Category if Category is not None else pythoncom.Missing,
            NameLocal if NameLocal is not None else pythoncom.Missing,
            RefersToLocal if RefersToLocal is not None else pythoncom.Missing,
            CategoryLocal if CategoryLocal is not None else pythoncom.Missing,
            RefersToR1C1 if RefersToR1C1 is not None else pythoncom.Missing,
            RefersToR1C1Local if RefersToR1C1Local is not None else pythoncom.Missing,
        ]
        return Name(self.names.Add(*params))

    def Item(self, Index=None, IndexLocal=None, RefersTo=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
            IndexLocal if IndexLocal is not None else pythoncom.Missing,
            RefersTo if RefersTo is not None else pythoncom.Missing,
        ]
        return self.names.Item(*params)


class NegativeBarFormat:

    def __init__(self, negativebarformat=None):
        self.negativebarformat = negativebarformat

    @property
    def Application(self):
        return self.negativebarformat.Application

    @property
    def BorderColor(self):
        return FormatColor(self.negativebarformat.BorderColor)

    @property
    def BorderColorType(self):
        return self.negativebarformat.BorderColorType

    @BorderColorType.setter
    def BorderColorType(self, value):
        self.negativebarformat.BorderColorType = value

    @property
    def Color(self):
        return FormatColor(self.negativebarformat.Color)

    @property
    def ColorType(self):
        return self.negativebarformat.ColorType

    @ColorType.setter
    def ColorType(self, value):
        self.negativebarformat.ColorType = value

    @property
    def Creator(self):
        return self.negativebarformat.Creator

    @property
    def Parent(self):
        return self.negativebarformat.Parent


class ODBCConnection:

    def __init__(self, odbcconnection=None):
        self.odbcconnection = odbcconnection

    @property
    def AlwaysUseConnectionFile(self):
        return self.odbcconnection.AlwaysUseConnectionFile

    @AlwaysUseConnectionFile.setter
    def AlwaysUseConnectionFile(self, value):
        self.odbcconnection.AlwaysUseConnectionFile = value

    @property
    def Application(self):
        return self.odbcconnection.Application

    @property
    def BackgroundQuery(self):
        return self.odbcconnection.BackgroundQuery

    @BackgroundQuery.setter
    def BackgroundQuery(self, value):
        self.odbcconnection.BackgroundQuery = value

    @property
    def CommandText(self):
        return self.odbcconnection.CommandText

    @CommandText.setter
    def CommandText(self, value):
        self.odbcconnection.CommandText = value

    @property
    def CommandType(self):
        return XlCmdType(self.odbcconnection.CommandType)

    @CommandType.setter
    def CommandType(self, value):
        self.odbcconnection.CommandType = value

    @property
    def Connection(self):
        return self.odbcconnection.Connection

    @Connection.setter
    def Connection(self, value):
        self.odbcconnection.Connection = value

    @property
    def Creator(self):
        return self.odbcconnection.Creator

    @property
    def EnableRefresh(self):
        return self.odbcconnection.EnableRefresh

    @EnableRefresh.setter
    def EnableRefresh(self, value):
        self.odbcconnection.EnableRefresh = value

    @property
    def Parent(self):
        return self.odbcconnection.Parent

    @property
    def RefreshDate(self):
        return self.odbcconnection.RefreshDate

    @property
    def Refreshing(self):
        return self.odbcconnection.Refreshing

    @Refreshing.setter
    def Refreshing(self, value):
        self.odbcconnection.Refreshing = value

    @property
    def RefreshOnFileOpen(self):
        return self.odbcconnection.RefreshOnFileOpen

    @RefreshOnFileOpen.setter
    def RefreshOnFileOpen(self, value):
        self.odbcconnection.RefreshOnFileOpen = value

    @property
    def RefreshPeriod(self):
        return self.odbcconnection.RefreshPeriod

    @RefreshPeriod.setter
    def RefreshPeriod(self, value):
        self.odbcconnection.RefreshPeriod = value

    @property
    def RobustConnect(self):
        return XlRobustConnect(self.odbcconnection.RobustConnect)

    @RobustConnect.setter
    def RobustConnect(self, value):
        self.odbcconnection.RobustConnect = value

    @property
    def SavePassword(self):
        return self.odbcconnection.SavePassword

    @SavePassword.setter
    def SavePassword(self, value):
        self.odbcconnection.SavePassword = value

    @property
    def ServerCredentialsMethod(self):
        return XlCredentialsMethod(self.odbcconnection.ServerCredentialsMethod)

    @ServerCredentialsMethod.setter
    def ServerCredentialsMethod(self, value):
        self.odbcconnection.ServerCredentialsMethod = value

    @property
    def ServerSSOApplicationID(self):
        return self.odbcconnection.ServerSSOApplicationID

    @ServerSSOApplicationID.setter
    def ServerSSOApplicationID(self, value):
        self.odbcconnection.ServerSSOApplicationID = value

    @property
    def SourceConnectionFile(self):
        return self.odbcconnection.SourceConnectionFile

    @SourceConnectionFile.setter
    def SourceConnectionFile(self, value):
        self.odbcconnection.SourceConnectionFile = value

    @property
    def SourceData(self):
        return self.odbcconnection.SourceData

    @SourceData.setter
    def SourceData(self, value):
        self.odbcconnection.SourceData = value

    @property
    def SourceDataFile(self):
        return self.odbcconnection.SourceDataFile

    @SourceDataFile.setter
    def SourceDataFile(self, value):
        self.odbcconnection.SourceDataFile = value

    def CancelRefresh(self):
        self.odbcconnection.CancelRefresh()

    def Refresh(self):
        self.odbcconnection.Refresh()

    def SaveAsODC(self, ODCFileName=None, Description=None, Keywords=None):
        params = [
            ODCFileName if ODCFileName is not None else pythoncom.Missing,
            Description if Description is not None else pythoncom.Missing,
            Keywords if Keywords is not None else pythoncom.Missing,
        ]
        return self.odbcconnection.SaveAsODC(*params)


class ODBCError:

    def __init__(self, odbcerror=None):
        self.odbcerror = odbcerror

    @property
    def Application(self):
        return self.odbcerror.Application

    @property
    def Creator(self):
        return self.odbcerror.Creator

    @property
    def ErrorString(self):
        return self.odbcerror.ErrorString

    @property
    def Parent(self):
        return self.odbcerror.Parent

    @property
    def SqlState(self):
        return self.odbcerror.SqlState


class ODBCErrors:

    def __init__(self, odbcerrors=None):
        self.odbcerrors = odbcerrors

    def __call__(self, item):
        return ODBCError(self.odbcerrors(item))

    @property
    def Application(self):
        return self.odbcerrors.Application

    @property
    def Count(self):
        return self.odbcerrors.Count

    @property
    def Creator(self):
        return self.odbcerrors.Creator

    @property
    def Parent(self):
        return self.odbcerrors.Parent

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return ODBCError(self.odbcerrors.Item(*params))


class OLEDBConnection:

    def __init__(self, oledbconnection=None):
        self.oledbconnection = oledbconnection

    @property
    def ADOConnection(self):
        return self.oledbconnection.ADOConnection

    @property
    def AlwaysUseConnectionFile(self):
        return self.oledbconnection.AlwaysUseConnectionFile

    @AlwaysUseConnectionFile.setter
    def AlwaysUseConnectionFile(self, value):
        self.oledbconnection.AlwaysUseConnectionFile = value

    @property
    def Application(self):
        return self.oledbconnection.Application

    @property
    def BackgroundQuery(self):
        return self.oledbconnection.BackgroundQuery

    @BackgroundQuery.setter
    def BackgroundQuery(self, value):
        self.oledbconnection.BackgroundQuery = value

    @property
    def CalculatedMembers(self):
        return CalculatedMembers(self.oledbconnection.CalculatedMembers)

    @property
    def CommandText(self):
        return self.oledbconnection.CommandText

    @CommandText.setter
    def CommandText(self, value):
        self.oledbconnection.CommandText = value

    @property
    def CommandType(self):
        return XlCmdType(self.oledbconnection.CommandType)

    @CommandType.setter
    def CommandType(self, value):
        self.oledbconnection.CommandType = value

    @property
    def Connection(self):
        return self.oledbconnection.Connection

    @Connection.setter
    def Connection(self, value):
        self.oledbconnection.Connection = value

    @property
    def Creator(self):
        return self.oledbconnection.Creator

    @property
    def EnableRefresh(self):
        return self.oledbconnection.EnableRefresh

    @EnableRefresh.setter
    def EnableRefresh(self, value):
        self.oledbconnection.EnableRefresh = value

    @property
    def IsConnected(self):
        return self.oledbconnection.IsConnected

    @property
    def LocalConnection(self):
        return self.oledbconnection.LocalConnection

    @LocalConnection.setter
    def LocalConnection(self, value):
        self.oledbconnection.LocalConnection = value

    @property
    def LocaleID(self):
        return self.oledbconnection.LocaleID

    @LocaleID.setter
    def LocaleID(self, value):
        self.oledbconnection.LocaleID = value

    @property
    def MaintainConnection(self):
        return self.oledbconnection.MaintainConnection

    @MaintainConnection.setter
    def MaintainConnection(self, value):
        self.oledbconnection.MaintainConnection = value

    @property
    def MaxDrillthroughRecords(self):
        return self.oledbconnection.MaxDrillthroughRecords

    @MaxDrillthroughRecords.setter
    def MaxDrillthroughRecords(self, value):
        self.oledbconnection.MaxDrillthroughRecords = value

    @property
    def OLAP(self):
        return self.oledbconnection.OLAP

    @property
    def Parent(self):
        return self.oledbconnection.Parent

    @property
    def RefreshDate(self):
        return self.oledbconnection.RefreshDate

    @property
    def Refreshing(self):
        return self.oledbconnection.Refreshing

    @Refreshing.setter
    def Refreshing(self, value):
        self.oledbconnection.Refreshing = value

    @property
    def RefreshOnFileOpen(self):
        return self.oledbconnection.RefreshOnFileOpen

    @RefreshOnFileOpen.setter
    def RefreshOnFileOpen(self, value):
        self.oledbconnection.RefreshOnFileOpen = value

    @property
    def RefreshPeriod(self):
        return self.oledbconnection.RefreshPeriod

    @RefreshPeriod.setter
    def RefreshPeriod(self, value):
        self.oledbconnection.RefreshPeriod = value

    @property
    def RetrieveInOfficeUILang(self):
        return self.oledbconnection.RetrieveInOfficeUILang

    @RetrieveInOfficeUILang.setter
    def RetrieveInOfficeUILang(self, value):
        self.oledbconnection.RetrieveInOfficeUILang = value

    @property
    def RobustConnect(self):
        return XlRobustConnect(self.oledbconnection.RobustConnect)

    @RobustConnect.setter
    def RobustConnect(self, value):
        self.oledbconnection.RobustConnect = value

    @property
    def SavePassword(self):
        return self.oledbconnection.SavePassword

    @SavePassword.setter
    def SavePassword(self, value):
        self.oledbconnection.SavePassword = value

    @property
    def ServerCredentialsMethod(self):
        return XlCredentialsMethod(self.oledbconnection.ServerCredentialsMethod)

    @ServerCredentialsMethod.setter
    def ServerCredentialsMethod(self, value):
        self.oledbconnection.ServerCredentialsMethod = value

    @property
    def ServerFillColor(self):
        return self.oledbconnection.ServerFillColor

    @ServerFillColor.setter
    def ServerFillColor(self, value):
        self.oledbconnection.ServerFillColor = value

    @property
    def ServerFontStyle(self):
        return self.oledbconnection.ServerFontStyle

    @ServerFontStyle.setter
    def ServerFontStyle(self, value):
        self.oledbconnection.ServerFontStyle = value

    @property
    def ServerNumberFormat(self):
        return self.oledbconnection.ServerNumberFormat

    @ServerNumberFormat.setter
    def ServerNumberFormat(self, value):
        self.oledbconnection.ServerNumberFormat = value

    @property
    def ServerSSOApplicationID(self):
        return self.oledbconnection.ServerSSOApplicationID

    @ServerSSOApplicationID.setter
    def ServerSSOApplicationID(self, value):
        self.oledbconnection.ServerSSOApplicationID = value

    @property
    def ServerTextColor(self):
        return self.oledbconnection.ServerTextColor

    @ServerTextColor.setter
    def ServerTextColor(self, value):
        self.oledbconnection.ServerTextColor = value

    @property
    def SourceConnectionFile(self):
        return self.oledbconnection.SourceConnectionFile

    @SourceConnectionFile.setter
    def SourceConnectionFile(self, value):
        self.oledbconnection.SourceConnectionFile = value

    @property
    def SourceDataFile(self):
        return self.oledbconnection.SourceDataFile

    @SourceDataFile.setter
    def SourceDataFile(self, value):
        self.oledbconnection.SourceDataFile = value

    @property
    def UseLocalConnection(self):
        return self.oledbconnection.UseLocalConnection

    @UseLocalConnection.setter
    def UseLocalConnection(self, value):
        self.oledbconnection.UseLocalConnection = value

    def CancelRefresh(self):
        self.oledbconnection.CancelRefresh()

    def MakeConnection(self):
        return self.oledbconnection.MakeConnection()

    def Reconnect(self):
        self.oledbconnection.Reconnect()

    def Refresh(self):
        self.oledbconnection.Refresh()

    def SaveAsODC(self, ODCFileName=None, Description=None, Keywords=None):
        params = [
            ODCFileName if ODCFileName is not None else pythoncom.Missing,
            Description if Description is not None else pythoncom.Missing,
            Keywords if Keywords is not None else pythoncom.Missing,
        ]
        self.oledbconnection.SaveAsODC(*params)


class OLEDBError:

    def __init__(self, oledberror=None):
        self.oledberror = oledberror

    @property
    def Application(self):
        return self.oledberror.Application

    @property
    def Creator(self):
        return self.oledberror.Creator

    @property
    def ErrorString(self):
        return self.oledberror.ErrorString

    @property
    def Native(self):
        return self.oledberror.Native

    @property
    def Number(self):
        return self.oledberror.Number

    @property
    def Parent(self):
        return self.oledberror.Parent

    @property
    def SqlState(self):
        return self.oledberror.SqlState

    @property
    def Stage(self):
        return self.oledberror.Stage


class OLEDBErrors:

    def __init__(self, oledberrors=None):
        self.oledberrors = oledberrors

    def __call__(self, item):
        return OLEDBError(self.oledberrors(item))

    @property
    def Application(self):
        return self.oledberrors.Application

    @property
    def Count(self):
        return self.oledberrors.Count

    @property
    def Creator(self):
        return self.oledberrors.Creator

    @property
    def Parent(self):
        return self.oledberrors.Parent

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return OLEDBError(self.oledberrors.Item(*params))


class OLEFormat:

    def __init__(self, oleformat=None):
        self.oleformat = oleformat

    @property
    def Application(self):
        return self.oleformat.Application

    @property
    def Creator(self):
        return self.oleformat.Creator

    @property
    def Object(self):
        return self.oleformat.Object

    @property
    def Parent(self):
        return self.oleformat.Parent

    @property
    def progID(self):
        return self.oleformat.progID

    def Activate(self):
        self.oleformat.Activate()

    def Verb(self, Verb=None):
        params = [
            Verb if Verb is not None else pythoncom.Missing,
        ]
        self.oleformat.Verb(*params)


class OLEObject:

    def __init__(self, oleobject=None):
        self.oleobject = oleobject

    @property
    def Application(self):
        return self.oleobject.Application

    @property
    def AutoLoad(self):
        return self.oleobject.AutoLoad

    @AutoLoad.setter
    def AutoLoad(self, value):
        self.oleobject.AutoLoad = value

    @property
    def AutoUpdate(self):
        return self.oleobject.AutoUpdate

    @property
    def Border(self):
        return Border(self.oleobject.Border)

    @property
    def BottomRightCell(self):
        return Range(self.oleobject.BottomRightCell)

    @property
    def Creator(self):
        return self.oleobject.Creator

    @property
    def Enabled(self):
        return self.oleobject.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.oleobject.Enabled = value

    @property
    def Height(self):
        return self.oleobject.Height

    @Height.setter
    def Height(self, value):
        self.oleobject.Height = value

    @property
    def Index(self):
        return self.oleobject.Index

    @property
    def Interior(self):
        return Interior(self.oleobject.Interior)

    @property
    def Left(self):
        return self.oleobject.Left

    @Left.setter
    def Left(self, value):
        self.oleobject.Left = value

    @property
    def LinkedCell(self):
        return self.oleobject.LinkedCell

    @LinkedCell.setter
    def LinkedCell(self, value):
        self.oleobject.LinkedCell = value

    @property
    def ListFillRange(self):
        return self.oleobject.ListFillRange

    @ListFillRange.setter
    def ListFillRange(self, value):
        self.oleobject.ListFillRange = value

    @property
    def Locked(self):
        return self.oleobject.Locked

    @Locked.setter
    def Locked(self, value):
        self.oleobject.Locked = value

    @property
    def Name(self):
        return self.oleobject.Name

    @Name.setter
    def Name(self, value):
        self.oleobject.Name = value

    @property
    def Object(self):
        return self.oleobject.Object

    @property
    def OLEType(self):
        return XlOLEType(self.oleobject.OLEType)

    @property
    def Parent(self):
        return self.oleobject.Parent

    @property
    def Placement(self):
        return XlPlacement(self.oleobject.Placement)

    @Placement.setter
    def Placement(self, value):
        self.oleobject.Placement = value

    @property
    def PrintObject(self):
        return self.oleobject.PrintObject

    @PrintObject.setter
    def PrintObject(self, value):
        self.oleobject.PrintObject = value

    @property
    def progID(self):
        return self.oleobject.progID

    @property
    def Shadow(self):
        return self.oleobject.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.oleobject.Shadow = value

    @property
    def ShapeRange(self):
        return ShapeRange(self.oleobject.ShapeRange)

    @property
    def SourceName(self):
        return self.oleobject.SourceName

    @SourceName.setter
    def SourceName(self, value):
        self.oleobject.SourceName = value

    @property
    def Top(self):
        return self.oleobject.Top

    @Top.setter
    def Top(self, value):
        self.oleobject.Top = value

    @property
    def TopLeftCell(self):
        return Range(self.oleobject.TopLeftCell)

    @property
    def Visible(self):
        return self.oleobject.Visible

    @Visible.setter
    def Visible(self, value):
        self.oleobject.Visible = value

    @property
    def Width(self):
        return self.oleobject.Width

    @Width.setter
    def Width(self, value):
        self.oleobject.Width = value

    @property
    def ZOrder(self):
        return self.oleobject.ZOrder

    def Activate(self):
        return self.oleobject.Activate()

    def BringToFront(self):
        return self.oleobject.BringToFront()

    def Copy(self):
        return self.oleobject.Copy()

    def CopyPicture(self, Appearance=None, Format=None):
        params = [
            Appearance if Appearance is not None else pythoncom.Missing,
            Format if Format is not None else pythoncom.Missing,
        ]
        return self.oleobject.CopyPicture(*params)

    def Cut(self):
        return self.oleobject.Cut()

    def Delete(self):
        return self.oleobject.Delete()

    def Duplicate(self):
        return self.oleobject.Duplicate()

    def Select(self, Replace=None):
        params = [
            Replace if Replace is not None else pythoncom.Missing,
        ]
        return self.oleobject.Select(*params)

    def SendToBack(self):
        return self.oleobject.SendToBack()

    def Update(self):
        return self.oleobject.Update()

    def Verb(self, Verb=None):
        params = [
            Verb if Verb is not None else pythoncom.Missing,
        ]
        return self.oleobject.Verb(*params)


class OLEObjects:

    def __init__(self, oleobjects=None):
        self.oleobjects = oleobjects

    def __call__(self, item):
        return OLEObject(self.oleobjects(item))

    @property
    def Application(self):
        return self.oleobjects.Application

    @property
    def AutoLoad(self):
        return self.oleobjects.AutoLoad

    @AutoLoad.setter
    def AutoLoad(self, value):
        self.oleobjects.AutoLoad = value

    @property
    def Border(self):
        return Border(self.oleobjects.Border)

    @property
    def Count(self):
        return self.oleobjects.Count

    @property
    def Creator(self):
        return self.oleobjects.Creator

    @property
    def Enabled(self):
        return self.oleobjects.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.oleobjects.Enabled = value

    @property
    def Height(self):
        return self.oleobjects.Height

    @Height.setter
    def Height(self, value):
        self.oleobjects.Height = value

    @property
    def Interior(self):
        return Interior(self.oleobjects.Interior)

    @property
    def Left(self):
        return self.oleobjects.Left

    @Left.setter
    def Left(self, value):
        self.oleobjects.Left = value

    @property
    def Locked(self):
        return self.oleobjects.Locked

    @Locked.setter
    def Locked(self, value):
        self.oleobjects.Locked = value

    @property
    def Parent(self):
        return self.oleobjects.Parent

    @property
    def Placement(self):
        return XlPlacement(self.oleobjects.Placement)

    @Placement.setter
    def Placement(self, value):
        self.oleobjects.Placement = value

    @property
    def PrintObject(self):
        return self.oleobjects.PrintObject

    @PrintObject.setter
    def PrintObject(self, value):
        self.oleobjects.PrintObject = value

    @property
    def Shadow(self):
        return self.oleobjects.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.oleobjects.Shadow = value

    @property
    def ShapeRange(self):
        return ShapeRange(self.oleobjects.ShapeRange)

    @property
    def SourceName(self):
        return self.oleobjects.SourceName

    @SourceName.setter
    def SourceName(self, value):
        self.oleobjects.SourceName = value

    @property
    def Top(self):
        return self.oleobjects.Top

    @Top.setter
    def Top(self, value):
        self.oleobjects.Top = value

    @property
    def Visible(self):
        return self.oleobjects.Visible

    @Visible.setter
    def Visible(self, value):
        self.oleobjects.Visible = value

    @property
    def Width(self):
        return self.oleobjects.Width

    @Width.setter
    def Width(self, value):
        self.oleobjects.Width = value

    @property
    def ZOrder(self):
        return self.oleobjects.ZOrder

    def Add(self, ClassType=None, FileName=None, Link=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Left=None, Top=None, Width=None, Height=None):
        params = [
            ClassType if ClassType is not None else pythoncom.Missing,
            FileName if FileName is not None else pythoncom.Missing,
            Link if Link is not None else pythoncom.Missing,
            DisplayAsIcon if DisplayAsIcon is not None else pythoncom.Missing,
            IconFileName if IconFileName is not None else pythoncom.Missing,
            IconIndex if IconIndex is not None else pythoncom.Missing,
            IconLabel if IconLabel is not None else pythoncom.Missing,
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            Width if Width is not None else pythoncom.Missing,
            Height if Height is not None else pythoncom.Missing,
        ]
        return OLEObject(self.oleobjects.Add(*params))

    def BringToFront(self):
        return self.oleobjects.BringToFront()

    def Copy(self):
        return self.oleobjects.Copy()

    def CopyPicture(self, Appearance=None, Format=None):
        params = [
            Appearance if Appearance is not None else pythoncom.Missing,
            Format if Format is not None else pythoncom.Missing,
        ]
        return self.oleobjects.CopyPicture(*params)

    def Cut(self):
        return self.oleobjects.Cut()

    def Delete(self):
        return self.oleobjects.Delete()

    def Duplicate(self):
        return self.oleobjects.Duplicate()

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.oleobjects.Item(*params)

    def Select(self, Replace=None):
        params = [
            Replace if Replace is not None else pythoncom.Missing,
        ]
        return self.oleobjects.Select(*params)

    def SendToBack(self):
        return self.oleobjects.SendToBack()


class Outline:

    def __init__(self, outline=None):
        self.outline = outline

    @property
    def Application(self):
        return self.outline.Application

    @property
    def AutomaticStyles(self):
        return self.outline.AutomaticStyles

    @AutomaticStyles.setter
    def AutomaticStyles(self, value):
        self.outline.AutomaticStyles = value

    @property
    def Creator(self):
        return self.outline.Creator

    @property
    def Parent(self):
        return self.outline.Parent

    @property
    def SummaryColumn(self):
        return XlSummaryColumn(self.outline.SummaryColumn)

    @SummaryColumn.setter
    def SummaryColumn(self, value):
        self.outline.SummaryColumn = value

    @property
    def SummaryRow(self):
        return XlSummaryRow(self.outline.SummaryRow)

    @SummaryRow.setter
    def SummaryRow(self, value):
        self.outline.SummaryRow = value

    def ShowLevels(self, RowLevels=None, ColumnLevels=None):
        params = [
            RowLevels if RowLevels is not None else pythoncom.Missing,
            ColumnLevels if ColumnLevels is not None else pythoncom.Missing,
        ]
        return self.outline.ShowLevels(*params)


class Page:

    def __init__(self, page=None):
        self.page = page

    @property
    def CenterFooter(self):
        return self.page.CenterFooter

    @property
    def CenterHeader(self):
        return self.page.CenterHeader

    @property
    def LeftFooter(self):
        return self.page.LeftFooter

    @property
    def LeftHeader(self):
        return self.page.LeftHeader

    @property
    def RightFooter(self):
        return self.page.RightFooter

    @property
    def RightHeader(self):
        return self.page.RightHeader


class Pages:

    def __init__(self, pages=None):
        self.pages = pages

    def __call__(self, item):
        return Page(self.pages(item))

    @property
    def Count(self):
        return self.pages.Count

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.pages.Item):
            return Page(self.pages.Item(*params))
        else:
            return Page(self.pages.GetItem(*params))


class PageSetup:

    def __init__(self, pagesetup=None):
        self.pagesetup = pagesetup

    @property
    def AlignMarginsHeaderFooter(self):
        return self.pagesetup.AlignMarginsHeaderFooter

    @AlignMarginsHeaderFooter.setter
    def AlignMarginsHeaderFooter(self, value):
        self.pagesetup.AlignMarginsHeaderFooter = value

    @property
    def Application(self):
        return self.pagesetup.Application

    @property
    def BlackAndWhite(self):
        return self.pagesetup.BlackAndWhite

    @BlackAndWhite.setter
    def BlackAndWhite(self, value):
        self.pagesetup.BlackAndWhite = value

    @property
    def BottomMargin(self):
        return self.pagesetup.BottomMargin

    @BottomMargin.setter
    def BottomMargin(self, value):
        self.pagesetup.BottomMargin = value

    @property
    def CenterFooter(self):
        return self.pagesetup.CenterFooter

    @CenterFooter.setter
    def CenterFooter(self, value):
        self.pagesetup.CenterFooter = value

    @property
    def CenterFooterPicture(self):
        return Graphic(self.pagesetup.CenterFooterPicture)

    @property
    def CenterHeader(self):
        return self.pagesetup.CenterHeader

    @CenterHeader.setter
    def CenterHeader(self, value):
        self.pagesetup.CenterHeader = value

    @property
    def CenterHeaderPicture(self):
        return Graphic(self.pagesetup.CenterHeaderPicture)

    @property
    def CenterHorizontally(self):
        return self.pagesetup.CenterHorizontally

    @CenterHorizontally.setter
    def CenterHorizontally(self, value):
        self.pagesetup.CenterHorizontally = value

    @property
    def CenterVertically(self):
        return self.pagesetup.CenterVertically

    @CenterVertically.setter
    def CenterVertically(self, value):
        self.pagesetup.CenterVertically = value

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
    def Draft(self):
        return self.pagesetup.Draft

    @Draft.setter
    def Draft(self, value):
        self.pagesetup.Draft = value

    @property
    def EvenPage(self):
        return self.pagesetup.EvenPage

    @EvenPage.setter
    def EvenPage(self, value):
        self.pagesetup.EvenPage = value

    @property
    def FirstPage(self):
        return self.pagesetup.FirstPage

    @FirstPage.setter
    def FirstPage(self, value):
        self.pagesetup.FirstPage = value

    @property
    def FirstPageNumber(self):
        return self.pagesetup.FirstPageNumber

    @FirstPageNumber.setter
    def FirstPageNumber(self, value):
        self.pagesetup.FirstPageNumber = value

    @property
    def FitToPagesTall(self):
        return self.pagesetup.FitToPagesTall

    @FitToPagesTall.setter
    def FitToPagesTall(self, value):
        self.pagesetup.FitToPagesTall = value

    @property
    def FitToPagesWide(self):
        return self.pagesetup.FitToPagesWide

    @FitToPagesWide.setter
    def FitToPagesWide(self, value):
        self.pagesetup.FitToPagesWide = value

    @property
    def FooterMargin(self):
        return self.pagesetup.FooterMargin

    @FooterMargin.setter
    def FooterMargin(self, value):
        self.pagesetup.FooterMargin = value

    @property
    def HeaderMargin(self):
        return self.pagesetup.HeaderMargin

    @HeaderMargin.setter
    def HeaderMargin(self, value):
        self.pagesetup.HeaderMargin = value

    @property
    def LeftFooter(self):
        return self.pagesetup.LeftFooter

    @LeftFooter.setter
    def LeftFooter(self, value):
        self.pagesetup.LeftFooter = value

    @property
    def LeftFooterPicture(self):
        return Graphic(self.pagesetup.LeftFooterPicture)

    @property
    def LeftHeader(self):
        return self.pagesetup.LeftHeader

    @LeftHeader.setter
    def LeftHeader(self, value):
        self.pagesetup.LeftHeader = value

    @property
    def LeftHeaderPicture(self):
        return Graphic(self.pagesetup.LeftHeaderPicture)

    @property
    def LeftMargin(self):
        return self.pagesetup.LeftMargin

    @LeftMargin.setter
    def LeftMargin(self, value):
        self.pagesetup.LeftMargin = value

    @property
    def OddAndEvenPagesHeaderFooter(self):
        return self.pagesetup.OddAndEvenPagesHeaderFooter

    @OddAndEvenPagesHeaderFooter.setter
    def OddAndEvenPagesHeaderFooter(self, value):
        self.pagesetup.OddAndEvenPagesHeaderFooter = value

    @property
    def Order(self):
        return XlOrder(self.pagesetup.Order)

    @Order.setter
    def Order(self, value):
        self.pagesetup.Order = value

    @property
    def Orientation(self):
        return XlPageOrientation(self.pagesetup.Orientation)

    @Orientation.setter
    def Orientation(self, value):
        self.pagesetup.Orientation = value

    @property
    def Pages(self):
        return Pages(self.pagesetup.Pages)

    @Pages.setter
    def Pages(self, value):
        self.pagesetup.Pages = value

    @property
    def PaperSize(self):
        return XlPaperSize(self.pagesetup.PaperSize)

    @PaperSize.setter
    def PaperSize(self, value):
        self.pagesetup.PaperSize = value

    @property
    def Parent(self):
        return self.pagesetup.Parent

    @property
    def PrintArea(self):
        return self.pagesetup.PrintArea

    @PrintArea.setter
    def PrintArea(self, value):
        self.pagesetup.PrintArea = value

    @property
    def PrintComments(self):
        return XlPrintLocation(self.pagesetup.PrintComments)

    @PrintComments.setter
    def PrintComments(self, value):
        self.pagesetup.PrintComments = value

    @property
    def PrintErrors(self):
        return self.pagesetup.PrintErrors

    @PrintErrors.setter
    def PrintErrors(self, value):
        self.pagesetup.PrintErrors = value

    @property
    def PrintGridlines(self):
        return self.pagesetup.PrintGridlines

    @PrintGridlines.setter
    def PrintGridlines(self, value):
        self.pagesetup.PrintGridlines = value

    @property
    def PrintHeadings(self):
        return self.pagesetup.PrintHeadings

    @PrintHeadings.setter
    def PrintHeadings(self, value):
        self.pagesetup.PrintHeadings = value

    @property
    def PrintNotes(self):
        return self.pagesetup.PrintNotes

    @PrintNotes.setter
    def PrintNotes(self, value):
        self.pagesetup.PrintNotes = value

    @property
    def PrintQuality(self):
        return self.pagesetup.PrintQuality

    @PrintQuality.setter
    def PrintQuality(self, value):
        self.pagesetup.PrintQuality = value

    @property
    def PrintTitleColumns(self):
        return self.pagesetup.PrintTitleColumns

    @PrintTitleColumns.setter
    def PrintTitleColumns(self, value):
        self.pagesetup.PrintTitleColumns = value

    @property
    def PrintTitleRows(self):
        return self.pagesetup.PrintTitleRows

    @PrintTitleRows.setter
    def PrintTitleRows(self, value):
        self.pagesetup.PrintTitleRows = value

    @property
    def RightFooter(self):
        return self.pagesetup.RightFooter

    @RightFooter.setter
    def RightFooter(self, value):
        self.pagesetup.RightFooter = value

    @property
    def RightFooterPicture(self):
        return Graphic(self.pagesetup.RightFooterPicture)

    @property
    def RightHeader(self):
        return self.pagesetup.RightHeader

    @RightHeader.setter
    def RightHeader(self, value):
        self.pagesetup.RightHeader = value

    @property
    def RightHeaderPicture(self):
        return Graphic(self.pagesetup.RightHeaderPicture)

    @property
    def RightMargin(self):
        return self.pagesetup.RightMargin

    @RightMargin.setter
    def RightMargin(self, value):
        self.pagesetup.RightMargin = value

    @property
    def ScaleWithDocHeaderFooter(self):
        return self.pagesetup.ScaleWithDocHeaderFooter

    @ScaleWithDocHeaderFooter.setter
    def ScaleWithDocHeaderFooter(self, value):
        self.pagesetup.ScaleWithDocHeaderFooter = value

    @property
    def TopMargin(self):
        return self.pagesetup.TopMargin

    @TopMargin.setter
    def TopMargin(self, value):
        self.pagesetup.TopMargin = value

    @property
    def Zoom(self):
        return self.pagesetup.Zoom

    @Zoom.setter
    def Zoom(self, value):
        self.pagesetup.Zoom = value


class Pane:

    def __init__(self, pane=None):
        self.pane = pane

    @property
    def Application(self):
        return self.pane.Application

    @property
    def Creator(self):
        return self.pane.Creator

    @property
    def Index(self):
        return self.pane.Index

    @property
    def Parent(self):
        return self.pane.Parent

    @property
    def ScrollColumn(self):
        return self.pane.ScrollColumn

    @ScrollColumn.setter
    def ScrollColumn(self, value):
        self.pane.ScrollColumn = value

    @property
    def ScrollRow(self):
        return self.pane.ScrollRow

    @ScrollRow.setter
    def ScrollRow(self, value):
        self.pane.ScrollRow = value

    @property
    def VisibleRange(self):
        return Range(self.pane.VisibleRange)

    def Activate(self):
        return self.pane.Activate()

    def LargeScroll(self, Down=None, Up=None, ToRight=None, ToLeft=None):
        params = [
            Down if Down is not None else pythoncom.Missing,
            Up if Up is not None else pythoncom.Missing,
            ToRight if ToRight is not None else pythoncom.Missing,
            ToLeft if ToLeft is not None else pythoncom.Missing,
        ]
        return self.pane.LargeScroll(*params)

    def PointsToScreenPixelsX(self, Points=None):
        params = [
            Points if Points is not None else pythoncom.Missing,
        ]
        return self.pane.PointsToScreenPixelsX(*params)

    def PointsToScreenPixelsY(self, Points=None):
        params = [
            Points if Points is not None else pythoncom.Missing,
        ]
        return self.pane.PointsToScreenPixelsY(*params)

    def ScrollIntoView(self, Left=None, Top=None, Width=None, Height=None, Start=None):
        params = [
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            Width if Width is not None else pythoncom.Missing,
            Height if Height is not None else pythoncom.Missing,
            Start if Start is not None else pythoncom.Missing,
        ]
        self.pane.ScrollIntoView(*params)

    def SmallScroll(self, Down=None, Up=None, ToRight=None, ToLeft=None):
        params = [
            Down if Down is not None else pythoncom.Missing,
            Up if Up is not None else pythoncom.Missing,
            ToRight if ToRight is not None else pythoncom.Missing,
            ToLeft if ToLeft is not None else pythoncom.Missing,
        ]
        return self.pane.SmallScroll(*params)


class Panes:

    def __init__(self, panes=None):
        self.panes = panes

    def __call__(self, item):
        return Pane(self.panes(item))

    @property
    def Application(self):
        return self.panes.Application

    @property
    def Count(self):
        return self.panes.Count

    @property
    def Creator(self):
        return self.panes.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.panes.Item):
            return self.panes.Item(*params)
        else:
            return self.panes.GetItem(*params)

    @property
    def Parent(self):
        return self.panes.Parent


class Parameter:

    def __init__(self, parameter=None):
        self.parameter = parameter

    @property
    def Application(self):
        return self.parameter.Application

    @property
    def Creator(self):
        return self.parameter.Creator

    @property
    def DataType(self):
        return XlParameterDataType(self.parameter.DataType)

    @DataType.setter
    def DataType(self, value):
        self.parameter.DataType = value

    @property
    def Name(self):
        return self.parameter.Name

    @Name.setter
    def Name(self, value):
        self.parameter.Name = value

    @property
    def Parent(self):
        return self.parameter.Parent

    @property
    def PromptString(self):
        return self.parameter.PromptString

    @property
    def RefreshOnChange(self):
        return self.parameter.RefreshOnChange

    @RefreshOnChange.setter
    def RefreshOnChange(self, value):
        self.parameter.RefreshOnChange = value

    @property
    def SourceRange(self):
        return Range(self.parameter.SourceRange)

    @property
    def Type(self):
        return XlParameterType(self.parameter.Type)

    @property
    def Value(self):
        return self.parameter.Value

    def SetParam(self, Type=None, Value=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            Value if Value is not None else pythoncom.Missing,
        ]
        self.parameter.SetParam(*params)


class Parameters:

    def __init__(self, parameters=None):
        self.parameters = parameters

    def __call__(self, item):
        return Parameter(self.parameters(item))

    @property
    def Application(self):
        return self.parameters.Application

    @property
    def Count(self):
        return self.parameters.Count

    @property
    def Creator(self):
        return self.parameters.Creator

    @property
    def Parent(self):
        return self.parameters.Parent

    def Add(self, Name=None, iDataType=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
            iDataType if iDataType is not None else pythoncom.Missing,
        ]
        return Parameter(self.parameters.Add(*params))

    def Delete(self):
        self.parameters.Delete()

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return Parameter(self.parameters.Item(*params))


class Phonetic:

    def __init__(self, phonetic=None):
        self.phonetic = phonetic

    @property
    def Alignment(self):
        return self.phonetic.Alignment

    @Alignment.setter
    def Alignment(self, value):
        self.phonetic.Alignment = value

    @property
    def Application(self):
        return self.phonetic.Application

    @property
    def CharacterType(self):
        return XlPhoneticCharacterType(self.phonetic.CharacterType)

    @CharacterType.setter
    def CharacterType(self, value):
        self.phonetic.CharacterType = value

    @property
    def Creator(self):
        return self.phonetic.Creator

    @property
    def Font(self):
        return Font(self.phonetic.Font)

    @property
    def Parent(self):
        return self.phonetic.Parent

    @property
    def Text(self):
        return self.phonetic.Text

    @Text.setter
    def Text(self, value):
        self.phonetic.Text = value

    @property
    def Visible(self):
        return self.phonetic.Visible

    @Visible.setter
    def Visible(self, value):
        self.phonetic.Visible = value


class Phonetics:

    def __init__(self, phonetics=None):
        self.phonetics = phonetics

    def __call__(self, item):
        return Phonetic(self.phonetics(item))

    @property
    def Alignment(self):
        return self.phonetics.Alignment

    @Alignment.setter
    def Alignment(self, value):
        self.phonetics.Alignment = value

    @property
    def Application(self):
        return self.phonetics.Application

    @property
    def CharacterType(self):
        return XlPhoneticCharacterType(self.phonetics.CharacterType)

    @CharacterType.setter
    def CharacterType(self, value):
        self.phonetics.CharacterType = value

    @property
    def Count(self):
        return self.phonetics.Count

    @property
    def Creator(self):
        return self.phonetics.Creator

    @property
    def Font(self):
        return Font(self.phonetics.Font)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.phonetics.Item):
            return self.phonetics.Item(*params)
        else:
            return self.phonetics.GetItem(*params)

    @property
    def Length(self):
        return self.phonetics.Length

    @property
    def Parent(self):
        return self.phonetics.Parent

    @property
    def Start(self):
        return self.phonetics.Start

    @property
    def Text(self):
        return self.phonetics.Text

    @Text.setter
    def Text(self, value):
        self.phonetics.Text = value

    @property
    def Visible(self):
        return self.phonetics.Visible

    @Visible.setter
    def Visible(self, value):
        self.phonetics.Visible = value

    def Add(self, Start=None, Length=None, Text=None):
        params = [
            Start if Start is not None else pythoncom.Missing,
            Length if Length is not None else pythoncom.Missing,
            Text if Text is not None else pythoncom.Missing,
        ]
        self.phonetics.Add(*params)

    def Delete(self):
        self.phonetics.Delete()


class PictureFormat:

    def __init__(self, pictureformat=None):
        self.pictureformat = pictureformat

    @property
    def Application(self):
        return self.pictureformat.Application

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
        params = [
            Increment if Increment is not None else pythoncom.Missing,
        ]
        self.pictureformat.IncrementBrightness(*params)

    def IncrementContrast(self, Increment=None):
        params = [
            Increment if Increment is not None else pythoncom.Missing,
        ]
        self.pictureformat.IncrementContrast(*params)


class PivotAxis:

    def __init__(self, pivotaxis=None):
        self.pivotaxis = pivotaxis

    @property
    def Application(self):
        return self.pivotaxis.Application

    @property
    def Creator(self):
        return self.pivotaxis.Creator

    @property
    def Parent(self):
        return PivotAxis(self.pivotaxis.Parent)

    @property
    def PivotLines(self):
        return PivotLines(self.pivotaxis.PivotLines)


class PivotCache:

    def __init__(self, pivotcache=None):
        self.pivotcache = pivotcache

    @property
    def ADOConnection(self):
        return self.pivotcache.ADOConnection

    @property
    def Application(self):
        return self.pivotcache.Application

    @property
    def BackgroundQuery(self):
        return self.pivotcache.BackgroundQuery

    @BackgroundQuery.setter
    def BackgroundQuery(self, value):
        self.pivotcache.BackgroundQuery = value

    @property
    def CommandText(self):
        return self.pivotcache.CommandText

    @CommandText.setter
    def CommandText(self, value):
        self.pivotcache.CommandText = value

    @property
    def CommandType(self):
        return XlCmdType(self.pivotcache.CommandType)

    @CommandType.setter
    def CommandType(self, value):
        self.pivotcache.CommandType = value

    @property
    def Connection(self):
        return self.pivotcache.Connection

    @Connection.setter
    def Connection(self, value):
        self.pivotcache.Connection = value

    @property
    def Creator(self):
        return self.pivotcache.Creator

    @property
    def EnableRefresh(self):
        return self.pivotcache.EnableRefresh

    @EnableRefresh.setter
    def EnableRefresh(self, value):
        self.pivotcache.EnableRefresh = value

    @property
    def Index(self):
        return self.pivotcache.Index

    @property
    def IsConnected(self):
        return self.pivotcache.IsConnected

    @property
    def LocalConnection(self):
        return self.pivotcache.LocalConnection

    @LocalConnection.setter
    def LocalConnection(self, value):
        self.pivotcache.LocalConnection = value

    @property
    def MaintainConnection(self):
        return self.pivotcache.MaintainConnection

    @MaintainConnection.setter
    def MaintainConnection(self, value):
        self.pivotcache.MaintainConnection = value

    @property
    def MemoryUsed(self):
        return self.pivotcache.MemoryUsed

    @property
    def MissingItemsLimit(self):
        return XlPivotTableMissingItems(self.pivotcache.MissingItemsLimit)

    @MissingItemsLimit.setter
    def MissingItemsLimit(self, value):
        self.pivotcache.MissingItemsLimit = value

    @property
    def OLAP(self):
        return self.pivotcache.OLAP

    @property
    def OptimizeCache(self):
        return self.pivotcache.OptimizeCache

    @OptimizeCache.setter
    def OptimizeCache(self, value):
        self.pivotcache.OptimizeCache = value

    @property
    def Parent(self):
        return self.pivotcache.Parent

    @property
    def QueryType(self):
        return self.pivotcache.QueryType

    @property
    def RecordCount(self):
        return self.pivotcache.RecordCount

    @property
    def Recordset(self):
        return self.pivotcache.Recordset

    @Recordset.setter
    def Recordset(self, value):
        self.pivotcache.Recordset = value

    @property
    def RefreshDate(self):
        return self.pivotcache.RefreshDate

    @property
    def RefreshName(self):
        return self.pivotcache.RefreshName

    @property
    def RefreshOnFileOpen(self):
        return self.pivotcache.RefreshOnFileOpen

    @RefreshOnFileOpen.setter
    def RefreshOnFileOpen(self, value):
        self.pivotcache.RefreshOnFileOpen = value

    @property
    def RefreshPeriod(self):
        return self.pivotcache.RefreshPeriod

    @RefreshPeriod.setter
    def RefreshPeriod(self, value):
        self.pivotcache.RefreshPeriod = value

    @property
    def RobustConnect(self):
        return XlRobustConnect(self.pivotcache.RobustConnect)

    @RobustConnect.setter
    def RobustConnect(self, value):
        self.pivotcache.RobustConnect = value

    @property
    def SavePassword(self):
        return self.pivotcache.SavePassword

    @SavePassword.setter
    def SavePassword(self, value):
        self.pivotcache.SavePassword = value

    @property
    def SourceConnectionFile(self):
        return self.pivotcache.SourceConnectionFile

    @SourceConnectionFile.setter
    def SourceConnectionFile(self, value):
        self.pivotcache.SourceConnectionFile = value

    @property
    def SourceData(self):
        return self.pivotcache.SourceData

    @SourceData.setter
    def SourceData(self, value):
        self.pivotcache.SourceData = value

    @property
    def SourceDataFile(self):
        return self.pivotcache.SourceDataFile

    @property
    def SourceType(self):
        return XlPivotTableSourceType(self.pivotcache.SourceType)

    @property
    def UpgradeOnRefresh(self):
        return self.pivotcache.UpgradeOnRefresh

    @UpgradeOnRefresh.setter
    def UpgradeOnRefresh(self, value):
        self.pivotcache.UpgradeOnRefresh = value

    @property
    def UseLocalConnection(self):
        return self.pivotcache.UseLocalConnection

    @UseLocalConnection.setter
    def UseLocalConnection(self, value):
        self.pivotcache.UseLocalConnection = value

    @property
    def Version(self):
        return XlPivotTableVersionList(self.pivotcache.Version)

    @property
    def WorkbookConnection(self):
        return self.pivotcache.WorkbookConnection

    def CreatePivotTable(self, TableDestination=None, TableName=None, ReadData=None, DefaultVersion=None):
        params = [
            TableDestination if TableDestination is not None else pythoncom.Missing,
            TableName if TableName is not None else pythoncom.Missing,
            ReadData if ReadData is not None else pythoncom.Missing,
            DefaultVersion if DefaultVersion is not None else pythoncom.Missing,
        ]
        return self.pivotcache.CreatePivotTable(*params)

    def MakeConnection(self):
        self.pivotcache.MakeConnection()

    def Refresh(self):
        self.pivotcache.Refresh()

    def ResetTimer(self):
        self.pivotcache.ResetTimer()

    def SaveAsODC(self, ODCFileName=None, Description=None, Keywords=None):
        params = [
            ODCFileName if ODCFileName is not None else pythoncom.Missing,
            Description if Description is not None else pythoncom.Missing,
            Keywords if Keywords is not None else pythoncom.Missing,
        ]
        self.pivotcache.SaveAsODC(*params)


class PivotCaches:

    def __init__(self, pivotcaches=None):
        self.pivotcaches = pivotcaches

    @property
    def Application(self):
        return self.pivotcaches.Application

    @property
    def Count(self):
        return self.pivotcaches.Count

    @property
    def Creator(self):
        return self.pivotcaches.Creator

    @property
    def Parent(self):
        return self.pivotcaches.Parent

    def Create(self, SourceType=None, SourceData=None, Version=None):
        params = [
            SourceType if SourceType is not None else pythoncom.Missing,
            SourceData if SourceData is not None else pythoncom.Missing,
            Version if Version is not None else pythoncom.Missing,
        ]
        return self.pivotcaches.Create(*params)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return PivotCache(self.pivotcaches.Item(*params))


class PivotCell:

    def __init__(self, pivotcell=None):
        self.pivotcell = pivotcell

    @property
    def Application(self):
        return self.pivotcell.Application

    @property
    def CellChanged(self):
        return self.pivotcell.CellChanged

    @property
    def ColumnItems(self):
        return PivotItemList(self.pivotcell.ColumnItems)

    @property
    def Creator(self):
        return self.pivotcell.Creator

    @property
    def CustomSubtotalFunction(self):
        return XlConsolidationFunction(self.pivotcell.CustomSubtotalFunction)

    @property
    def DataField(self):
        return PivotField(self.pivotcell.DataField)

    @property
    def DataSourceValue(self):
        return self.pivotcell.DataSourceValue

    @property
    def MDX(self):
        return self.pivotcell.MDX

    @property
    def Parent(self):
        return self.pivotcell.Parent

    @property
    def PivotCellType(self):
        return XlPivotCellType(self.pivotcell.PivotCellType)

    @property
    def PivotColumnLine(self):
        return PivotLine(self.pivotcell.PivotColumnLine)

    @property
    def PivotField(self):
        return PivotField(self.pivotcell.PivotField)

    @property
    def PivotItem(self):
        return PivotItem(self.pivotcell.PivotItem)

    @property
    def PivotRowLine(self):
        return PivotLine(self.pivotcell.PivotRowLine)

    @property
    def PivotTable(self):
        return PivotTable(self.pivotcell.PivotTable)

    @property
    def Range(self):
        return Range(self.pivotcell.Range)

    @property
    def RowItems(self):
        return PivotItemList(self.pivotcell.RowItems)

    def AllocateChange(self):
        return self.pivotcell.AllocateChange()

    def DiscardChange(self):
        return self.pivotcell.DiscardChange()


class PivotField:

    def __init__(self, pivotfield=None):
        self.pivotfield = pivotfield

    @property
    def AllItemsVisible(self):
        return self.pivotfield.AllItemsVisible

    @property
    def Application(self):
        return self.pivotfield.Application

    @property
    def AutoShowCount(self):
        return self.pivotfield.AutoShowCount

    @property
    def AutoShowField(self):
        return self.pivotfield.AutoShowField

    @property
    def AutoShowRange(self):
        return self.pivotfield.AutoShowRange

    @property
    def AutoShowType(self):
        return self.pivotfield.AutoShowType

    @property
    def AutoSortCustomSubtotal(self):
        return self.pivotfield.AutoSortCustomSubtotal

    @property
    def AutoSortField(self):
        return self.pivotfield.AutoSortField

    @property
    def AutoSortOrder(self):
        return XlSortOrder(self.pivotfield.AutoSortOrder)

    @property
    def AutoSortPivotLine(self):
        return PivotLine(self.pivotfield.AutoSortPivotLine)

    @property
    def BaseField(self):
        return self.pivotfield.BaseField

    @BaseField.setter
    def BaseField(self, value):
        self.pivotfield.BaseField = value

    @property
    def BaseItem(self):
        return self.pivotfield.BaseItem

    @BaseItem.setter
    def BaseItem(self, value):
        self.pivotfield.BaseItem = value

    @property
    def Calculation(self):
        return XlPivotFieldCalculation(self.pivotfield.Calculation)

    @Calculation.setter
    def Calculation(self, value):
        self.pivotfield.Calculation = value

    @property
    def Caption(self):
        return self.pivotfield.Caption

    @property
    def ChildField(self):
        return PivotField(self.pivotfield.ChildField)

    def ChildItems(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.pivotfield.ChildItems):
            return PivotItem(self.pivotfield.ChildItems(*params))
        else:
            return PivotItem(self.pivotfield.GetChildItems(*params))

    @property
    def Creator(self):
        return self.pivotfield.Creator

    @property
    def CubeField(self):
        return CubeField(self.pivotfield.CubeField)

    @property
    def CurrentPage(self):
        return PivotItem(self.pivotfield.CurrentPage)

    @CurrentPage.setter
    def CurrentPage(self, value):
        self.pivotfield.CurrentPage = value

    @property
    def CurrentPageList(self):
        return self.pivotfield.CurrentPageList

    @CurrentPageList.setter
    def CurrentPageList(self, value):
        self.pivotfield.CurrentPageList = value

    @property
    def CurrentPageName(self):
        return self.pivotfield.CurrentPageName

    @CurrentPageName.setter
    def CurrentPageName(self, value):
        self.pivotfield.CurrentPageName = value

    @property
    def DatabaseSort(self):
        return self.pivotfield.DatabaseSort

    @DatabaseSort.setter
    def DatabaseSort(self, value):
        self.pivotfield.DatabaseSort = value

    @property
    def DataRange(self):
        return Range(self.pivotfield.DataRange)

    @property
    def DataType(self):
        return XlPivotFieldDataType(self.pivotfield.DataType)

    @property
    def DisplayAsCaption(self):
        return self.pivotfield.DisplayAsCaption

    @property
    def DisplayAsTooltip(self):
        return self.pivotfield.DisplayAsTooltip

    @DisplayAsTooltip.setter
    def DisplayAsTooltip(self, value):
        self.pivotfield.DisplayAsTooltip = value

    @property
    def DisplayInReport(self):
        return self.pivotfield.DisplayInReport

    @DisplayInReport.setter
    def DisplayInReport(self, value):
        self.pivotfield.DisplayInReport = value

    @property
    def DragToColumn(self):
        return self.pivotfield.DragToColumn

    @DragToColumn.setter
    def DragToColumn(self, value):
        self.pivotfield.DragToColumn = value

    @property
    def DragToData(self):
        return self.pivotfield.DragToData

    @DragToData.setter
    def DragToData(self, value):
        self.pivotfield.DragToData = value

    @property
    def DragToHide(self):
        return self.pivotfield.DragToHide

    @DragToHide.setter
    def DragToHide(self, value):
        self.pivotfield.DragToHide = value

    @property
    def DragToPage(self):
        return self.pivotfield.DragToPage

    @DragToPage.setter
    def DragToPage(self, value):
        self.pivotfield.DragToPage = value

    @property
    def DragToRow(self):
        return self.pivotfield.DragToRow

    @DragToRow.setter
    def DragToRow(self, value):
        self.pivotfield.DragToRow = value

    @property
    def DrilledDown(self):
        return self.pivotfield.DrilledDown

    @DrilledDown.setter
    def DrilledDown(self, value):
        self.pivotfield.DrilledDown = value

    @property
    def EnableItemSelection(self):
        return self.pivotfield.EnableItemSelection

    @EnableItemSelection.setter
    def EnableItemSelection(self, value):
        self.pivotfield.EnableItemSelection = value

    @property
    def EnableMultiplePageItems(self):
        return self.pivotfield.EnableMultiplePageItems

    @EnableMultiplePageItems.setter
    def EnableMultiplePageItems(self, value):
        self.pivotfield.EnableMultiplePageItems = value

    @property
    def Formula(self):
        return self.pivotfield.Formula

    @Formula.setter
    def Formula(self, value):
        self.pivotfield.Formula = value

    @property
    def Function(self):
        return XlConsolidationFunction(self.pivotfield.Function)

    @Function.setter
    def Function(self, value):
        self.pivotfield.Function = value

    @property
    def GroupLevel(self):
        return self.pivotfield.GroupLevel

    @property
    def Hidden(self):
        return self.pivotfield.Hidden

    @Hidden.setter
    def Hidden(self, value):
        self.pivotfield.Hidden = value

    def HiddenItems(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.pivotfield.HiddenItems):
            return PivotItem(self.pivotfield.HiddenItems(*params))
        else:
            return PivotItem(self.pivotfield.GetHiddenItems(*params))

    @property
    def HiddenItemsList(self):
        return self.pivotfield.HiddenItemsList

    @HiddenItemsList.setter
    def HiddenItemsList(self, value):
        self.pivotfield.HiddenItemsList = value

    @property
    def IncludeNewItemsInFilter(self):
        return self.pivotfield.IncludeNewItemsInFilter

    @IncludeNewItemsInFilter.setter
    def IncludeNewItemsInFilter(self, value):
        self.pivotfield.IncludeNewItemsInFilter = value

    @property
    def IsCalculated(self):
        return self.pivotfield.IsCalculated

    @property
    def IsMemberProperty(self):
        return self.pivotfield.IsMemberProperty

    @property
    def LabelRange(self):
        return Range(self.pivotfield.LabelRange)

    @property
    def LayoutBlankLine(self):
        return self.pivotfield.LayoutBlankLine

    @LayoutBlankLine.setter
    def LayoutBlankLine(self, value):
        self.pivotfield.LayoutBlankLine = value

    @property
    def LayoutCompactRow(self):
        return self.pivotfield.LayoutCompactRow

    @LayoutCompactRow.setter
    def LayoutCompactRow(self, value):
        self.pivotfield.LayoutCompactRow = value

    @property
    def LayoutForm(self):
        return XlLayoutFormType(self.pivotfield.LayoutForm)

    @LayoutForm.setter
    def LayoutForm(self, value):
        self.pivotfield.LayoutForm = value

    @property
    def LayoutPageBreak(self):
        return self.pivotfield.LayoutPageBreak

    @LayoutPageBreak.setter
    def LayoutPageBreak(self, value):
        self.pivotfield.LayoutPageBreak = value

    @property
    def LayoutSubtotalLocation(self):
        return XlSubtotalLocationType(self.pivotfield.LayoutSubtotalLocation)

    @LayoutSubtotalLocation.setter
    def LayoutSubtotalLocation(self, value):
        self.pivotfield.LayoutSubtotalLocation = value

    @property
    def MemberPropertyCaption(self):
        return self.pivotfield.MemberPropertyCaption

    @MemberPropertyCaption.setter
    def MemberPropertyCaption(self, value):
        self.pivotfield.MemberPropertyCaption = value

    @property
    def MemoryUsed(self):
        return self.pivotfield.MemoryUsed

    @property
    def Name(self):
        return self.pivotfield.Name

    @Name.setter
    def Name(self, value):
        self.pivotfield.Name = value

    @property
    def NumberFormat(self):
        return self.pivotfield.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value):
        self.pivotfield.NumberFormat = value

    @property
    def Orientation(self):
        return XlPivotFieldOrientation(self.pivotfield.Orientation)

    @Orientation.setter
    def Orientation(self, value):
        self.pivotfield.Orientation = value

    @property
    def Parent(self):
        return self.pivotfield.Parent

    @property
    def ParentField(self):
        return PivotField(self.pivotfield.ParentField)

    def ParentItems(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.pivotfield.ParentItems):
            return PivotItem(self.pivotfield.ParentItems(*params))
        else:
            return PivotItem(self.pivotfield.GetParentItems(*params))

    @property
    def PivotFilters(self):
        return PivotField(self.pivotfield.PivotFilters)

    @PivotFilters.setter
    def PivotFilters(self, value):
        self.pivotfield.PivotFilters = value

    @property
    def Position(self):
        return self.pivotfield.Position

    @Position.setter
    def Position(self, value):
        self.pivotfield.Position = value

    @property
    def PropertyOrder(self):
        return self.pivotfield.PropertyOrder

    @PropertyOrder.setter
    def PropertyOrder(self, value):
        self.pivotfield.PropertyOrder = value

    @property
    def PropertyParentField(self):
        return PivotField(self.pivotfield.PropertyParentField)

    @property
    def RepeatLabels(self):
        return self.pivotfield.RepeatLabels

    @RepeatLabels.setter
    def RepeatLabels(self, value):
        self.pivotfield.RepeatLabels = value

    @property
    def ServerBased(self):
        return self.pivotfield.ServerBased

    @ServerBased.setter
    def ServerBased(self, value):
        self.pivotfield.ServerBased = value

    @property
    def ShowAllItems(self):
        return self.pivotfield.ShowAllItems

    @ShowAllItems.setter
    def ShowAllItems(self, value):
        self.pivotfield.ShowAllItems = value

    @property
    def ShowDetail(self):
        return self.pivotfield.ShowDetail

    @ShowDetail.setter
    def ShowDetail(self, value):
        self.pivotfield.ShowDetail = value

    @property
    def ShowingInAxis(self):
        return self.pivotfield.ShowingInAxis

    @property
    def SourceCaption(self):
        return self.pivotfield.SourceCaption

    @property
    def SourceName(self):
        return self.pivotfield.SourceName

    @property
    def StandardFormula(self):
        return self.pivotfield.StandardFormula

    @StandardFormula.setter
    def StandardFormula(self, value):
        self.pivotfield.StandardFormula = value

    @property
    def SubtotalName(self):
        return self.pivotfield.SubtotalName

    @SubtotalName.setter
    def SubtotalName(self, value):
        self.pivotfield.SubtotalName = value

    @property
    def Subtotals(self):
        return self.pivotfield.Subtotals

    @Subtotals.setter
    def Subtotals(self, value):
        self.pivotfield.Subtotals = value

    @property
    def TotalLevels(self):
        return self.pivotfield.TotalLevels

    @property
    def UseMemberPropertyAsCaption(self):
        return self.pivotfield.UseMemberPropertyAsCaption

    @UseMemberPropertyAsCaption.setter
    def UseMemberPropertyAsCaption(self, value):
        self.pivotfield.UseMemberPropertyAsCaption = value

    @property
    def Value(self):
        return self.pivotfield.Value

    @Value.setter
    def Value(self, value):
        self.pivotfield.Value = value

    def VisibleItems(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.pivotfield.VisibleItems):
            return PivotItem(self.pivotfield.VisibleItems(*params))
        else:
            return PivotItem(self.pivotfield.GetVisibleItems(*params))

    @property
    def VisibleItemsList(self):
        return self.pivotfield.VisibleItemsList

    @VisibleItemsList.setter
    def VisibleItemsList(self, value):
        self.pivotfield.VisibleItemsList = value

    def AddPageItem(self, Item=None, ClearList=None):
        params = [
            Item if Item is not None else pythoncom.Missing,
            ClearList if ClearList is not None else pythoncom.Missing,
        ]
        self.pivotfield.AddPageItem(*params)

    def AutoShow(self, Type=None, Range=None, Count=None, Field=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            Range if Range is not None else pythoncom.Missing,
            Count if Count is not None else pythoncom.Missing,
            Field if Field is not None else pythoncom.Missing,
        ]
        self.pivotfield.AutoShow(*params)

    def AutoSort(self, Order=None, Field=None, PivotLine=None, CustomSubtotal=None):
        params = [
            Order if Order is not None else pythoncom.Missing,
            Field if Field is not None else pythoncom.Missing,
            PivotLine if PivotLine is not None else pythoncom.Missing,
            CustomSubtotal if CustomSubtotal is not None else pythoncom.Missing,
        ]
        self.pivotfield.AutoSort(*params)

    def CalculatedItems(self):
        return self.pivotfield.CalculatedItems()

    def ClearAllFilters(self):
        self.pivotfield.ClearAllFilters()

    def ClearLabelFilters(self):
        self.pivotfield.ClearLabelFilters()

    def ClearManualFilter(self):
        self.pivotfield.ClearManualFilter()

    def ClearValueFilters(self):
        self.pivotfield.ClearValueFilters()

    def Delete(self):
        self.pivotfield.Delete()

    def DrillTo(self, PivotFieldName=None):
        params = [
            PivotFieldName if PivotFieldName is not None else pythoncom.Missing,
        ]
        self.pivotfield.DrillTo(*params)

    def PivotItems(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.pivotfield.PivotItems(*params)


class PivotFields:

    def __init__(self, pivotfields=None):
        self.pivotfields = pivotfields

    def __call__(self, item):
        return PivotField(self.pivotfields(item))

    @property
    def Application(self):
        return self.pivotfields.Application

    @property
    def Count(self):
        return self.pivotfields.Count

    @property
    def Creator(self):
        return self.pivotfields.Creator

    @property
    def Parent(self):
        return self.pivotfields.Parent

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.pivotfields.Item(*params)


class PivotFilter:

    def __init__(self, pivotfilter=None):
        self.pivotfilter = pivotfilter

    @property
    def Active(self):
        return self.pivotfilter.Active

    @property
    def Application(self):
        return self.pivotfilter.Application

    @property
    def Creator(self):
        return self.pivotfilter.Creator

    @property
    def DataCubeField(self):
        return self.pivotfilter.DataCubeField

    @DataCubeField.setter
    def DataCubeField(self, value):
        self.pivotfilter.DataCubeField = value

    @property
    def DataField(self):
        return self.pivotfilter.DataField

    @DataField.setter
    def DataField(self, value):
        self.pivotfilter.DataField = value

    @property
    def Description(self):
        return self.pivotfilter.Description

    @property
    def FilterType(self):
        return self.pivotfilter.FilterType

    @property
    def IsMemberPropertyFilter(self):
        return self.pivotfilter.IsMemberPropertyFilter

    @property
    def MemberPropertyField(self):
        return self.pivotfilter.MemberPropertyField

    @MemberPropertyField.setter
    def MemberPropertyField(self, value):
        self.pivotfilter.MemberPropertyField = value

    @property
    def Name(self):
        return self.pivotfilter.Name

    @property
    def Order(self):
        return self.pivotfilter.Order

    @Order.setter
    def Order(self, value):
        self.pivotfilter.Order = value

    @property
    def Parent(self):
        return PivotFilter(self.pivotfilter.Parent)

    @property
    def PivotField(self):
        return self.pivotfilter.PivotField

    @property
    def Value1(self):
        return self.pivotfilter.Value1

    @Value1.setter
    def Value1(self, value):
        self.pivotfilter.Value1 = value

    @property
    def Value2(self):
        return self.pivotfilter.Value2

    @Value2.setter
    def Value2(self, value):
        self.pivotfilter.Value2 = value

    def Delete(self):
        self.pivotfilter.Delete()


class PivotFilters:

    def __init__(self, pivotfilters=None):
        self.pivotfilters = pivotfilters

    @property
    def Application(self):
        return self.pivotfilters.Application

    @property
    def Count(self):
        return PivotFilters(self.pivotfilters.Count)

    @property
    def Creator(self):
        return self.pivotfilters.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.pivotfilters.Item):
            return PivotFilters(self.pivotfilters.Item(*params))
        else:
            return PivotFilters(self.pivotfilters.GetItem(*params))

    @property
    def Parent(self):
        return PivotFilters(self.pivotfilters.Parent)

    def Add(self, Type=None, DataField=None, Value1=None, Value2=None, Order=None, Name=None, Description=None, MemberPropertyField=None, WholeDayFilter=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            DataField if DataField is not None else pythoncom.Missing,
            Value1 if Value1 is not None else pythoncom.Missing,
            Value2 if Value2 is not None else pythoncom.Missing,
            Order if Order is not None else pythoncom.Missing,
            Name if Name is not None else pythoncom.Missing,
            Description if Description is not None else pythoncom.Missing,
            MemberPropertyField if MemberPropertyField is not None else pythoncom.Missing,
            WholeDayFilter if WholeDayFilter is not None else pythoncom.Missing,
        ]
        return self.pivotfilters.Add(*params)


class PivotFormula:

    def __init__(self, pivotformula=None):
        self.pivotformula = pivotformula

    @property
    def Application(self):
        return self.pivotformula.Application

    @property
    def Creator(self):
        return self.pivotformula.Creator

    @property
    def Formula(self):
        return self.pivotformula.Formula

    @Formula.setter
    def Formula(self, value):
        self.pivotformula.Formula = value

    @property
    def Index(self):
        return PivotFormulas(self.pivotformula.Index)

    @Index.setter
    def Index(self, value):
        self.pivotformula.Index = value

    @property
    def Parent(self):
        return self.pivotformula.Parent

    @property
    def StandardFormula(self):
        return self.pivotformula.StandardFormula

    @StandardFormula.setter
    def StandardFormula(self, value):
        self.pivotformula.StandardFormula = value

    @property
    def Value(self):
        return self.pivotformula.Value

    @Value.setter
    def Value(self, value):
        self.pivotformula.Value = value

    def Delete(self):
        self.pivotformula.Delete()


class PivotFormulas:

    def __init__(self, pivotformulas=None):
        self.pivotformulas = pivotformulas

    @property
    def Application(self):
        return self.pivotformulas.Application

    @property
    def Count(self):
        return self.pivotformulas.Count

    @property
    def Creator(self):
        return self.pivotformulas.Creator

    @property
    def Parent(self):
        return self.pivotformulas.Parent

    def Add(self, Formula=None, UseStandardFormula=None):
        params = [
            Formula if Formula is not None else pythoncom.Missing,
            UseStandardFormula if UseStandardFormula is not None else pythoncom.Missing,
        ]
        return PivotFormula(self.pivotformulas.Add(*params))

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return PivotFormula(self.pivotformulas.Item(*params))


class PivotItem:

    def __init__(self, pivotitem=None):
        self.pivotitem = pivotitem

    @property
    def Application(self):
        return self.pivotitem.Application

    @property
    def Caption(self):
        return self.pivotitem.Caption

    def ChildItems(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.pivotitem.ChildItems):
            return PivotItem(self.pivotitem.ChildItems(*params))
        else:
            return PivotItem(self.pivotitem.GetChildItems(*params))

    @property
    def Creator(self):
        return self.pivotitem.Creator

    @property
    def DataRange(self):
        return Range(self.pivotitem.DataRange)

    @property
    def DrilledDown(self):
        return self.pivotitem.DrilledDown

    @DrilledDown.setter
    def DrilledDown(self, value):
        self.pivotitem.DrilledDown = value

    @property
    def Formula(self):
        return self.pivotitem.Formula

    @Formula.setter
    def Formula(self, value):
        self.pivotitem.Formula = value

    @property
    def IsCalculated(self):
        return self.pivotitem.IsCalculated

    @property
    def LabelRange(self):
        return Range(self.pivotitem.LabelRange)

    @property
    def Name(self):
        return self.pivotitem.Name

    @Name.setter
    def Name(self, value):
        self.pivotitem.Name = value

    @property
    def Parent(self):
        return self.pivotitem.Parent

    @property
    def ParentItem(self):
        return PivotField(self.pivotitem.ParentItem)

    @property
    def ParentShowDetail(self):
        return self.pivotitem.ParentShowDetail

    @property
    def Position(self):
        return self.pivotitem.Position

    @Position.setter
    def Position(self, value):
        self.pivotitem.Position = value

    @property
    def RecordCount(self):
        return self.pivotitem.RecordCount

    @property
    def ShowDetail(self):
        return self.pivotitem.ShowDetail

    @ShowDetail.setter
    def ShowDetail(self, value):
        self.pivotitem.ShowDetail = value

    @property
    def SourceName(self):
        return self.pivotitem.SourceName

    @property
    def SourceNameStandard(self):
        return self.pivotitem.SourceNameStandard

    @property
    def StandardFormula(self):
        return self.pivotitem.StandardFormula

    @StandardFormula.setter
    def StandardFormula(self, value):
        self.pivotitem.StandardFormula = value

    @property
    def Value(self):
        return self.pivotitem.Value

    @Value.setter
    def Value(self, value):
        self.pivotitem.Value = value

    @property
    def Visible(self):
        return self.pivotitem.Visible

    @Visible.setter
    def Visible(self, value):
        self.pivotitem.Visible = value

    def Delete(self):
        self.pivotitem.Delete()

    def DrillTo(self, PivotItemName=None):
        params = [
            PivotItemName if PivotItemName is not None else pythoncom.Missing,
        ]
        self.pivotitem.DrillTo(*params)


class PivotItemList:

    def __init__(self, pivotitemlist=None):
        self.pivotitemlist = pivotitemlist

    def __call__(self, item):
        return PivotItemLis(self.pivotitemlist(item))

    @property
    def Application(self):
        return self.pivotitemlist.Application

    @property
    def Count(self):
        return self.pivotitemlist.Count

    @property
    def Creator(self):
        return self.pivotitemlist.Creator

    @property
    def Parent(self):
        return self.pivotitemlist.Parent

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return PivotItem(self.pivotitemlist.Item(*params))


class PivotItems:

    def __init__(self, pivotitems=None):
        self.pivotitems = pivotitems

    def __call__(self, item):
        return PivotItem(self.pivotitems(item))

    @property
    def Application(self):
        return self.pivotitems.Application

    @property
    def Count(self):
        return self.pivotitems.Count

    @property
    def Creator(self):
        return self.pivotitems.Creator

    @property
    def Parent(self):
        return self.pivotitems.Parent

    def Add(self, Name=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
        ]
        self.pivotitems.Add(*params)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.pivotitems.Item(*params)


class PivotLayout:

    def __init__(self, pivotlayout=None):
        self.pivotlayout = pivotlayout

    @property
    def Application(self):
        return self.pivotlayout.Application

    @property
    def Creator(self):
        return self.pivotlayout.Creator

    @property
    def Parent(self):
        return self.pivotlayout.Parent

    @property
    def PivotTable(self):
        return PivotTable(self.pivotlayout.PivotTable)


class PivotLine:

    def __init__(self, pivotline=None):
        self.pivotline = pivotline

    @property
    def Application(self):
        return self.pivotline.Application

    @property
    def Creator(self):
        return self.pivotline.Creator

    @property
    def LineType(self):
        return XlPivotLineType(self.pivotline.LineType)

    @property
    def Parent(self):
        return PivotLine(self.pivotline.Parent)

    @property
    def PivotLineCells(self):
        return PivotCell(self.pivotline.PivotLineCells)

    @property
    def Position(self):
        return PivotLine(self.pivotline.Position)

    @Position.setter
    def Position(self, value):
        self.pivotline.Position = value


class PivotLineCells:

    def __init__(self, pivotlinecells=None):
        self.pivotlinecells = pivotlinecells

    @property
    def Application(self):
        return self.pivotlinecells.Application

    @property
    def Count(self):
        return PivotLineCells(self.pivotlinecells.Count)

    @property
    def Creator(self):
        return self.pivotlinecells.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.pivotlinecells.Item):
            return PivotLineCells(self.pivotlinecells.Item(*params))
        else:
            return PivotLineCells(self.pivotlinecells.GetItem(*params))

    @property
    def Parent(self):
        return PivotLineCells(self.pivotlinecells.Parent)


class PivotLines:

    def __init__(self, pivotlines=None):
        self.pivotlines = pivotlines

    @property
    def Application(self):
        return self.pivotlines.Application

    @property
    def Count(self):
        return PivotLines(self.pivotlines.Count)

    @property
    def Creator(self):
        return self.pivotlines.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.pivotlines.Item):
            return PivotLines(self.pivotlines.Item(*params))
        else:
            return PivotLines(self.pivotlines.GetItem(*params))

    @property
    def Parent(self):
        return PivotLines(self.pivotlines.Parent)


class PivotTable:

    def __init__(self, pivottable=None):
        self.pivottable = pivottable

    @property
    def ActiveFilters(self):
        return self.pivottable.ActiveFilters

    @property
    def Allocation(self):
        return self.pivottable.Allocation

    @Allocation.setter
    def Allocation(self, value):
        self.pivottable.Allocation = value

    @property
    def AllocationMethod(self):
        return self.pivottable.AllocationMethod

    @AllocationMethod.setter
    def AllocationMethod(self, value):
        self.pivottable.AllocationMethod = value

    @property
    def AllocationValue(self):
        return self.pivottable.AllocationValue

    @AllocationValue.setter
    def AllocationValue(self, value):
        self.pivottable.AllocationValue = value

    @property
    def AllocationWeightExpression(self):
        return self.pivottable.AllocationWeightExpression

    @AllocationWeightExpression.setter
    def AllocationWeightExpression(self, value):
        self.pivottable.AllocationWeightExpression = value

    @property
    def AllowMultipleFilters(self):
        return self.pivottable.AllowMultipleFilters

    @AllowMultipleFilters.setter
    def AllowMultipleFilters(self, value):
        self.pivottable.AllowMultipleFilters = value

    @property
    def AlternativeText(self):
        return self.pivottable.AlternativeText

    @AlternativeText.setter
    def AlternativeText(self, value):
        self.pivottable.AlternativeText = value

    @property
    def Application(self):
        return self.pivottable.Application

    @property
    def CacheIndex(self):
        return self.pivottable.CacheIndex

    @CacheIndex.setter
    def CacheIndex(self, value):
        self.pivottable.CacheIndex = value

    @property
    def CalculatedMembers(self):
        return CalculatedMembers(self.pivottable.CalculatedMembers)

    @property
    def CalculatedMembersInFilters(self):
        return self.pivottable.CalculatedMembersInFilters

    @CalculatedMembersInFilters.setter
    def CalculatedMembersInFilters(self, value):
        self.pivottable.CalculatedMembersInFilters = value

    @property
    def ChangeList(self):
        return PivotTableChangeList(self.pivottable.ChangeList)

    def ColumnFields(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.pivottable.ColumnFields):
            return PivotField(self.pivottable.ColumnFields(*params))
        else:
            return PivotField(self.pivottable.GetColumnFields(*params))

    @property
    def ColumnGrand(self):
        return self.pivottable.ColumnGrand

    @ColumnGrand.setter
    def ColumnGrand(self, value):
        self.pivottable.ColumnGrand = value

    @property
    def ColumnRange(self):
        return Range(self.pivottable.ColumnRange)

    @property
    def CompactLayoutColumnHeader(self):
        return self.pivottable.CompactLayoutColumnHeader

    @CompactLayoutColumnHeader.setter
    def CompactLayoutColumnHeader(self, value):
        self.pivottable.CompactLayoutColumnHeader = value

    @property
    def CompactLayoutRowHeader(self):
        return self.pivottable.CompactLayoutRowHeader

    @CompactLayoutRowHeader.setter
    def CompactLayoutRowHeader(self, value):
        self.pivottable.CompactLayoutRowHeader = value

    @property
    def CompactRowIndent(self):
        return self.pivottable.CompactRowIndent

    @CompactRowIndent.setter
    def CompactRowIndent(self, value):
        self.pivottable.CompactRowIndent = value

    @property
    def Creator(self):
        return self.pivottable.Creator

    @property
    def CubeFields(self):
        return CubeFields(self.pivottable.CubeFields)

    @property
    def DataBodyRange(self):
        return Range(self.pivottable.DataBodyRange)

    def DataFields(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.pivottable.DataFields):
            return PivotField(self.pivottable.DataFields(*params))
        else:
            return PivotField(self.pivottable.GetDataFields(*params))

    @property
    def DataLabelRange(self):
        return Range(self.pivottable.DataLabelRange)

    @property
    def DataPivotField(self):
        return PivotField(self.pivottable.DataPivotField)

    @property
    def DisplayContextTooltips(self):
        return self.pivottable.DisplayContextTooltips

    @DisplayContextTooltips.setter
    def DisplayContextTooltips(self, value):
        self.pivottable.DisplayContextTooltips = value

    @property
    def DisplayEmptyColumn(self):
        return self.pivottable.DisplayEmptyColumn

    @DisplayEmptyColumn.setter
    def DisplayEmptyColumn(self, value):
        self.pivottable.DisplayEmptyColumn = value

    @property
    def DisplayEmptyRow(self):
        return self.pivottable.DisplayEmptyRow

    @DisplayEmptyRow.setter
    def DisplayEmptyRow(self, value):
        self.pivottable.DisplayEmptyRow = value

    @property
    def DisplayErrorString(self):
        return self.pivottable.DisplayErrorString

    @DisplayErrorString.setter
    def DisplayErrorString(self, value):
        self.pivottable.DisplayErrorString = value

    @property
    def DisplayFieldCaptions(self):
        return self.pivottable.DisplayFieldCaptions

    @DisplayFieldCaptions.setter
    def DisplayFieldCaptions(self, value):
        self.pivottable.DisplayFieldCaptions = value

    @property
    def DisplayImmediateItems(self):
        return self.pivottable.DisplayImmediateItems

    @DisplayImmediateItems.setter
    def DisplayImmediateItems(self, value):
        self.pivottable.DisplayImmediateItems = value

    @property
    def DisplayMemberPropertyTooltips(self):
        return self.pivottable.DisplayMemberPropertyTooltips

    @DisplayMemberPropertyTooltips.setter
    def DisplayMemberPropertyTooltips(self, value):
        self.pivottable.DisplayMemberPropertyTooltips = value

    @property
    def DisplayNullString(self):
        return self.pivottable.DisplayNullString

    @DisplayNullString.setter
    def DisplayNullString(self, value):
        self.pivottable.DisplayNullString = value

    @property
    def EnableDataValueEditing(self):
        return self.pivottable.EnableDataValueEditing

    @EnableDataValueEditing.setter
    def EnableDataValueEditing(self, value):
        self.pivottable.EnableDataValueEditing = value

    @property
    def EnableDrilldown(self):
        return self.pivottable.EnableDrilldown

    @EnableDrilldown.setter
    def EnableDrilldown(self, value):
        self.pivottable.EnableDrilldown = value

    @property
    def EnableFieldDialog(self):
        return self.pivottable.EnableFieldDialog

    @EnableFieldDialog.setter
    def EnableFieldDialog(self, value):
        self.pivottable.EnableFieldDialog = value

    @property
    def EnableFieldList(self):
        return self.pivottable.EnableFieldList

    @EnableFieldList.setter
    def EnableFieldList(self, value):
        self.pivottable.EnableFieldList = value

    @property
    def EnableWizard(self):
        return self.pivottable.EnableWizard

    @EnableWizard.setter
    def EnableWizard(self, value):
        self.pivottable.EnableWizard = value

    @property
    def EnableWriteback(self):
        return self.pivottable.EnableWriteback

    @EnableWriteback.setter
    def EnableWriteback(self, value):
        self.pivottable.EnableWriteback = value

    @property
    def ErrorString(self):
        return self.pivottable.ErrorString

    @ErrorString.setter
    def ErrorString(self, value):
        self.pivottable.ErrorString = value

    @property
    def FieldListSortAscending(self):
        return self.pivottable.FieldListSortAscending

    @FieldListSortAscending.setter
    def FieldListSortAscending(self, value):
        self.pivottable.FieldListSortAscending = value

    @property
    def GrandTotalName(self):
        return self.pivottable.GrandTotalName

    @GrandTotalName.setter
    def GrandTotalName(self, value):
        self.pivottable.GrandTotalName = value

    def HiddenFields(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.pivottable.HiddenFields):
            return PivotField(self.pivottable.HiddenFields(*params))
        else:
            return PivotField(self.pivottable.GetHiddenFields(*params))

    @property
    def InGridDropZones(self):
        return self.pivottable.InGridDropZones

    @InGridDropZones.setter
    def InGridDropZones(self, value):
        self.pivottable.InGridDropZones = value

    @property
    def InnerDetail(self):
        return self.pivottable.InnerDetail

    @InnerDetail.setter
    def InnerDetail(self, value):
        self.pivottable.InnerDetail = value

    @property
    def LayoutRowDefault(self):
        return self.pivottable.LayoutRowDefault

    @LayoutRowDefault.setter
    def LayoutRowDefault(self, value):
        self.pivottable.LayoutRowDefault = value

    @property
    def Location(self):
        return self.pivottable.Location

    @Location.setter
    def Location(self, value):
        self.pivottable.Location = value

    @property
    def ManualUpdate(self):
        return self.pivottable.ManualUpdate

    @ManualUpdate.setter
    def ManualUpdate(self, value):
        self.pivottable.ManualUpdate = value

    @property
    def MDX(self):
        return self.pivottable.MDX

    @property
    def MergeLabels(self):
        return self.pivottable.MergeLabels

    @MergeLabels.setter
    def MergeLabels(self, value):
        self.pivottable.MergeLabels = value

    @property
    def Name(self):
        return self.pivottable.Name

    @Name.setter
    def Name(self, value):
        self.pivottable.Name = value

    @property
    def NullString(self):
        return self.pivottable.NullString

    @NullString.setter
    def NullString(self, value):
        self.pivottable.NullString = value

    @property
    def PageFieldOrder(self):
        return XlOrder(self.pivottable.PageFieldOrder)

    @PageFieldOrder.setter
    def PageFieldOrder(self, value):
        self.pivottable.PageFieldOrder = value

    def PageFields(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.pivottable.PageFields):
            return PivotField(self.pivottable.PageFields(*params))
        else:
            return PivotField(self.pivottable.GetPageFields(*params))

    @property
    def PageFieldStyle(self):
        return self.pivottable.PageFieldStyle

    @PageFieldStyle.setter
    def PageFieldStyle(self, value):
        self.pivottable.PageFieldStyle = value

    @property
    def PageFieldWrapCount(self):
        return self.pivottable.PageFieldWrapCount

    @PageFieldWrapCount.setter
    def PageFieldWrapCount(self, value):
        self.pivottable.PageFieldWrapCount = value

    @property
    def PageRange(self):
        return Range(self.pivottable.PageRange)

    @property
    def PageRangeCells(self):
        return Range(self.pivottable.PageRangeCells)

    @property
    def Parent(self):
        return self.pivottable.Parent

    @property
    def PivotColumnAxis(self):
        return PivotAxis(self.pivottable.PivotColumnAxis)

    @property
    def PivotFormulas(self):
        return PivotFormulas(self.pivottable.PivotFormulas)

    @property
    def PivotRowAxis(self):
        return PivotAxis(self.pivottable.PivotRowAxis)

    @property
    def PivotSelection(self):
        return self.pivottable.PivotSelection

    @PivotSelection.setter
    def PivotSelection(self, value):
        self.pivottable.PivotSelection = value

    @property
    def PivotSelectionStandard(self):
        return self.pivottable.PivotSelectionStandard

    @PivotSelectionStandard.setter
    def PivotSelectionStandard(self, value):
        self.pivottable.PivotSelectionStandard = value

    @property
    def PreserveFormatting(self):
        return self.pivottable.PreserveFormatting

    @property
    def PrintDrillIndicators(self):
        return self.pivottable.PrintDrillIndicators

    @PrintDrillIndicators.setter
    def PrintDrillIndicators(self, value):
        self.pivottable.PrintDrillIndicators = value

    @property
    def PrintTitles(self):
        return self.pivottable.PrintTitles

    @PrintTitles.setter
    def PrintTitles(self, value):
        self.pivottable.PrintTitles = value

    @property
    def RefreshDate(self):
        return self.pivottable.RefreshDate

    @property
    def RefreshName(self):
        return self.pivottable.RefreshName

    @property
    def RepeatItemsOnEachPrintedPage(self):
        return self.pivottable.RepeatItemsOnEachPrintedPage

    @RepeatItemsOnEachPrintedPage.setter
    def RepeatItemsOnEachPrintedPage(self, value):
        self.pivottable.RepeatItemsOnEachPrintedPage = value

    def RowFields(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.pivottable.RowFields):
            return PivotField(self.pivottable.RowFields(*params))
        else:
            return PivotField(self.pivottable.GetRowFields(*params))

    @property
    def RowGrand(self):
        return self.pivottable.RowGrand

    @RowGrand.setter
    def RowGrand(self, value):
        self.pivottable.RowGrand = value

    @property
    def RowRange(self):
        return Range(self.pivottable.RowRange)

    @property
    def SaveData(self):
        return self.pivottable.SaveData

    @SaveData.setter
    def SaveData(self, value):
        self.pivottable.SaveData = value

    @property
    def SelectionMode(self):
        return XlPTSelectionMode(self.pivottable.SelectionMode)

    @SelectionMode.setter
    def SelectionMode(self, value):
        self.pivottable.SelectionMode = value

    @property
    def ShowDrillIndicators(self):
        return self.pivottable.ShowDrillIndicators

    @ShowDrillIndicators.setter
    def ShowDrillIndicators(self, value):
        self.pivottable.ShowDrillIndicators = value

    @property
    def ShowPageMultipleItemLabel(self):
        return self.pivottable.ShowPageMultipleItemLabel

    @ShowPageMultipleItemLabel.setter
    def ShowPageMultipleItemLabel(self, value):
        self.pivottable.ShowPageMultipleItemLabel = value

    @property
    def ShowTableStyleColumnHeaders(self):
        return self.pivottable.ShowTableStyleColumnHeaders

    @ShowTableStyleColumnHeaders.setter
    def ShowTableStyleColumnHeaders(self, value):
        self.pivottable.ShowTableStyleColumnHeaders = value

    @property
    def ShowTableStyleColumnStripes(self):
        return self.pivottable.ShowTableStyleColumnStripes

    @ShowTableStyleColumnStripes.setter
    def ShowTableStyleColumnStripes(self, value):
        self.pivottable.ShowTableStyleColumnStripes = value

    @property
    def ShowTableStyleRowHeaders(self):
        return self.pivottable.ShowTableStyleRowHeaders

    @ShowTableStyleRowHeaders.setter
    def ShowTableStyleRowHeaders(self, value):
        self.pivottable.ShowTableStyleRowHeaders = value

    @property
    def ShowTableStyleRowStripes(self):
        return self.pivottable.ShowTableStyleRowStripes

    @ShowTableStyleRowStripes.setter
    def ShowTableStyleRowStripes(self, value):
        self.pivottable.ShowTableStyleRowStripes = value

    @property
    def ShowValuesRow(self):
        return self.pivottable.ShowValuesRow

    @ShowValuesRow.setter
    def ShowValuesRow(self, value):
        self.pivottable.ShowValuesRow = value

    @property
    def Slicers(self):
        return Slicers(self.pivottable.Slicers)

    @property
    def SmallGrid(self):
        return self.pivottable.SmallGrid

    @SmallGrid.setter
    def SmallGrid(self, value):
        self.pivottable.SmallGrid = value

    @property
    def SortUsingCustomLists(self):
        return self.pivottable.SortUsingCustomLists

    @SortUsingCustomLists.setter
    def SortUsingCustomLists(self, value):
        self.pivottable.SortUsingCustomLists = value

    @property
    def SourceData(self):
        return self.pivottable.SourceData

    @SourceData.setter
    def SourceData(self, value):
        self.pivottable.SourceData = value

    @property
    def SubtotalHiddenPageItems(self):
        return self.pivottable.SubtotalHiddenPageItems

    @SubtotalHiddenPageItems.setter
    def SubtotalHiddenPageItems(self, value):
        self.pivottable.SubtotalHiddenPageItems = value

    @property
    def Summary(self):
        return self.pivottable.Summary

    @Summary.setter
    def Summary(self, value):
        self.pivottable.Summary = value

    @property
    def TableRange1(self):
        return Range(self.pivottable.TableRange1)

    @property
    def TableRange2(self):
        return Range(self.pivottable.TableRange2)

    @property
    def TableStyle2(self):
        return self.pivottable.TableStyle2

    @TableStyle2.setter
    def TableStyle2(self, value):
        self.pivottable.TableStyle2 = value

    @property
    def Tag(self):
        return self.pivottable.Tag

    @Tag.setter
    def Tag(self, value):
        self.pivottable.Tag = value

    @property
    def TotalsAnnotation(self):
        return self.pivottable.TotalsAnnotation

    @TotalsAnnotation.setter
    def TotalsAnnotation(self, value):
        self.pivottable.TotalsAnnotation = value

    @property
    def VacatedStyle(self):
        return self.pivottable.VacatedStyle

    @VacatedStyle.setter
    def VacatedStyle(self, value):
        self.pivottable.VacatedStyle = value

    @property
    def Value(self):
        return self.pivottable.Value

    @Value.setter
    def Value(self, value):
        self.pivottable.Value = value

    @property
    def Version(self):
        return XlPivotTableVersionList(self.pivottable.Version)

    @property
    def ViewCalculatedMembers(self):
        return self.pivottable.ViewCalculatedMembers

    @ViewCalculatedMembers.setter
    def ViewCalculatedMembers(self, value):
        self.pivottable.ViewCalculatedMembers = value

    def VisibleFields(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.pivottable.VisibleFields):
            return PivotField(self.pivottable.VisibleFields(*params))
        else:
            return PivotField(self.pivottable.GetVisibleFields(*params))

    @property
    def VisualTotals(self):
        return self.pivottable.VisualTotals

    @VisualTotals.setter
    def VisualTotals(self, value):
        self.pivottable.VisualTotals = value

    @property
    def VisualTotalsForSets(self):
        return self.pivottable.VisualTotalsForSets

    @VisualTotalsForSets.setter
    def VisualTotalsForSets(self, value):
        self.pivottable.VisualTotalsForSets = value

    def AddDataField(self, Field=None, Caption=None, Function=None):
        params = [
            Field if Field is not None else pythoncom.Missing,
            Caption if Caption is not None else pythoncom.Missing,
            Function if Function is not None else pythoncom.Missing,
        ]
        return self.pivottable.AddDataField(*params)

    def AddFields(self, RowFields=None, ColumnFields=None, PageFields=None, AddToTable=None):
        params = [
            RowFields if RowFields is not None else pythoncom.Missing,
            ColumnFields if ColumnFields is not None else pythoncom.Missing,
            PageFields if PageFields is not None else pythoncom.Missing,
            AddToTable if AddToTable is not None else pythoncom.Missing,
        ]
        return self.pivottable.AddFields(*params)

    def AllocateChanges(self):
        return self.pivottable.AllocateChanges()

    def CalculatedFields(self):
        return self.pivottable.CalculatedFields()

    def ChangeConnection(self, conn=None):
        params = [
            conn if conn is not None else pythoncom.Missing,
        ]
        self.pivottable.ChangeConnection(*params)

    def ChangePivotCache(self, bstr=None):
        params = [
            bstr if bstr is not None else pythoncom.Missing,
        ]
        self.pivottable.ChangePivotCache(*params)

    def ClearAllFilters(self):
        self.pivottable.ClearAllFilters()

    def ClearTable(self):
        self.pivottable.ClearTable()

    def CommitChanges(self):
        return self.pivottable.CommitChanges()

    def ConvertToFormulas(self, ConvertFilters=None):
        params = [
            ConvertFilters if ConvertFilters is not None else pythoncom.Missing,
        ]
        self.pivottable.ConvertToFormulas(*params)

    def CreateCubeFile(self, File=None, Measures=None, Levels=None, Members=None, Properties=None):
        params = [
            File if File is not None else pythoncom.Missing,
            Measures if Measures is not None else pythoncom.Missing,
            Levels if Levels is not None else pythoncom.Missing,
            Members if Members is not None else pythoncom.Missing,
            Properties if Properties is not None else pythoncom.Missing,
        ]
        return self.pivottable.CreateCubeFile(*params)

    def DiscardChanges(self):
        return self.pivottable.DiscardChanges()

    def GetData(self, Name=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
        ]
        return self.pivottable.GetData(*params)

    def GetPivotData(self, DataField=None, Field1=None, Item1=None, Field2=None, Item2=None, Field3=None, Item3=None, Field4=None, Item4=None, Field5=None, Item5=None, Field6=None, Item6=None, Field7=None, Item7=None, Field8=None, Item8=None, Field9=None, Item9=None, Field10=None, Item10=None, Field11=None, Item11=None, Field12=None, Item12=None, Field13=None, Item13=None, Field14=None, Item14=None):
        params = [
            DataField if DataField is not None else pythoncom.Missing,
            Field1 if Field1 is not None else pythoncom.Missing,
            Item1 if Item1 is not None else pythoncom.Missing,
            Field2 if Field2 is not None else pythoncom.Missing,
            Item2 if Item2 is not None else pythoncom.Missing,
            Field3 if Field3 is not None else pythoncom.Missing,
            Item3 if Item3 is not None else pythoncom.Missing,
            Field4 if Field4 is not None else pythoncom.Missing,
            Item4 if Item4 is not None else pythoncom.Missing,
            Field5 if Field5 is not None else pythoncom.Missing,
            Item5 if Item5 is not None else pythoncom.Missing,
            Field6 if Field6 is not None else pythoncom.Missing,
            Item6 if Item6 is not None else pythoncom.Missing,
            Field7 if Field7 is not None else pythoncom.Missing,
            Item7 if Item7 is not None else pythoncom.Missing,
            Field8 if Field8 is not None else pythoncom.Missing,
            Item8 if Item8 is not None else pythoncom.Missing,
            Field9 if Field9 is not None else pythoncom.Missing,
            Item9 if Item9 is not None else pythoncom.Missing,
            Field10 if Field10 is not None else pythoncom.Missing,
            Item10 if Item10 is not None else pythoncom.Missing,
            Field11 if Field11 is not None else pythoncom.Missing,
            Item11 if Item11 is not None else pythoncom.Missing,
            Field12 if Field12 is not None else pythoncom.Missing,
            Item12 if Item12 is not None else pythoncom.Missing,
            Field13 if Field13 is not None else pythoncom.Missing,
            Item13 if Item13 is not None else pythoncom.Missing,
            Field14 if Field14 is not None else pythoncom.Missing,
            Item14 if Item14 is not None else pythoncom.Missing,
        ]
        return self.pivottable.GetPivotData(*params)

    def ListFormulas(self):
        self.pivottable.ListFormulas()

    def PivotCache(self):
        return self.pivottable.PivotCache()

    def PivotFields(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.pivottable.PivotFields(*params)

    def PivotSelect(self, Name=None, Mode=None, UseStandardName=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
            Mode if Mode is not None else pythoncom.Missing,
            UseStandardName if UseStandardName is not None else pythoncom.Missing,
        ]
        self.pivottable.PivotSelect(*params)

    def PivotTableWizard(self, SourceType=None, SourceData=None, TableDestination=None, TableName=None, RowGrand=None, ColumnGrand=None, SaveData=None, HasAutoFormat=None, AutoPage=None, Reserved=None, BackgroundQuery=None, OptimizeCache=None, PageFieldOrder=None, PageFieldWrapCount=None, ReadData=None, Connection=None):
        params = [
            SourceType if SourceType is not None else pythoncom.Missing,
            SourceData if SourceData is not None else pythoncom.Missing,
            TableDestination if TableDestination is not None else pythoncom.Missing,
            TableName if TableName is not None else pythoncom.Missing,
            RowGrand if RowGrand is not None else pythoncom.Missing,
            ColumnGrand if ColumnGrand is not None else pythoncom.Missing,
            SaveData if SaveData is not None else pythoncom.Missing,
            HasAutoFormat if HasAutoFormat is not None else pythoncom.Missing,
            AutoPage if AutoPage is not None else pythoncom.Missing,
            Reserved if Reserved is not None else pythoncom.Missing,
            BackgroundQuery if BackgroundQuery is not None else pythoncom.Missing,
            OptimizeCache if OptimizeCache is not None else pythoncom.Missing,
            PageFieldOrder if PageFieldOrder is not None else pythoncom.Missing,
            PageFieldWrapCount if PageFieldWrapCount is not None else pythoncom.Missing,
            ReadData if ReadData is not None else pythoncom.Missing,
            Connection if Connection is not None else pythoncom.Missing,
        ]
        self.pivottable.PivotTableWizard(*params)

    def RefreshDataSourceValues(self):
        return self.pivottable.RefreshDataSourceValues()

    def RefreshTable(self):
        return self.pivottable.RefreshTable()

    def RepeatAllLabels(self, Repeat=None):
        params = [
            Repeat if Repeat is not None else pythoncom.Missing,
        ]
        return self.pivottable.RepeatAllLabels(*params)

    def RowAxisLayout(self, RowLayout=None):
        params = [
            RowLayout if RowLayout is not None else pythoncom.Missing,
        ]
        self.pivottable.RowAxisLayout(*params)

    def ShowPages(self, PageField=None):
        params = [
            PageField if PageField is not None else pythoncom.Missing,
        ]
        return self.pivottable.ShowPages(*params)

    def SubtotalLocation(self, Location=None):
        params = [
            Location if Location is not None else pythoncom.Missing,
        ]
        self.pivottable.SubtotalLocation(*params)

    def Update(self):
        self.pivottable.Update()


class PivotTableChangeList:

    def __init__(self, pivottablechangelist=None):
        self.pivottablechangelist = pivottablechangelist

    @property
    def Application(self):
        return self.pivottablechangelist.Application

    @property
    def Count(self):
        return self.pivottablechangelist.Count

    @property
    def Creator(self):
        return self.pivottablechangelist.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.pivottablechangelist.Item):
            return ValueChange(self.pivottablechangelist.Item(*params))
        else:
            return ValueChange(self.pivottablechangelist.GetItem(*params))

    @property
    def Parent(self):
        return PivotTable(self.pivottablechangelist.Parent)

    def Add(self, Tuple=None, Value=None, AllocationValue=None, AllocationMethod=None, AllocationWeightExpression=None):
        params = [
            Tuple if Tuple is not None else pythoncom.Missing,
            Value if Value is not None else pythoncom.Missing,
            AllocationValue if AllocationValue is not None else pythoncom.Missing,
            AllocationMethod if AllocationMethod is not None else pythoncom.Missing,
            AllocationWeightExpression if AllocationWeightExpression is not None else pythoncom.Missing,
        ]
        return self.pivottablechangelist.Add(*params)


class PivotTables:

    def __init__(self, pivottables=None):
        self.pivottables = pivottables

    def __call__(self, item):
        return PivotTable(self.pivottables(item))

    @property
    def Application(self):
        return self.pivottables.Application

    @property
    def Count(self):
        return self.pivottables.Count

    @property
    def Creator(self):
        return self.pivottables.Creator

    @property
    def Parent(self):
        return self.pivottables.Parent

    def Add(self, PivotCache=None, TableDestination=None, TableName=None, ReadData=None, DefaultVersion=None):
        params = [
            PivotCache if PivotCache is not None else pythoncom.Missing,
            TableDestination if TableDestination is not None else pythoncom.Missing,
            TableName if TableName is not None else pythoncom.Missing,
            ReadData if ReadData is not None else pythoncom.Missing,
            DefaultVersion if DefaultVersion is not None else pythoncom.Missing,
        ]
        return PivotTable(self.pivottables.Add(*params))

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return PivotTable(self.pivottables.Item(*params))


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
        return XlChartElementPosition(self.plotarea.Position)

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
        return self.plotarea.ClearFormats()

    def Select(self):
        return self.plotarea.Select()


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
        return XlColorIndex(self.point.MarkerBackgroundColorIndex)

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
        return XlColorIndex(self.point.MarkerForegroundColorIndex)

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
        return XlMarkerStyle(self.point.MarkerStyle)

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
        return XlChartPictureType(self.point.PictureType)

    @PictureType.setter
    def PictureType(self, value):
        self.point.PictureType = value

    @property
    def PictureUnit2(self):
        return PictureType(self.point.PictureUnit2)

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
        params = [
            Type if Type is not None else pythoncom.Missing,
            LegendKey if LegendKey is not None else pythoncom.Missing,
            AutoText if AutoText is not None else pythoncom.Missing,
            HasLeaderLines if HasLeaderLines is not None else pythoncom.Missing,
            ShowSeriesName if ShowSeriesName is not None else pythoncom.Missing,
            ShowCategoryName if ShowCategoryName is not None else pythoncom.Missing,
            ShowValue if ShowValue is not None else pythoncom.Missing,
            ShowPercentage if ShowPercentage is not None else pythoncom.Missing,
            ShowBubbleSize if ShowBubbleSize is not None else pythoncom.Missing,
            Separator if Separator is not None else pythoncom.Missing,
        ]
        self.point.ApplyDataLabels(*params)

    def ClearFormats(self):
        return self.point.ClearFormats()

    def Copy(self):
        return self.point.Copy()

    def Delete(self):
        return self.point.Delete()

    def Paste(self):
        return self.point.Paste()

    def PieSliceLocation(self, loc=None, Index=None):
        params = [
            loc if loc is not None else pythoncom.Missing,
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.point.PieSliceLocation(*params)

    def Select(self):
        return self.point.Select()


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
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return Point(self.points.Item(*params))


class ProtectedViewWindow:

    def __init__(self, protectedviewwindow=None):
        self.protectedviewwindow = protectedviewwindow

    @property
    def Caption(self):
        return self.protectedviewwindow.Caption

    @Caption.setter
    def Caption(self, value):
        self.protectedviewwindow.Caption = value

    @property
    def EnableResize(self):
        return self.protectedviewwindow.EnableResize

    @EnableResize.setter
    def EnableResize(self, value):
        self.protectedviewwindow.EnableResize = value

    @property
    def Height(self):
        return self.protectedviewwindow.Height

    @Height.setter
    def Height(self, value):
        self.protectedviewwindow.Height = value

    @property
    def Left(self):
        return self.protectedviewwindow.Left

    @Left.setter
    def Left(self, value):
        self.protectedviewwindow.Left = value

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

    @property
    def Workbook(self):
        return self.protectedviewwindow.Workbook

    def Activate(self):
        return self.protectedviewwindow.Activate()

    def Close(self):
        return self.protectedviewwindow.Close()

    def Edit(self, WriteResPassword=None, UpdateLinks=None):
        params = [
            WriteResPassword if WriteResPassword is not None else pythoncom.Missing,
            UpdateLinks if UpdateLinks is not None else pythoncom.Missing,
        ]
        return Workbook(self.protectedviewwindow.Edit(*params))


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

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.protectedviewwindows.Item):
            return self.protectedviewwindows.Item(*params)
        else:
            return self.protectedviewwindows.GetItem(*params)

    @property
    def Parent(self):
        return self.protectedviewwindows.Parent

    def Open(self, FileName=None, Password=None, AddToMru=None, RepairMode=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            Password if Password is not None else pythoncom.Missing,
            AddToMru if AddToMru is not None else pythoncom.Missing,
            RepairMode if RepairMode is not None else pythoncom.Missing,
        ]
        return ProtectedViewWindow(self.protectedviewwindows.Open(*params))


class Protection:

    def __init__(self, protection=None):
        self.protection = protection

    @property
    def AllowDeletingColumns(self):
        return self.protection.AllowDeletingColumns

    @property
    def AllowDeletingRows(self):
        return self.protection.AllowDeletingRows

    @property
    def AllowEditRanges(self):
        return AllowEditRanges(self.protection.AllowEditRanges)

    @property
    def AllowFiltering(self):
        return self.protection.AllowFiltering

    @property
    def AllowFormattingCells(self):
        return self.protection.AllowFormattingCells

    @property
    def AllowFormattingColumns(self):
        return self.protection.AllowFormattingColumns

    @property
    def AllowFormattingRows(self):
        return self.protection.AllowFormattingRows

    @property
    def AllowInsertingColumns(self):
        return self.protection.AllowInsertingColumns

    @property
    def AllowInsertingHyperlinks(self):
        return self.protection.AllowInsertingHyperlinks

    @property
    def AllowInsertingRows(self):
        return self.protection.AllowInsertingRows

    @property
    def AllowSorting(self):
        return self.protection.AllowSorting

    @property
    def AllowUsingPivotTables(self):
        return self.protection.AllowUsingPivotTables


class PublishObject:

    def __init__(self, publishobject=None):
        self.publishobject = publishobject

    @property
    def Application(self):
        return self.publishobject.Application

    @property
    def AutoRepublish(self):
        return self.publishobject.AutoRepublish

    @AutoRepublish.setter
    def AutoRepublish(self, value):
        self.publishobject.AutoRepublish = value

    @property
    def Creator(self):
        return self.publishobject.Creator

    @property
    def DivID(self):
        return self.publishobject.DivID

    @property
    def Filename(self):
        return self.publishobject.Filename

    @Filename.setter
    def Filename(self, value):
        self.publishobject.Filename = value

    @property
    def HtmlType(self):
        return XlHtmlType(self.publishobject.HtmlType)

    @HtmlType.setter
    def HtmlType(self, value):
        self.publishobject.HtmlType = value

    @property
    def Parent(self):
        return self.publishobject.Parent

    @property
    def Sheet(self):
        return PublishObject(self.publishobject.Sheet)

    @property
    def Source(self):
        return self.publishobject.Source

    @property
    def SourceType(self):
        return XlSourceType(self.publishobject.SourceType)

    @property
    def Title(self):
        return self.publishobject.Title

    @Title.setter
    def Title(self, value):
        self.publishobject.Title = value

    def Delete(self):
        self.publishobject.Delete()

    def Publish(self, Create=None):
        params = [
            Create if Create is not None else pythoncom.Missing,
        ]
        self.publishobject.Publish(*params)


class PublishObjects:

    def __init__(self, publishobjects=None):
        self.publishobjects = publishobjects

    def __call__(self, item):
        return PublishObject(self.publishobjects(item))

    @property
    def Application(self):
        return self.publishobjects.Application

    @property
    def Count(self):
        return self.publishobjects.Count

    @property
    def Creator(self):
        return self.publishobjects.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.publishobjects.Item):
            return self.publishobjects.Item(*params)
        else:
            return self.publishobjects.GetItem(*params)

    @property
    def Parent(self):
        return self.publishobjects.Parent

    def Add(self, SourceType=None, FileName=None, Sheet=None, Source=None, HtmlType=None, DivID=None, Title=None):
        params = [
            SourceType if SourceType is not None else pythoncom.Missing,
            FileName if FileName is not None else pythoncom.Missing,
            Sheet if Sheet is not None else pythoncom.Missing,
            Source if Source is not None else pythoncom.Missing,
            HtmlType if HtmlType is not None else pythoncom.Missing,
            DivID if DivID is not None else pythoncom.Missing,
            Title if Title is not None else pythoncom.Missing,
        ]
        return PublishObject(self.publishobjects.Add(*params))

    def Delete(self):
        self.publishobjects.Delete()

    def Publish(self):
        self.publishobjects.Publish()


class QueryTable:

    def __init__(self, querytable=None):
        self.querytable = querytable

    @property
    def AdjustColumnWidth(self):
        return self.querytable.AdjustColumnWidth

    @AdjustColumnWidth.setter
    def AdjustColumnWidth(self, value):
        self.querytable.AdjustColumnWidth = value

    @property
    def Application(self):
        return self.querytable.Application

    @property
    def BackgroundQuery(self):
        return self.querytable.BackgroundQuery

    @BackgroundQuery.setter
    def BackgroundQuery(self, value):
        self.querytable.BackgroundQuery = value

    @property
    def CommandText(self):
        return self.querytable.CommandText

    @CommandText.setter
    def CommandText(self, value):
        self.querytable.CommandText = value

    @property
    def CommandType(self):
        return XlCmdType(self.querytable.CommandType)

    @CommandType.setter
    def CommandType(self, value):
        self.querytable.CommandType = value

    @property
    def Connection(self):
        return self.querytable.Connection

    @Connection.setter
    def Connection(self, value):
        self.querytable.Connection = value

    @property
    def Creator(self):
        return self.querytable.Creator

    @property
    def Destination(self):
        return Range(self.querytable.Destination)

    @property
    def EditWebPage(self):
        return self.querytable.EditWebPage

    @EditWebPage.setter
    def EditWebPage(self, value):
        self.querytable.EditWebPage = value

    @property
    def EnableEditing(self):
        return self.querytable.EnableEditing

    @EnableEditing.setter
    def EnableEditing(self, value):
        self.querytable.EnableEditing = value

    @property
    def EnableRefresh(self):
        return self.querytable.EnableRefresh

    @EnableRefresh.setter
    def EnableRefresh(self, value):
        self.querytable.EnableRefresh = value

    @property
    def FetchedRowOverflow(self):
        return self.querytable.FetchedRowOverflow

    @property
    def FieldNames(self):
        return self.querytable.FieldNames

    @FieldNames.setter
    def FieldNames(self, value):
        self.querytable.FieldNames = value

    @property
    def FillAdjacentFormulas(self):
        return self.querytable.FillAdjacentFormulas

    @FillAdjacentFormulas.setter
    def FillAdjacentFormulas(self, value):
        self.querytable.FillAdjacentFormulas = value

    @property
    def ListObject(self):
        return ListObject(self.querytable.ListObject)

    @property
    def MaintainConnection(self):
        return self.querytable.MaintainConnection

    @MaintainConnection.setter
    def MaintainConnection(self, value):
        self.querytable.MaintainConnection = value

    @property
    def Name(self):
        return self.querytable.Name

    @Name.setter
    def Name(self, value):
        self.querytable.Name = value

    @property
    def Parameters(self):
        return Parameters(self.querytable.Parameters)

    @property
    def Parent(self):
        return self.querytable.Parent

    @property
    def PostText(self):
        return self.querytable.PostText

    @PostText.setter
    def PostText(self, value):
        self.querytable.PostText = value

    @property
    def PreserveColumnInfo(self):
        return self.querytable.PreserveColumnInfo

    @PreserveColumnInfo.setter
    def PreserveColumnInfo(self, value):
        self.querytable.PreserveColumnInfo = value

    @property
    def PreserveFormatting(self):
        return self.querytable.PreserveFormatting

    @property
    def QueryType(self):
        return self.querytable.QueryType

    @property
    def Recordset(self):
        return self.querytable.Recordset

    @Recordset.setter
    def Recordset(self, value):
        self.querytable.Recordset = value

    @property
    def Refreshing(self):
        return self.querytable.Refreshing

    @property
    def RefreshOnFileOpen(self):
        return self.querytable.RefreshOnFileOpen

    @RefreshOnFileOpen.setter
    def RefreshOnFileOpen(self, value):
        self.querytable.RefreshOnFileOpen = value

    @property
    def RefreshPeriod(self):
        return self.querytable.RefreshPeriod

    @RefreshPeriod.setter
    def RefreshPeriod(self, value):
        self.querytable.RefreshPeriod = value

    @property
    def RefreshStyle(self):
        return XlCellInsertionMode(self.querytable.RefreshStyle)

    @RefreshStyle.setter
    def RefreshStyle(self, value):
        self.querytable.RefreshStyle = value

    @property
    def ResultRange(self):
        return Range(self.querytable.ResultRange)

    @property
    def RobustConnect(self):
        return XlRobustConnect(self.querytable.RobustConnect)

    @RobustConnect.setter
    def RobustConnect(self, value):
        self.querytable.RobustConnect = value

    @property
    def RowNumbers(self):
        return self.querytable.RowNumbers

    @RowNumbers.setter
    def RowNumbers(self, value):
        self.querytable.RowNumbers = value

    @property
    def SaveData(self):
        return self.querytable.SaveData

    @SaveData.setter
    def SaveData(self, value):
        self.querytable.SaveData = value

    @property
    def SavePassword(self):
        return self.querytable.SavePassword

    @SavePassword.setter
    def SavePassword(self, value):
        self.querytable.SavePassword = value

    @property
    def Sort(self):
        return self.querytable.Sort

    @property
    def SourceConnectionFile(self):
        return self.querytable.SourceConnectionFile

    @SourceConnectionFile.setter
    def SourceConnectionFile(self, value):
        self.querytable.SourceConnectionFile = value

    @property
    def SourceDataFile(self):
        return self.querytable.SourceDataFile

    @SourceDataFile.setter
    def SourceDataFile(self, value):
        self.querytable.SourceDataFile = value

    @property
    def TextFileColumnDataTypes(self):
        return self.querytable.TextFileColumnDataTypes

    @TextFileColumnDataTypes.setter
    def TextFileColumnDataTypes(self, value):
        self.querytable.TextFileColumnDataTypes = value

    @property
    def TextFileCommaDelimiter(self):
        return self.querytable.TextFileCommaDelimiter

    @TextFileCommaDelimiter.setter
    def TextFileCommaDelimiter(self, value):
        self.querytable.TextFileCommaDelimiter = value

    @property
    def TextFileConsecutiveDelimiter(self):
        return self.querytable.TextFileConsecutiveDelimiter

    @TextFileConsecutiveDelimiter.setter
    def TextFileConsecutiveDelimiter(self, value):
        self.querytable.TextFileConsecutiveDelimiter = value

    @property
    def TextFileDecimalSeparator(self):
        return self.querytable.TextFileDecimalSeparator

    @TextFileDecimalSeparator.setter
    def TextFileDecimalSeparator(self, value):
        self.querytable.TextFileDecimalSeparator = value

    @property
    def TextFileFixedColumnWidths(self):
        return self.querytable.TextFileFixedColumnWidths

    @TextFileFixedColumnWidths.setter
    def TextFileFixedColumnWidths(self, value):
        self.querytable.TextFileFixedColumnWidths = value

    @property
    def TextFileOtherDelimiter(self):
        return self.querytable.TextFileOtherDelimiter

    @TextFileOtherDelimiter.setter
    def TextFileOtherDelimiter(self, value):
        self.querytable.TextFileOtherDelimiter = value

    @property
    def TextFileParseType(self):
        return XlTextParsingType(self.querytable.TextFileParseType)

    @TextFileParseType.setter
    def TextFileParseType(self, value):
        self.querytable.TextFileParseType = value

    @property
    def TextFilePlatform(self):
        return XlPlatform(self.querytable.TextFilePlatform)

    @TextFilePlatform.setter
    def TextFilePlatform(self, value):
        self.querytable.TextFilePlatform = value

    @property
    def TextFilePromptOnRefresh(self):
        return self.querytable.TextFilePromptOnRefresh

    @TextFilePromptOnRefresh.setter
    def TextFilePromptOnRefresh(self, value):
        self.querytable.TextFilePromptOnRefresh = value

    @property
    def TextFileSemicolonDelimiter(self):
        return self.querytable.TextFileSemicolonDelimiter

    @TextFileSemicolonDelimiter.setter
    def TextFileSemicolonDelimiter(self, value):
        self.querytable.TextFileSemicolonDelimiter = value

    @property
    def TextFileSpaceDelimiter(self):
        return self.querytable.TextFileSpaceDelimiter

    @TextFileSpaceDelimiter.setter
    def TextFileSpaceDelimiter(self, value):
        self.querytable.TextFileSpaceDelimiter = value

    @property
    def TextFileStartRow(self):
        return self.querytable.TextFileStartRow

    @TextFileStartRow.setter
    def TextFileStartRow(self, value):
        self.querytable.TextFileStartRow = value

    @property
    def TextFileTabDelimiter(self):
        return self.querytable.TextFileTabDelimiter

    @TextFileTabDelimiter.setter
    def TextFileTabDelimiter(self, value):
        self.querytable.TextFileTabDelimiter = value

    @property
    def TextFileTextQualifier(self):
        return XlTextQualifier(self.querytable.TextFileTextQualifier)

    @TextFileTextQualifier.setter
    def TextFileTextQualifier(self, value):
        self.querytable.TextFileTextQualifier = value

    @property
    def TextFileThousandsSeparator(self):
        return self.querytable.TextFileThousandsSeparator

    @TextFileThousandsSeparator.setter
    def TextFileThousandsSeparator(self, value):
        self.querytable.TextFileThousandsSeparator = value

    @property
    def TextFileTrailingMinusNumbers(self):
        return self.querytable.TextFileTrailingMinusNumbers

    @TextFileTrailingMinusNumbers.setter
    def TextFileTrailingMinusNumbers(self, value):
        self.querytable.TextFileTrailingMinusNumbers = value

    @property
    def TextFileVisualLayout(self):
        return XlTextVisualLayoutType(self.querytable.TextFileVisualLayout)

    @TextFileVisualLayout.setter
    def TextFileVisualLayout(self, value):
        self.querytable.TextFileVisualLayout = value

    @property
    def WebConsecutiveDelimitersAsOne(self):
        return self.querytable.WebConsecutiveDelimitersAsOne

    @WebConsecutiveDelimitersAsOne.setter
    def WebConsecutiveDelimitersAsOne(self, value):
        self.querytable.WebConsecutiveDelimitersAsOne = value

    @property
    def WebDisableDateRecognition(self):
        return self.querytable.WebDisableDateRecognition

    @WebDisableDateRecognition.setter
    def WebDisableDateRecognition(self, value):
        self.querytable.WebDisableDateRecognition = value

    @property
    def WebDisableRedirections(self):
        return self.querytable.WebDisableRedirections

    @WebDisableRedirections.setter
    def WebDisableRedirections(self, value):
        self.querytable.WebDisableRedirections = value

    @property
    def WebFormatting(self):
        return XlWebFormatting(self.querytable.WebFormatting)

    @WebFormatting.setter
    def WebFormatting(self, value):
        self.querytable.WebFormatting = value

    @property
    def WebPreFormattedTextToColumns(self):
        return self.querytable.WebPreFormattedTextToColumns

    @WebPreFormattedTextToColumns.setter
    def WebPreFormattedTextToColumns(self, value):
        self.querytable.WebPreFormattedTextToColumns = value

    @property
    def WebSelectionType(self):
        return XlWebSelectionType(self.querytable.WebSelectionType)

    @WebSelectionType.setter
    def WebSelectionType(self, value):
        self.querytable.WebSelectionType = value

    @property
    def WebSingleBlockTextImport(self):
        return self.querytable.WebSingleBlockTextImport

    @WebSingleBlockTextImport.setter
    def WebSingleBlockTextImport(self, value):
        self.querytable.WebSingleBlockTextImport = value

    @property
    def WebTables(self):
        return self.querytable.WebTables

    @WebTables.setter
    def WebTables(self, value):
        self.querytable.WebTables = value

    @property
    def WorkbookConnection(self):
        return WorkbookConnection(self.querytable.WorkbookConnection)

    def CancelRefresh(self):
        self.querytable.CancelRefresh()

    def Delete(self):
        self.querytable.Delete()

    def Refresh(self, BackgroundQuery=None):
        params = [
            BackgroundQuery if BackgroundQuery is not None else pythoncom.Missing,
        ]
        return self.querytable.Refresh(*params)

    def ResetTimer(self):
        self.querytable.ResetTimer()

    def SaveAsODC(self, ODCFileName=None, Description=None, Keywords=None):
        params = [
            ODCFileName if ODCFileName is not None else pythoncom.Missing,
            Description if Description is not None else pythoncom.Missing,
            Keywords if Keywords is not None else pythoncom.Missing,
        ]
        self.querytable.SaveAsODC(*params)


class QueryTables:

    def __init__(self, querytables=None):
        self.querytables = querytables

    def __call__(self, item):
        return QueryTable(self.querytables(item))

    @property
    def Application(self):
        return self.querytables.Application

    @property
    def Count(self):
        return self.querytables.Count

    @property
    def Creator(self):
        return self.querytables.Creator

    @property
    def Parent(self):
        return self.querytables.Parent

    def Add(self, Connection=None, Destination=None, Sql=None):
        params = [
            Connection if Connection is not None else pythoncom.Missing,
            Destination if Destination is not None else pythoncom.Missing,
            Sql if Sql is not None else pythoncom.Missing,
        ]
        return QueryTable(self.querytables.Add(*params))

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return QueryTable(self.querytables.Item(*params))


class Range:

    def __init__(self, range=None):
        self.range = range

    @property
    def AddIndent(self):
        return self.range.AddIndent

    @AddIndent.setter
    def AddIndent(self, value):
        self.range.AddIndent = value

    def Address(self, RowAbsolute=None, ColumnAbsolute=None, ReferenceStyle=None, External=None, RelativeTo=None):
        params = [
            RowAbsolute if RowAbsolute is not None else pythoncom.Missing,
            ColumnAbsolute if ColumnAbsolute is not None else pythoncom.Missing,
            ReferenceStyle if ReferenceStyle is not None else pythoncom.Missing,
            External if External is not None else pythoncom.Missing,
            RelativeTo if RelativeTo is not None else pythoncom.Missing,
        ]
        if callable(self.range.Address):
            return self.range.Address(*params)
        else:
            return self.range.GetAddress(*params)

    def AddressLocal(self, RowAbsolute=None, ColumnAbsolute=None, ReferenceStyle=None, External=None, RelativeTo=None):
        params = [
            RowAbsolute if RowAbsolute is not None else pythoncom.Missing,
            ColumnAbsolute if ColumnAbsolute is not None else pythoncom.Missing,
            ReferenceStyle if ReferenceStyle is not None else pythoncom.Missing,
            External if External is not None else pythoncom.Missing,
            RelativeTo if RelativeTo is not None else pythoncom.Missing,
        ]
        if callable(self.range.AddressLocal):
            return self.range.AddressLocal(*params)
        else:
            return self.range.GetAddressLocal(*params)

    @property
    def AllowEdit(self):
        return self.range.AllowEdit

    @property
    def Application(self):
        return self.range.Application

    @property
    def Areas(self):
        return Areas(self.range.Areas)

    @property
    def Borders(self):
        return Borders(self.range.Borders)

    def Cells(self, RowIndex=None, ColumnIndex=None):
        params = [
            RowIndex if RowIndex is not None else pythoncom.Missing,
            ColumnIndex if ColumnIndex is not None else pythoncom.Missing,
        ]
        if callable(self.range.Cells):
            return Range(self.range.Cells(*params))
        else:
            return Range(self.range.GetCells(*params))

    def Characters(self, Start=None, Length=None):
        params = [
            Start if Start is not None else pythoncom.Missing,
            Length if Length is not None else pythoncom.Missing,
        ]
        if callable(self.range.Characters):
            return Characters(self.range.Characters(*params))
        else:
            return Characters(self.range.GetCharacters(*params))

    @property
    def Column(self):
        return self.range.Column

    @property
    def Columns(self):
        return Range(self.range.Columns)

    @property
    def ColumnWidth(self):
        return self.range.ColumnWidth

    @ColumnWidth.setter
    def ColumnWidth(self, value):
        self.range.ColumnWidth = value

    @property
    def Comment(self):
        return Comment(self.range.Comment)

    @property
    def CommentThreaded(self):
        return CommentThreaded(self.range.CommentThreaded)

    @property
    def Count(self):
        return self.range.Count

    @property
    def CountLarge(self):
        return self.range.CountLarge

    @property
    def Creator(self):
        return self.range.Creator

    @property
    def CurrentArray(self):
        return self.range.CurrentArray

    @property
    def CurrentRegion(self):
        return Range(self.range.CurrentRegion)

    @property
    def Dependents(self):
        return Range(self.range.Dependents)

    @property
    def DirectDependents(self):
        return Range(self.range.DirectDependents)

    @property
    def DirectPrecedents(self):
        return Range(self.range.DirectPrecedents)

    @property
    def DisplayFormat(self):
        return DisplayFormat(self.range.DisplayFormat)

    def End(self, Direction=None):
        params = [
            Direction if Direction is not None else pythoncom.Missing,
        ]
        if callable(self.range.End):
            return Range(self.range.End(*params))
        else:
            return Range(self.range.GetEnd(*params))

    @property
    def EntireColumn(self):
        return Range(self.range.EntireColumn)

    @property
    def EntireRow(self):
        return Range(self.range.EntireRow)

    @property
    def Errors(self):
        return self.range.Errors

    @property
    def Font(self):
        return Font(self.range.Font)

    @property
    def FormatConditions(self):
        return FormatConditions(self.range.FormatConditions)

    @property
    def Formula(self):
        return self.range.Formula

    @Formula.setter
    def Formula(self, value):
        self.range.Formula = value

    @property
    def Formula2(self):
        return self.range.Formula2

    @Formula2.setter
    def Formula2(self, value):
        self.range.Formula2 = value

    @property
    def Formula2Local(self):
        return self.range.Formula2Local

    @property
    def Formula2R1C1(self):
        return self.range.Formula2R1C1

    @property
    def Formula2R1C1Local(self):
        return self.range.Formula2R1C1Local

    @property
    def FormulaArray(self):
        return self.range.FormulaArray

    @FormulaArray.setter
    def FormulaArray(self, value):
        self.range.FormulaArray = value

    @property
    def FormulaHidden(self):
        return self.range.FormulaHidden

    @FormulaHidden.setter
    def FormulaHidden(self, value):
        self.range.FormulaHidden = value

    @property
    def FormulaLocal(self):
        return self.range.FormulaLocal

    @FormulaLocal.setter
    def FormulaLocal(self, value):
        self.range.FormulaLocal = value

    @property
    def FormulaR1C1(self):
        return self.range.FormulaR1C1

    @FormulaR1C1.setter
    def FormulaR1C1(self, value):
        self.range.FormulaR1C1 = value

    @property
    def FormulaR1C1Local(self):
        return self.range.FormulaR1C1Local

    @FormulaR1C1Local.setter
    def FormulaR1C1Local(self, value):
        self.range.FormulaR1C1Local = value

    @property
    def HasArray(self):
        return self.range.HasArray

    @property
    def HasFormula(self):
        return self.range.HasFormula

    @property
    def HasRichDataType(self):
        return self.range.HasRichDataType

    @property
    def HasSpill(self):
        return self.range.HasSpill

    @property
    def Height(self):
        return self.range.Height

    @property
    def Hidden(self):
        return self.range.Hidden

    @Hidden.setter
    def Hidden(self, value):
        self.range.Hidden = value

    @property
    def HorizontalAlignment(self):
        return self.range.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.range.HorizontalAlignment = value

    @property
    def Hyperlinks(self):
        return Hyperlinks(self.range.Hyperlinks)

    @property
    def ID(self):
        return self.range.ID

    @ID.setter
    def ID(self, value):
        self.range.ID = value

    @property
    def IndentLevel(self):
        return self.range.IndentLevel

    @IndentLevel.setter
    def IndentLevel(self, value):
        self.range.IndentLevel = value

    @property
    def Interior(self):
        return Interior(self.range.Interior)

    def Item(self, RowIndex=None, ColumnIndex=None):
        params = [
            RowIndex if RowIndex is not None else pythoncom.Missing,
            ColumnIndex if ColumnIndex is not None else pythoncom.Missing,
        ]
        if callable(self.range.Item):
            return Range(self.range.Item(*params))
        else:
            return Range(self.range.GetItem(*params))

    @property
    def Left(self):
        return self.range.Left

    @property
    def LinkedDataTypeState(self):
        return XlLinkedDataTypeState(self.range.LinkedDataTypeState)

    @property
    def ListHeaderRows(self):
        return self.range.ListHeaderRows

    @property
    def ListObject(self):
        return ListObject(self.range.ListObject)

    @property
    def LocationInTable(self):
        return PivotTable(self.range.LocationInTable)

    @property
    def Locked(self):
        return self.range.Locked

    @Locked.setter
    def Locked(self, value):
        self.range.Locked = value

    @property
    def MDX(self):
        return Range(self.range.MDX)

    @property
    def MergeArea(self):
        return Range(self.range.MergeArea)

    @property
    def MergeCells(self):
        return self.range.MergeCells

    @MergeCells.setter
    def MergeCells(self, value):
        self.range.MergeCells = value

    @property
    def Name(self):
        return self.range.Name

    @Name.setter
    def Name(self, value):
        self.range.Name = value

    @property
    def Next(self):
        return Range(self.range.Next)

    @property
    def NumberFormat(self):
        return self.range.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value):
        self.range.NumberFormat = value

    @property
    def NumberFormatLocal(self):
        return self.range.NumberFormatLocal

    @NumberFormatLocal.setter
    def NumberFormatLocal(self, value):
        self.range.NumberFormatLocal = value

    def Offset(self, RowOffset=None, ColumnOffset=None):
        params = [
            RowOffset if RowOffset is not None else pythoncom.Missing,
            ColumnOffset if ColumnOffset is not None else pythoncom.Missing,
        ]
        if callable(self.range.Offset):
            return Range(self.range.Offset(*params))
        else:
            return Range(self.range.GetOffset(*params))

    @property
    def Orientation(self):
        return self.range.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.range.Orientation = value

    @property
    def OutlineLevel(self):
        return self.range.OutlineLevel

    @OutlineLevel.setter
    def OutlineLevel(self, value):
        self.range.OutlineLevel = value

    @property
    def PageBreak(self):
        return XlPageBreak(self.range.PageBreak)

    @PageBreak.setter
    def PageBreak(self, value):
        self.range.PageBreak = value

    @property
    def Parent(self):
        return self.range.Parent

    @property
    def Phonetic(self):
        return Phonetic(self.range.Phonetic)

    @property
    def Phonetics(self):
        return Phonetics(self.range.Phonetics)

    @property
    def PivotCell(self):
        return PivotCell(self.range.PivotCell)

    @property
    def PivotField(self):
        return PivotField(self.range.PivotField)

    @property
    def PivotItem(self):
        return PivotItem(self.range.PivotItem)

    @property
    def PivotTable(self):
        return PivotTable(self.range.PivotTable)

    @property
    def Precedents(self):
        return Range(self.range.Precedents)

    @property
    def PrefixCharacter(self):
        return self.range.PrefixCharacter

    @property
    def Previous(self):
        return Range(self.range.Previous)

    @property
    def QueryTable(self):
        return QueryTable(self.range.QueryTable)

    def Range(self, Cell1=None, Cell2=None):
        params = [
            Cell1 if Cell1 is not None else pythoncom.Missing,
            Cell2 if Cell2 is not None else pythoncom.Missing,
        ]
        if callable(self.range.Range):
            return Range(self.range.Range(*params))
        else:
            return Range(self.range.GetRange(*params))

    @property
    def ReadingOrder(self):
        return self.range.ReadingOrder

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.range.ReadingOrder = value

    def Resize(self, RowSize=None, ColumnSize=None):
        params = [
            RowSize if RowSize is not None else pythoncom.Missing,
            ColumnSize if ColumnSize is not None else pythoncom.Missing,
        ]
        if callable(self.range.Resize):
            return Range(self.range.Resize(*params))
        else:
            return Range(self.range.GetResize(*params))

    @property
    def Row(self):
        return self.range.Row

    @property
    def RowHeight(self):
        return self.range.RowHeight

    @RowHeight.setter
    def RowHeight(self, value):
        self.range.RowHeight = value

    @property
    def Rows(self):
        return Range(self.range.Rows)

    @property
    def SavedAsArray(self):
        return self.range.SavedAsArray

    @property
    def ServerActions(self):
        return self.range.ServerActions

    @property
    def ShowDetail(self):
        return self.range.ShowDetail

    @ShowDetail.setter
    def ShowDetail(self, value):
        self.range.ShowDetail = value

    @property
    def ShrinkToFit(self):
        return self.range.ShrinkToFit

    @ShrinkToFit.setter
    def ShrinkToFit(self, value):
        self.range.ShrinkToFit = value

    @property
    def SoundNote(self):
        return self.range.SoundNote

    @property
    def SparklineGroups(self):
        return SparklineGroups(self.range.SparklineGroups)

    @property
    def SpillParent(self):
        return self.range.SpillParent

    @property
    def Style(self):
        return Style(self.range.Style)

    @Style.setter
    def Style(self, value):
        self.range.Style = value

    @property
    def Summary(self):
        return self.range.Summary

    @property
    def Text(self):
        return self.range.Text

    @property
    def Top(self):
        return self.range.Top

    @property
    def UseStandardHeight(self):
        return self.range.UseStandardHeight

    @UseStandardHeight.setter
    def UseStandardHeight(self, value):
        self.range.UseStandardHeight = value

    @property
    def UseStandardWidth(self):
        return self.range.UseStandardWidth

    @UseStandardWidth.setter
    def UseStandardWidth(self, value):
        self.range.UseStandardWidth = value

    @property
    def Validation(self):
        return Validation(self.range.Validation)

    @property
    def Value(self):
        return self.range.Value

    @Value.setter
    def Value(self, value):
        self.range.Value = value

    @property
    def Value2(self):
        return self.range.Value2

    @Value2.setter
    def Value2(self, value):
        self.range.Value2 = value

    @property
    def VerticalAlignment(self):
        return self.range.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.range.VerticalAlignment = value

    @property
    def Width(self):
        return self.range.Width

    @property
    def Worksheet(self):
        return Worksheet(self.range.Worksheet)

    @property
    def WrapText(self):
        return self.range.WrapText

    @WrapText.setter
    def WrapText(self, value):
        self.range.WrapText = value

    @property
    def XPath(self):
        return XPath(self.range.XPath)

    def Activate(self):
        return self.range.Activate()

    def AddComment(self, Text=None):
        params = [
            Text if Text is not None else pythoncom.Missing,
        ]
        return self.range.AddComment(*params)

    def AddCommentThreaded(self, Text=None):
        params = [
            Text if Text is not None else pythoncom.Missing,
        ]
        return self.range.AddCommentThreaded(*params)

    def AdvancedFilter(self, Action=None, CriteriaRange=None, CopyToRange=None, Unique=None):
        params = [
            Action if Action is not None else pythoncom.Missing,
            CriteriaRange if CriteriaRange is not None else pythoncom.Missing,
            CopyToRange if CopyToRange is not None else pythoncom.Missing,
            Unique if Unique is not None else pythoncom.Missing,
        ]
        return self.range.AdvancedFilter(*params)

    def AllocateChanges(self):
        self.range.AllocateChanges()

    def ApplyNames(self, Names=None, IgnoreRelativeAbsolute=None, UseRowColumnNames=None, OmitColumn=None, OmitRow=None, Order=None, AppendLast=None):
        params = [
            Names if Names is not None else pythoncom.Missing,
            IgnoreRelativeAbsolute if IgnoreRelativeAbsolute is not None else pythoncom.Missing,
            UseRowColumnNames if UseRowColumnNames is not None else pythoncom.Missing,
            OmitColumn if OmitColumn is not None else pythoncom.Missing,
            OmitRow if OmitRow is not None else pythoncom.Missing,
            Order if Order is not None else pythoncom.Missing,
            AppendLast if AppendLast is not None else pythoncom.Missing,
        ]
        return self.range.ApplyNames(*params)

    def ApplyOutlineStyles(self):
        return self.range.ApplyOutlineStyles()

    def AutoComplete(self, String=None):
        params = [
            String if String is not None else pythoncom.Missing,
        ]
        return self.range.AutoComplete(*params)

    def AutoFill(self, Destination=None, Type=None):
        params = [
            Destination if Destination is not None else pythoncom.Missing,
            Type if Type is not None else pythoncom.Missing,
        ]
        return self.range.AutoFill(*params)

    def AutoFilter(self, Field=None, Criteria1=None, Operator=None, Criteria2=None, SubField=None, VisibleDropDown=None):
        params = [
            Field if Field is not None else pythoncom.Missing,
            Criteria1 if Criteria1 is not None else pythoncom.Missing,
            Operator if Operator is not None else pythoncom.Missing,
            Criteria2 if Criteria2 is not None else pythoncom.Missing,
            SubField if SubField is not None else pythoncom.Missing,
            VisibleDropDown if VisibleDropDown is not None else pythoncom.Missing,
        ]
        return self.range.AutoFilter(*params)

    def AutoFit(self):
        return self.range.AutoFit()

    def AutoOutline(self):
        return self.range.AutoOutline()

    def BorderAround(self, LineStyle=None, Weight=None, ColorIndex=None, Color=None, ThemeColor=None):
        params = [
            LineStyle if LineStyle is not None else pythoncom.Missing,
            Weight if Weight is not None else pythoncom.Missing,
            ColorIndex if ColorIndex is not None else pythoncom.Missing,
            Color if Color is not None else pythoncom.Missing,
            ThemeColor if ThemeColor is not None else pythoncom.Missing,
        ]
        return self.range.BorderAround(*params)

    def Calculate(self):
        return self.range.Calculate()

    def CalculateRowMajorOrder(self):
        return self.range.CalculateRowMajorOrder()

    def CheckSpelling(self, CustomDictionary=None, IgnoreUppercase=None, AlwaysSuggest=None, SpellLang=None):
        params = [
            CustomDictionary if CustomDictionary is not None else pythoncom.Missing,
            IgnoreUppercase if IgnoreUppercase is not None else pythoncom.Missing,
            AlwaysSuggest if AlwaysSuggest is not None else pythoncom.Missing,
            SpellLang if SpellLang is not None else pythoncom.Missing,
        ]
        return self.range.CheckSpelling(*params)

    def Clear(self):
        return self.range.Clear()

    def ClearComments(self):
        self.range.ClearComments()

    def ClearContents(self):
        return self.range.ClearContents()

    def ClearFormats(self):
        return self.range.ClearFormats()

    def ClearHyperlinks(self):
        return self.range.ClearHyperlinks()

    def ClearNotes(self):
        return self.range.ClearNotes()

    def ClearOutline(self):
        return self.range.ClearOutline()

    def ColumnDifferences(self, Comparison=None):
        params = [
            Comparison if Comparison is not None else pythoncom.Missing,
        ]
        return self.range.ColumnDifferences(*params)

    def Consolidate(self, Sources=None, Function=None, TopRow=None, LeftColumn=None, CreateLinks=None):
        params = [
            Sources if Sources is not None else pythoncom.Missing,
            Function if Function is not None else pythoncom.Missing,
            TopRow if TopRow is not None else pythoncom.Missing,
            LeftColumn if LeftColumn is not None else pythoncom.Missing,
            CreateLinks if CreateLinks is not None else pythoncom.Missing,
        ]
        return self.range.Consolidate(*params)

    def ConvertToLinkedDataType(self, ServiceID=None, LanguageCulture=None):
        params = [
            ServiceID if ServiceID is not None else pythoncom.Missing,
            LanguageCulture if LanguageCulture is not None else pythoncom.Missing,
        ]
        self.range.ConvertToLinkedDataType(*params)

    def Copy(self, Destination=None):
        params = [
            Destination if Destination is not None else pythoncom.Missing,
        ]
        return self.range.Copy(*params)

    def CopyFromRecordset(self, Data=None, MaxRows=None, MaxColumns=None):
        params = [
            Data if Data is not None else pythoncom.Missing,
            MaxRows if MaxRows is not None else pythoncom.Missing,
            MaxColumns if MaxColumns is not None else pythoncom.Missing,
        ]
        return self.range.CopyFromRecordset(*params)

    def CopyPicture(self, Appearance=None, Format=None):
        params = [
            Appearance if Appearance is not None else pythoncom.Missing,
            Format if Format is not None else pythoncom.Missing,
        ]
        return self.range.CopyPicture(*params)

    def CreateNames(self, Top=None, Left=None, Bottom=None, Right=None):
        params = [
            Top if Top is not None else pythoncom.Missing,
            Left if Left is not None else pythoncom.Missing,
            Bottom if Bottom is not None else pythoncom.Missing,
            Right if Right is not None else pythoncom.Missing,
        ]
        return self.range.CreateNames(*params)

    def Cut(self, Destination=None):
        params = [
            Destination if Destination is not None else pythoncom.Missing,
        ]
        return self.range.Cut(*params)

    def DataSeries(self, Rowcol=None, Type=None, Date=None, Step=None, Stop=None, Trend=None):
        params = [
            Rowcol if Rowcol is not None else pythoncom.Missing,
            Type if Type is not None else pythoncom.Missing,
            Date if Date is not None else pythoncom.Missing,
            Step if Step is not None else pythoncom.Missing,
            Stop if Stop is not None else pythoncom.Missing,
            Trend if Trend is not None else pythoncom.Missing,
        ]
        return self.range.DataSeries(*params)

    def DataTypeToText(self):
        self.range.DataTypeToText()

    def Delete(self, Shift=None):
        params = [
            Shift if Shift is not None else pythoncom.Missing,
        ]
        return self.range.Delete(*params)

    def DialogBox(self):
        return self.range.DialogBox()

    def Dirty(self):
        self.range.Dirty()

    def DiscardChanges(self):
        self.range.DiscardChanges()

    def EditionOptions(self, Type=None, Option=None, Name=None, Reference=None, Appearance=None, ChartSize=None, Format=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            Option if Option is not None else pythoncom.Missing,
            Name if Name is not None else pythoncom.Missing,
            Reference if Reference is not None else pythoncom.Missing,
            Appearance if Appearance is not None else pythoncom.Missing,
            ChartSize if ChartSize is not None else pythoncom.Missing,
            Format if Format is not None else pythoncom.Missing,
        ]
        return self.range.EditionOptions(*params)

    def ExportAsFixedFormat(self, Type=None, FileName=None, Quality=None, IncludeDocProperties=None, IgnorePrintAreas=None, From=None, To=None, OpenAfterPublish=None, FixedFormatExtClassPtr=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            FileName if FileName is not None else pythoncom.Missing,
            Quality if Quality is not None else pythoncom.Missing,
            IncludeDocProperties if IncludeDocProperties is not None else pythoncom.Missing,
            IgnorePrintAreas if IgnorePrintAreas is not None else pythoncom.Missing,
            From if From is not None else pythoncom.Missing,
            To if To is not None else pythoncom.Missing,
            OpenAfterPublish if OpenAfterPublish is not None else pythoncom.Missing,
            FixedFormatExtClassPtr if FixedFormatExtClassPtr is not None else pythoncom.Missing,
        ]
        self.range.ExportAsFixedFormat(*params)

    def FillDown(self):
        return self.range.FillDown()

    def FillLeft(self):
        return self.range.FillLeft()

    def FillRight(self):
        return self.range.FillRight()

    def FillUp(self):
        return self.range.FillUp()

    def Find(self, What=None, After=None, LookIn=None, LookAt=None, SearchOrder=None, SearchDirection=None, MatchCase=None, MatchByte=None, SearchFormat=None):
        params = [
            What if What is not None else pythoncom.Missing,
            After if After is not None else pythoncom.Missing,
            LookIn if LookIn is not None else pythoncom.Missing,
            LookAt if LookAt is not None else pythoncom.Missing,
            SearchOrder if SearchOrder is not None else pythoncom.Missing,
            SearchDirection if SearchDirection is not None else pythoncom.Missing,
            MatchCase if MatchCase is not None else pythoncom.Missing,
            MatchByte if MatchByte is not None else pythoncom.Missing,
            SearchFormat if SearchFormat is not None else pythoncom.Missing,
        ]
        return self.range.Find(*params)

    def FindNext(self, After=None):
        params = [
            After if After is not None else pythoncom.Missing,
        ]
        return self.range.FindNext(*params)

    def FindPrevious(self, Before=None):
        params = [
            Before if Before is not None else pythoncom.Missing,
        ]
        return self.range.FindPrevious(*params)

    def FunctionWizard(self):
        return self.range.FunctionWizard()

    def Group(self, Start=None, End=None, By=None, Periods=None):
        params = [
            Start if Start is not None else pythoncom.Missing,
            End if End is not None else pythoncom.Missing,
            By if By is not None else pythoncom.Missing,
            Periods if Periods is not None else pythoncom.Missing,
        ]
        return self.range.Group(*params)

    def Insert(self, Shift=None, CopyOrigin=None):
        params = [
            Shift if Shift is not None else pythoncom.Missing,
            CopyOrigin if CopyOrigin is not None else pythoncom.Missing,
        ]
        return self.range.Insert(*params)

    def InsertIndent(self, InsertAmount=None):
        params = [
            InsertAmount if InsertAmount is not None else pythoncom.Missing,
        ]
        self.range.InsertIndent(*params)

    def Justify(self):
        return self.range.Justify()

    def ListNames(self):
        return self.range.ListNames()

    def Merge(self, Across=None):
        params = [
            Across if Across is not None else pythoncom.Missing,
        ]
        self.range.Merge(*params)

    def NavigateArrow(self, TowardPrecedent=None, ArrowNumber=None, LinkNumber=None):
        params = [
            TowardPrecedent if TowardPrecedent is not None else pythoncom.Missing,
            ArrowNumber if ArrowNumber is not None else pythoncom.Missing,
            LinkNumber if LinkNumber is not None else pythoncom.Missing,
        ]
        return self.range.NavigateArrow(*params)

    def NoteText(self, Text=None, Start=None, Length=None):
        params = [
            Text if Text is not None else pythoncom.Missing,
            Start if Start is not None else pythoncom.Missing,
            Length if Length is not None else pythoncom.Missing,
        ]
        return self.range.NoteText(*params)

    def Parse(self, ParseLine=None, Destination=None):
        params = [
            ParseLine if ParseLine is not None else pythoncom.Missing,
            Destination if Destination is not None else pythoncom.Missing,
        ]
        return self.range.Parse(*params)

    def PasteSpecial(self, Paste=None, Operation=None, SkipBlanks=None, Transpose=None):
        params = [
            Paste if Paste is not None else pythoncom.Missing,
            Operation if Operation is not None else pythoncom.Missing,
            SkipBlanks if SkipBlanks is not None else pythoncom.Missing,
            Transpose if Transpose is not None else pythoncom.Missing,
        ]
        return self.range.PasteSpecial(*params)

    def PrintOut(self, From=None, To=None, Copies=None, Preview=None, ActivePrinter=None, PrintToFile=None, Collate=None, PrToFileName=None):
        params = [
            From if From is not None else pythoncom.Missing,
            To if To is not None else pythoncom.Missing,
            Copies if Copies is not None else pythoncom.Missing,
            Preview if Preview is not None else pythoncom.Missing,
            ActivePrinter if ActivePrinter is not None else pythoncom.Missing,
            PrintToFile if PrintToFile is not None else pythoncom.Missing,
            Collate if Collate is not None else pythoncom.Missing,
            PrToFileName if PrToFileName is not None else pythoncom.Missing,
        ]
        return self.range.PrintOut(*params)

    def PrintPreview(self, EnableChanges=None):
        params = [
            EnableChanges if EnableChanges is not None else pythoncom.Missing,
        ]
        return self.range.PrintPreview(*params)

    def RemoveDuplicates(self, Columns=None, Header=None):
        params = [
            Columns if Columns is not None else pythoncom.Missing,
            Header if Header is not None else pythoncom.Missing,
        ]
        self.range.RemoveDuplicates(*params)

    def RemoveSubtotal(self):
        return self.range.RemoveSubtotal()

    def Replace(self, What=None, Replacement=None, LookAt=None, SearchOrder=None, MatchCase=None, MatchByte=None, SearchFormat=None, ReplaceFormat=None, FormulaVersion=None):
        params = [
            What if What is not None else pythoncom.Missing,
            Replacement if Replacement is not None else pythoncom.Missing,
            LookAt if LookAt is not None else pythoncom.Missing,
            SearchOrder if SearchOrder is not None else pythoncom.Missing,
            MatchCase if MatchCase is not None else pythoncom.Missing,
            MatchByte if MatchByte is not None else pythoncom.Missing,
            SearchFormat if SearchFormat is not None else pythoncom.Missing,
            ReplaceFormat if ReplaceFormat is not None else pythoncom.Missing,
            FormulaVersion if FormulaVersion is not None else pythoncom.Missing,
        ]
        return self.range.Replace(*params)

    def ResetContents(self):
        self.range.ResetContents()

    def RowDifferences(self, Comparison=None):
        params = [
            Comparison if Comparison is not None else pythoncom.Missing,
        ]
        return self.range.RowDifferences(*params)

    def Run(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.range.Run(*params)

    def Select(self):
        return self.range.Select()

    def SetCellDataTypeFromCell(self, Range=None, LanguageCulture=None):
        params = [
            Range if Range is not None else pythoncom.Missing,
            LanguageCulture if LanguageCulture is not None else pythoncom.Missing,
        ]
        self.range.SetCellDataTypeFromCell(*params)

    def SetPhonetic(self):
        self.range.SetPhonetic()

    def Show(self):
        return self.range.Show()

    def ShowCard(self):
        self.range.ShowCard()

    def ShowDependents(self, Remove=None):
        params = [
            Remove if Remove is not None else pythoncom.Missing,
        ]
        return self.range.ShowDependents(*params)

    def ShowErrors(self):
        return self.range.ShowErrors()

    def ShowPrecedents(self, Remove=None):
        params = [
            Remove if Remove is not None else pythoncom.Missing,
        ]
        return self.range.ShowPrecedents(*params)

    def Sort(self, Key1=None, Order1=None, Key2=None, Type=None, Order2=None, Key3=None, Order3=None, Header=None, OrderCustom=None, MatchCase=None, Orientation=None, SortMethod=None, DataOption1=None, DataOption2=None, DataOption3=None):
        params = [
            Key1 if Key1 is not None else pythoncom.Missing,
            Order1 if Order1 is not None else pythoncom.Missing,
            Key2 if Key2 is not None else pythoncom.Missing,
            Type if Type is not None else pythoncom.Missing,
            Order2 if Order2 is not None else pythoncom.Missing,
            Key3 if Key3 is not None else pythoncom.Missing,
            Order3 if Order3 is not None else pythoncom.Missing,
            Header if Header is not None else pythoncom.Missing,
            OrderCustom if OrderCustom is not None else pythoncom.Missing,
            MatchCase if MatchCase is not None else pythoncom.Missing,
            Orientation if Orientation is not None else pythoncom.Missing,
            SortMethod if SortMethod is not None else pythoncom.Missing,
            DataOption1 if DataOption1 is not None else pythoncom.Missing,
            DataOption2 if DataOption2 is not None else pythoncom.Missing,
            DataOption3 if DataOption3 is not None else pythoncom.Missing,
        ]
        return self.range.Sort(*params)

    def SortSpecial(self, SortMethod=None, Key1=None, Order1=None, Type=None, Key2=None, Order2=None, Key3=None, Order3=None, Header=None, OrderCustom=None, MatchCase=None, Orientation=None, DataOption1=None, DataOption2=None, DataOption3=None):
        params = [
            SortMethod if SortMethod is not None else pythoncom.Missing,
            Key1 if Key1 is not None else pythoncom.Missing,
            Order1 if Order1 is not None else pythoncom.Missing,
            Type if Type is not None else pythoncom.Missing,
            Key2 if Key2 is not None else pythoncom.Missing,
            Order2 if Order2 is not None else pythoncom.Missing,
            Key3 if Key3 is not None else pythoncom.Missing,
            Order3 if Order3 is not None else pythoncom.Missing,
            Header if Header is not None else pythoncom.Missing,
            OrderCustom if OrderCustom is not None else pythoncom.Missing,
            MatchCase if MatchCase is not None else pythoncom.Missing,
            Orientation if Orientation is not None else pythoncom.Missing,
            DataOption1 if DataOption1 is not None else pythoncom.Missing,
            DataOption2 if DataOption2 is not None else pythoncom.Missing,
            DataOption3 if DataOption3 is not None else pythoncom.Missing,
        ]
        return self.range.SortSpecial(*params)

    def Speak(self, SpeakDirection=None, SpeakFormulas=None):
        params = [
            SpeakDirection if SpeakDirection is not None else pythoncom.Missing,
            SpeakFormulas if SpeakFormulas is not None else pythoncom.Missing,
        ]
        self.range.Speak(*params)

    def SpecialCells(self, Type=None, Value=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            Value if Value is not None else pythoncom.Missing,
        ]
        return self.range.SpecialCells(*params)

    def SubscribeTo(self, Edition=None, Format=None):
        params = [
            Edition if Edition is not None else pythoncom.Missing,
            Format if Format is not None else pythoncom.Missing,
        ]
        return self.range.SubscribeTo(*params)

    def Subtotal(self, GroupBy=None, Function=None, TotalList=None, Replace=None, PageBreaks=None, SummaryBelowData=None):
        params = [
            GroupBy if GroupBy is not None else pythoncom.Missing,
            Function if Function is not None else pythoncom.Missing,
            TotalList if TotalList is not None else pythoncom.Missing,
            Replace if Replace is not None else pythoncom.Missing,
            PageBreaks if PageBreaks is not None else pythoncom.Missing,
            SummaryBelowData if SummaryBelowData is not None else pythoncom.Missing,
        ]
        return self.range.Subtotal(*params)

    def Table(self, RowInput=None, ColumnInput=None):
        params = [
            RowInput if RowInput is not None else pythoncom.Missing,
            ColumnInput if ColumnInput is not None else pythoncom.Missing,
        ]
        return self.range.Table(*params)

    def TextToColumns(self, Destination=None, DataType=None, TextQualifier=None, ConsecutiveDelimiter=None, Tab=None, Semicolon=None, Comma=None, Space=None, Other=None, OtherChar=None, FieldInfo=None, DecimalSeparator=None, ThousandsSeparator=None, TrailingMinusNumbers=None):
        params = [
            Destination if Destination is not None else pythoncom.Missing,
            DataType if DataType is not None else pythoncom.Missing,
            TextQualifier if TextQualifier is not None else pythoncom.Missing,
            ConsecutiveDelimiter if ConsecutiveDelimiter is not None else pythoncom.Missing,
            Tab if Tab is not None else pythoncom.Missing,
            Semicolon if Semicolon is not None else pythoncom.Missing,
            Comma if Comma is not None else pythoncom.Missing,
            Space if Space is not None else pythoncom.Missing,
            Other if Other is not None else pythoncom.Missing,
            OtherChar if OtherChar is not None else pythoncom.Missing,
            FieldInfo if FieldInfo is not None else pythoncom.Missing,
            DecimalSeparator if DecimalSeparator is not None else pythoncom.Missing,
            ThousandsSeparator if ThousandsSeparator is not None else pythoncom.Missing,
            TrailingMinusNumbers if TrailingMinusNumbers is not None else pythoncom.Missing,
        ]
        return self.range.TextToColumns(*params)

    def Ungroup(self):
        return self.range.Ungroup()

    def UnMerge(self):
        self.range.UnMerge()


class CellControl:

    def __init__(self, cellcontrol=None):
        self.cellcontrol = cellcontrol


class Ranges:

    def __init__(self, ranges=None):
        self.ranges = ranges

    def __call__(self, item):
        return Range(self.ranges(item))

    @property
    def Application(self):
        return self.ranges.Application

    @property
    def Count(self):
        return self.ranges.Count

    @property
    def Creator(self):
        return self.ranges.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.ranges.Item):
            return Range(self.ranges.Item(*params))
        else:
            return Range(self.ranges.GetItem(*params))

    @property
    def Parent(self):
        return self.ranges.Parent


class RecentFile:

    def __init__(self, recentfile=None):
        self.recentfile = recentfile

    @property
    def Application(self):
        return self.recentfile.Application

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

    def Delete(self):
        self.recentfile.Delete()

    def Open(self):
        return Workbook(self.recentfile.Open())


class RecentFiles:

    def __init__(self, recentfiles=None):
        self.recentfiles = recentfiles

    @property
    def Application(self):
        return self.recentfiles.Application

    @property
    def Count(self):
        return self.recentfiles.Count

    @property
    def Creator(self):
        return self.recentfiles.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.recentfiles.Item):
            return self.recentfiles.Item(*params)
        else:
            return self.recentfiles.GetItem(*params)

    @property
    def Maximum(self):
        return self.recentfiles.Maximum

    @Maximum.setter
    def Maximum(self, value):
        self.recentfiles.Maximum = value

    @property
    def Parent(self):
        return self.recentfiles.Parent

    def Add(self, Name=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
        ]
        return RecentFile(self.recentfiles.Add(*params))


class RectangularGradient:

    def __init__(self, rectangulargradient=None):
        self.rectangulargradient = rectangulargradient

    @property
    def Application(self):
        return self.rectangulargradient.Application

    @property
    def ColorStops(self):
        return ColorStops(self.rectangulargradient.ColorStops)

    @property
    def Creator(self):
        return self.rectangulargradient.Creator

    @property
    def Parent(self):
        return self.rectangulargradient.Parent

    @property
    def RectangleBottom(self):
        return self.rectangulargradient.RectangleBottom

    @RectangleBottom.setter
    def RectangleBottom(self, value):
        self.rectangulargradient.RectangleBottom = value

    @property
    def RectangleLeft(self):
        return self.rectangulargradient.RectangleLeft

    @RectangleLeft.setter
    def RectangleLeft(self, value):
        self.rectangulargradient.RectangleLeft = value

    @property
    def RectangleRight(self):
        return self.rectangulargradient.RectangleRight

    @RectangleRight.setter
    def RectangleRight(self, value):
        self.rectangulargradient.RectangleRight = value

    @property
    def RectangleTop(self):
        return self.rectangulargradient.RectangleTop

    @RectangleTop.setter
    def RectangleTop(self, value):
        self.rectangulargradient.RectangleTop = value


class Research:

    def __init__(self, research=None):
        self.research = research

    @property
    def Application(self):
        return self.research.Application

    @property
    def Creator(self):
        return self.research.Creator

    @property
    def Parent(self):
        return self.research.Parent

    def IsResearchService(self, ServiceID=None):
        params = [
            ServiceID if ServiceID is not None else pythoncom.Missing,
        ]
        return self.research.IsResearchService(*params)

    def Query(self, ServiceID=None, QueryString=None, QueryLanguage=None, UseSelection=None, RequeryContextXML=None, NewQueryContextXML=None, LaunchQuery=None):
        params = [
            ServiceID if ServiceID is not None else pythoncom.Missing,
            QueryString if QueryString is not None else pythoncom.Missing,
            QueryLanguage if QueryLanguage is not None else pythoncom.Missing,
            UseSelection if UseSelection is not None else pythoncom.Missing,
            RequeryContextXML if RequeryContextXML is not None else pythoncom.Missing,
            NewQueryContextXML if NewQueryContextXML is not None else pythoncom.Missing,
            LaunchQuery if LaunchQuery is not None else pythoncom.Missing,
        ]
        return self.research.Query(*params)

    def SetLanguagePair(self, LanguageFrom=None, LanguageTo=None):
        params = [
            LanguageFrom if LanguageFrom is not None else pythoncom.Missing,
            LanguageTo if LanguageTo is not None else pythoncom.Missing,
        ]
        return self.research.SetLanguagePair(*params)


class RoutingSlip:

    def __init__(self, routingslip=None):
        self.routingslip = routingslip

    @property
    def Application(self):
        return self.routingslip.Application

    @property
    def Creator(self):
        return self.routingslip.Creator

    @property
    def Delivery(self):
        return self.routingslip.Delivery

    @property
    def Message(self):
        return self.routingslip.Message

    @property
    def Parent(self):
        return self.routingslip.Parent

    @property
    def Recipients(self):
        return self.routingslip.Recipients

    @property
    def ReturnWhenDone(self):
        return self.routingslip.ReturnWhenDone

    @property
    def Status(self):
        return self.routingslip.Status

    @property
    def Subject(self):
        return self.routingslip.Subject

    @property
    def TrackStatus(self):
        return self.routingslip.TrackStatus

    def Reset(self):
        self.routingslip.Reset()


class RTD:

    def __init__(self, rtd=None):
        self.rtd = rtd

    @property
    def ThrottleInterval(self):
        return self.rtd.ThrottleInterval

    @ThrottleInterval.setter
    def ThrottleInterval(self, value):
        self.rtd.ThrottleInterval = value

    def RefreshData(self):
        self.rtd.RefreshData()

    def RestartServers(self):
        self.rtd.RestartServers()


class Scenario:

    def __init__(self, scenario=None):
        self.scenario = scenario

    @property
    def Application(self):
        return self.scenario.Application

    @property
    def ChangingCells(self):
        return Range(self.scenario.ChangingCells)

    @property
    def Comment(self):
        return self.scenario.Comment

    @Comment.setter
    def Comment(self, value):
        self.scenario.Comment = value

    @property
    def Creator(self):
        return self.scenario.Creator

    @property
    def Hidden(self):
        return self.scenario.Hidden

    @Hidden.setter
    def Hidden(self, value):
        self.scenario.Hidden = value

    @property
    def Index(self):
        return self.scenario.Index

    @property
    def Locked(self):
        return self.scenario.Locked

    @Locked.setter
    def Locked(self, value):
        self.scenario.Locked = value

    @property
    def Name(self):
        return self.scenario.Name

    @Name.setter
    def Name(self, value):
        self.scenario.Name = value

    @property
    def Parent(self):
        return self.scenario.Parent

    def Values(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.scenario.Values):
            return self.scenario.Values(*params)
        else:
            return self.scenario.GetValues(*params)

    def ChangeScenario(self, ChangingCells=None, Values=None):
        params = [
            ChangingCells if ChangingCells is not None else pythoncom.Missing,
            Values if Values is not None else pythoncom.Missing,
        ]
        return self.scenario.ChangeScenario(*params)

    def Delete(self):
        return self.scenario.Delete()

    def Show(self):
        return self.scenario.Show()


class Scenarios:

    def __init__(self, scenarios=None):
        self.scenarios = scenarios

    def __call__(self, item):
        return Scenario(self.scenarios(item))

    @property
    def Application(self):
        return self.scenarios.Application

    @property
    def Count(self):
        return self.scenarios.Count

    @property
    def Creator(self):
        return self.scenarios.Creator

    @property
    def Parent(self):
        return self.scenarios.Parent

    def Add(self, Name=None, ChangingCells=None, Values=None, Comment=None, Locked=None, Hidden=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
            ChangingCells if ChangingCells is not None else pythoncom.Missing,
            Values if Values is not None else pythoncom.Missing,
            Comment if Comment is not None else pythoncom.Missing,
            Locked if Locked is not None else pythoncom.Missing,
            Hidden if Hidden is not None else pythoncom.Missing,
        ]
        return Scenario(self.scenarios.Add(*params))

    def CreateSummary(self, ReportType=None, ResultCells=None):
        params = [
            ReportType if ReportType is not None else pythoncom.Missing,
            ResultCells if ResultCells is not None else pythoncom.Missing,
        ]
        return self.scenarios.CreateSummary(*params)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return Scenario(self.scenarios.Item(*params))

    def Merge(self, Source=None):
        params = [
            Source if Source is not None else pythoncom.Missing,
        ]
        return self.scenarios.Merge(*params)


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
        return XlBarShape(self.series.BarShape)

    @BarShape.setter
    def BarShape(self, value):
        self.series.BarShape = value

    @property
    def BubbleSizes(self):
        return self.series.BubbleSizes

    @BubbleSizes.setter
    def BubbleSizes(self, value):
        self.series.BubbleSizes = value

    @property
    def ChartType(self):
        return XlChartType(self.series.ChartType)

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
        return XlColorIndex(self.series.MarkerBackgroundColorIndex)

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
        return XlColorIndex(self.series.MarkerForegroundColorIndex)

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
        return XlMarkerStyle(self.series.MarkerStyle)

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
        return XlChartPictureType(self.series.PictureType)

    @PictureType.setter
    def PictureType(self, value):
        self.series.PictureType = value

    @property
    def PictureUnit2(self):
        return PictureType(self.series.PictureUnit2)

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
        params = [
            Type if Type is not None else pythoncom.Missing,
            LegendKey if LegendKey is not None else pythoncom.Missing,
            AutoText if AutoText is not None else pythoncom.Missing,
            HasLeaderLines if HasLeaderLines is not None else pythoncom.Missing,
            ShowSeriesName if ShowSeriesName is not None else pythoncom.Missing,
            ShowCategoryName if ShowCategoryName is not None else pythoncom.Missing,
            ShowValue if ShowValue is not None else pythoncom.Missing,
            ShowPercentage if ShowPercentage is not None else pythoncom.Missing,
            ShowBubbleSize if ShowBubbleSize is not None else pythoncom.Missing,
            Separator if Separator is not None else pythoncom.Missing,
        ]
        self.series.ApplyDataLabels(*params)

    def ClearFormats(self):
        return self.series.ClearFormats()

    def Copy(self):
        return self.series.Copy()

    def DataLabels(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.series.DataLabels(*params)

    def Delete(self):
        return self.series.Delete()

    def ErrorBar(self, Direction=None, Include=None, Type=None, Amount=None, MinusValues=None):
        params = [
            Direction if Direction is not None else pythoncom.Missing,
            Include if Include is not None else pythoncom.Missing,
            Type if Type is not None else pythoncom.Missing,
            Amount if Amount is not None else pythoncom.Missing,
            MinusValues if MinusValues is not None else pythoncom.Missing,
        ]
        return self.series.ErrorBar(*params)

    def GeoMappingLevel(self):
        self.series.GeoMappingLevel()

    def GeoProjectionType(self):
        self.series.GeoProjectionType()

    def Paste(self):
        return self.series.Paste()

    def Points(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.series.Points(*params)

    def RegionLabelOptions(self):
        self.series.RegionLabelOptions()

    def Select(self):
        return self.series.Select()

    def Trendlines(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.series.Trendlines(*params)


class SeriesCollection:

    def __init__(self, seriescollection=None):
        self.seriescollection = seriescollection

    def __call__(self, item):
        return SeriesCollectio(self.seriescollection(item))

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
        params = [
            Source if Source is not None else pythoncom.Missing,
            Rowcol if Rowcol is not None else pythoncom.Missing,
            SeriesLabels if SeriesLabels is not None else pythoncom.Missing,
            CategoryLabels if CategoryLabels is not None else pythoncom.Missing,
            Replace if Replace is not None else pythoncom.Missing,
        ]
        return Series(self.seriescollection.Add(*params))

    def Extend(self, Source=None, RowCol=None, CategoryLabels=None):
        params = [
            Source if Source is not None else pythoncom.Missing,
            RowCol if RowCol is not None else pythoncom.Missing,
            CategoryLabels if CategoryLabels is not None else pythoncom.Missing,
        ]
        return self.seriescollection.Extend(*params)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return Series(self.seriescollection.Item(*params))

    def NewSeries(self):
        return self.seriescollection.NewSeries()

    def Paste(self, RowCol=None, SeriesLabels=None, CategoryLabels=None, Replace=None, NewSeries=None):
        params = [
            RowCol if RowCol is not None else pythoncom.Missing,
            SeriesLabels if SeriesLabels is not None else pythoncom.Missing,
            CategoryLabels if CategoryLabels is not None else pythoncom.Missing,
            Replace if Replace is not None else pythoncom.Missing,
            NewSeries if NewSeries is not None else pythoncom.Missing,
        ]
        return self.seriescollection.Paste(*params)


class SeriesLines:

    def __init__(self, serieslines=None):
        self.serieslines = serieslines

    @property
    def Application(self):
        return self.serieslines.Application

    @property
    def Border(self):
        return Border(self.serieslines.Border)

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
        return self.serieslines.Delete()

    def Select(self):
        return self.serieslines.Select()


class ServerViewableItems:

    def __init__(self, serverviewableitems=None):
        self.serverviewableitems = serverviewableitems

    def __call__(self, item):
        return ServerViewableItem(self.serverviewableitems(item))

    @property
    def Application(self):
        return self.serverviewableitems.Application

    @property
    def Count(self):
        return self.serverviewableitems.Count

    @property
    def Creator(self):
        return self.serverviewableitems.Creator

    @property
    def Parent(self):
        return self.serverviewableitems.Parent

    def Add(self, Obj=None):
        params = [
            Obj if Obj is not None else pythoncom.Missing,
        ]
        return ServerViewableItem(self.serverviewableitems.Add(*params))

    def Delete(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        self.serverviewableitems.Delete(*params)

    def DeleteAll(self):
        self.serverviewableitems.DeleteAll()

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.serverviewableitems.Item(*params)


class ShadowFormat:

    def __init__(self, shadowformat=None):
        self.shadowformat = shadowformat

    @property
    def Application(self):
        return self.shadowformat.Application

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

    @Visible.setter
    def Visible(self, value):
        self.shadowformat.Visible = value

    def IncrementOffsetX(self, Increment=None):
        params = [
            Increment if Increment is not None else pythoncom.Missing,
        ]
        self.shadowformat.IncrementOffsetX(*params)

    def IncrementOffsetY(self, Increment=None):
        params = [
            Increment if Increment is not None else pythoncom.Missing,
        ]
        self.shadowformat.IncrementOffsetY(*params)


class Shape:

    def __init__(self, shape=None):
        self.shape = shape

    @property
    def Adjustments(self):
        return Adjustments(self.shape.Adjustments)

    @property
    def AlternativeText(self):
        return Shape(self.shape.AlternativeText)

    @AlternativeText.setter
    def AlternativeText(self, value):
        self.shape.AlternativeText = value

    @property
    def Application(self):
        return self.shape.Application

    @property
    def AutoShapeType(self):
        return ShapeRange(self.shape.AutoShapeType)

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
    def BlackWhiteMode(self):
        return self.shape.BlackWhiteMode

    @BlackWhiteMode.setter
    def BlackWhiteMode(self, value):
        self.shape.BlackWhiteMode = value

    @property
    def BottomRightCell(self):
        return Range(self.shape.BottomRightCell)

    @property
    def Callout(self):
        return CalloutFormat(self.shape.Callout)

    @property
    def Chart(self):
        return Chart(self.shape.Chart)

    @property
    def Child(self):
        return self.shape.Child

    @property
    def ConnectionSiteCount(self):
        return self.shape.ConnectionSiteCount

    @property
    def Connector(self):
        return self.shape.Connector

    @property
    def ConnectorFormat(self):
        return ConnectorFormat(self.shape.ConnectorFormat)

    @property
    def ControlFormat(self):
        return ControlFormat(self.shape.ControlFormat)

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
    def FormControlType(self):
        return XlFormControl(self.shape.FormControlType)

    @property
    def Glow(self):
        return self.shape.Glow

    @property
    def GraphicStyle(self):
        return self.shape.GraphicStyle

    @GraphicStyle.setter
    def GraphicStyle(self, value):
        self.shape.GraphicStyle = value

    @property
    def GroupItems(self):
        return GroupShapes(self.shape.GroupItems)

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
    def HorizontalFlip(self):
        return self.shape.HorizontalFlip

    @property
    def Hyperlink(self):
        return Hyperlink(self.shape.Hyperlink)

    @property
    def ID(self):
        return self.shape.ID

    @property
    def Left(self):
        return self.shape.Left

    @Left.setter
    def Left(self, value):
        self.shape.Left = value

    @property
    def Line(self):
        return LineFormat(self.shape.Line)

    @property
    def LinkFormat(self):
        return LinkFormat(self.shape.LinkFormat)

    @property
    def LockAspectRatio(self):
        return self.shape.LockAspectRatio

    @LockAspectRatio.setter
    def LockAspectRatio(self, value):
        self.shape.LockAspectRatio = value

    @property
    def Locked(self):
        return self.shape.Locked

    @Locked.setter
    def Locked(self, value):
        self.shape.Locked = value

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
        return ShapeNodes(self.shape.Nodes)

    @property
    def OLEFormat(self):
        return OLEFormat(self.shape.OLEFormat)

    @property
    def OnAction(self):
        return self.shape.OnAction

    @OnAction.setter
    def OnAction(self, value):
        self.shape.OnAction = value

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
    def Placement(self):
        return XlPlacement(self.shape.Placement)

    @Placement.setter
    def Placement(self, value):
        self.shape.Placement = value

    @property
    def Reflection(self):
        return self.shape.Reflection

    @property
    def Rotation(self):
        return self.shape.Rotation

    @Rotation.setter
    def Rotation(self, value):
        self.shape.Rotation = value

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
        return self.shape.SoftEdge

    @property
    def TextEffect(self):
        return TextEffectFormat(self.shape.TextEffect)

    @property
    def TextFrame(self):
        return TextFrame(self.shape.TextFrame)

    @property
    def TextFrame2(self):
        return TextFrame2(self.shape.TextFrame2)

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
    def TopLeftCell(self):
        return Range(self.shape.TopLeftCell)

    @property
    def Type(self):
        return self.shape.Type

    @Type.setter
    def Type(self, value):
        self.shape.Type = value

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
    def ZOrderPosition(self):
        return self.shape.ZOrderPosition

    def Apply(self):
        self.shape.Apply()

    def Copy(self):
        self.shape.Copy()

    def CopyPicture(self, Appearance=None, Format=None):
        params = [
            Appearance if Appearance is not None else pythoncom.Missing,
            Format if Format is not None else pythoncom.Missing,
        ]
        self.shape.CopyPicture(*params)

    def Cut(self):
        return self.shape.Cut()

    def Delete(self):
        self.shape.Delete()

    def Duplicate(self):
        return self.shape.Duplicate()

    def Flip(self, FlipCmd=None):
        params = [
            FlipCmd if FlipCmd is not None else pythoncom.Missing,
        ]
        self.shape.Flip(*params)

    def IncrementLeft(self, Increment=None):
        params = [
            Increment if Increment is not None else pythoncom.Missing,
        ]
        self.shape.IncrementLeft(*params)

    def IncrementRotation(self, Increment=None):
        params = [
            Increment if Increment is not None else pythoncom.Missing,
        ]
        self.shape.IncrementRotation(*params)

    def IncrementTop(self, Increment=None):
        params = [
            Increment if Increment is not None else pythoncom.Missing,
        ]
        self.shape.IncrementTop(*params)

    def PickUp(self):
        self.shape.PickUp()

    def RerouteConnections(self):
        self.shape.RerouteConnections()

    def ScaleHeight(self, Factor=None, RelativeToOriginalSize=None, Scale=None):
        params = [
            Factor if Factor is not None else pythoncom.Missing,
            RelativeToOriginalSize if RelativeToOriginalSize is not None else pythoncom.Missing,
            Scale if Scale is not None else pythoncom.Missing,
        ]
        self.shape.ScaleHeight(*params)

    def ScaleWidth(self, Factor=None, RelativeToOriginalSize=None, Scale=None):
        params = [
            Factor if Factor is not None else pythoncom.Missing,
            RelativeToOriginalSize if RelativeToOriginalSize is not None else pythoncom.Missing,
            Scale if Scale is not None else pythoncom.Missing,
        ]
        self.shape.ScaleWidth(*params)

    def Select(self, Replace=None):
        params = [
            Replace if Replace is not None else pythoncom.Missing,
        ]
        self.shape.Select(*params)

    def SetShapesDefaultProperties(self):
        self.shape.SetShapesDefaultProperties()

    def Ungroup(self):
        return ShapeRange(self.shape.Ungroup())

    def ZOrder(self, ZOrderCmd=None):
        params = [
            ZOrderCmd if ZOrderCmd is not None else pythoncom.Missing,
        ]
        self.shape.ZOrder(*params)


class ShapeNode:

    def __init__(self, shapenode=None):
        self.shapenode = shapenode

    @property
    def Application(self):
        return self.shapenode.Application

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


class ShapeNodes:

    def __init__(self, shapenodes=None):
        self.shapenodes = shapenodes

    def __call__(self, item):
        return ShapeNode(self.shapenodes(item))

    @property
    def Application(self):
        return self.shapenodes.Application

    @property
    def Count(self):
        return self.shapenodes.Count

    @property
    def Creator(self):
        return self.shapenodes.Creator

    @property
    def Parent(self):
        return self.shapenodes.Parent

    def Delete(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        self.shapenodes.Delete(*params)

    def Insert(self, Index=None, SegmentType=None, EditingType=None, X1=None, Y1=None, X2=None, Y2=None, X3=None, Y3=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
            SegmentType if SegmentType is not None else pythoncom.Missing,
            EditingType if EditingType is not None else pythoncom.Missing,
            X1 if X1 is not None else pythoncom.Missing,
            Y1 if Y1 is not None else pythoncom.Missing,
            X2 if X2 is not None else pythoncom.Missing,
            Y2 if Y2 is not None else pythoncom.Missing,
            X3 if X3 is not None else pythoncom.Missing,
            Y3 if Y3 is not None else pythoncom.Missing,
        ]
        self.shapenodes.Insert(*params)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return ShapeNode(self.shapenodes.Item(*params))

    def SetEditingType(self, Index=None, EditingType=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
            EditingType if EditingType is not None else pythoncom.Missing,
        ]
        self.shapenodes.SetEditingType(*params)

    def SetPosition(self, Index=None, X1=None, Y1=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
            X1 if X1 is not None else pythoncom.Missing,
            Y1 if Y1 is not None else pythoncom.Missing,
        ]
        self.shapenodes.SetPosition(*params)

    def SetSegmentType(self, Index=None, SegmentType=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
            SegmentType if SegmentType is not None else pythoncom.Missing,
        ]
        self.shapenodes.SetSegmentType(*params)


class Shapes:

    def __init__(self, shapes=None):
        self.shapes = shapes

    def __call__(self, item):
        return Shape(self.shapes(item))

    @property
    def Application(self):
        return self.shapes.Application

    @property
    def Count(self):
        return self.shapes.Count

    @property
    def Creator(self):
        return self.shapes.Creator

    @property
    def Parent(self):
        return self.shapes.Parent

    def Range(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.shapes.Range):
            return ShapeRange(self.shapes.Range(*params))
        else:
            return ShapeRange(self.shapes.GetRange(*params))

    def Add3DModel(self, FileName=None, LinkToFile=None, SaveWithDocument=None, Left=None, Top=None, Width=None, Height=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            LinkToFile if LinkToFile is not None else pythoncom.Missing,
            SaveWithDocument if SaveWithDocument is not None else pythoncom.Missing,
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            Width if Width is not None else pythoncom.Missing,
            Height if Height is not None else pythoncom.Missing,
        ]
        return self.shapes.Add3DModel(*params)

    def AddCallout(self, Type=None, Left=None, Top=None, Width=None, Height=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            Width if Width is not None else pythoncom.Missing,
            Height if Height is not None else pythoncom.Missing,
        ]
        return self.shapes.AddCallout(*params)

    def AddConnector(self, Type=None, BeginX=None, BeginY=None, EndX=None, EndY=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            BeginX if BeginX is not None else pythoncom.Missing,
            BeginY if BeginY is not None else pythoncom.Missing,
            EndX if EndX is not None else pythoncom.Missing,
            EndY if EndY is not None else pythoncom.Missing,
        ]
        return self.shapes.AddConnector(*params)

    def AddCurve(self, SafeArrayOfPoints=None):
        params = [
            SafeArrayOfPoints if SafeArrayOfPoints is not None else pythoncom.Missing,
        ]
        return self.shapes.AddCurve(*params)

    def AddFormControl(self, Type=None, Left=None, Top=None, Width=None, Height=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            Width if Width is not None else pythoncom.Missing,
            Height if Height is not None else pythoncom.Missing,
        ]
        return self.shapes.AddFormControl(*params)

    def AddLabel(self, Orientation=None, Left=None, Top=None, Width=None, Height=None):
        params = [
            Orientation if Orientation is not None else pythoncom.Missing,
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            Width if Width is not None else pythoncom.Missing,
            Height if Height is not None else pythoncom.Missing,
        ]
        return self.shapes.AddLabel(*params)

    def AddLine(self, BeginX=None, BeginY=None, EndX=None, EndY=None):
        params = [
            BeginX if BeginX is not None else pythoncom.Missing,
            BeginY if BeginY is not None else pythoncom.Missing,
            EndX if EndX is not None else pythoncom.Missing,
            EndY if EndY is not None else pythoncom.Missing,
        ]
        return self.shapes.AddLine(*params)

    def AddOLEObject(self, ClassType=None, FileName=None, Link=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Left=None, Top=None, Width=None, Height=None):
        params = [
            ClassType if ClassType is not None else pythoncom.Missing,
            FileName if FileName is not None else pythoncom.Missing,
            Link if Link is not None else pythoncom.Missing,
            DisplayAsIcon if DisplayAsIcon is not None else pythoncom.Missing,
            IconFileName if IconFileName is not None else pythoncom.Missing,
            IconIndex if IconIndex is not None else pythoncom.Missing,
            IconLabel if IconLabel is not None else pythoncom.Missing,
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            Width if Width is not None else pythoncom.Missing,
            Height if Height is not None else pythoncom.Missing,
        ]
        return self.shapes.AddOLEObject(*params)

    def AddPicture(self, FileName=None, LinkToFile=None, SaveWithDocument=None, Left=None, Top=None, Width=None, Height=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            LinkToFile if LinkToFile is not None else pythoncom.Missing,
            SaveWithDocument if SaveWithDocument is not None else pythoncom.Missing,
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            Width if Width is not None else pythoncom.Missing,
            Height if Height is not None else pythoncom.Missing,
        ]
        return self.shapes.AddPicture(*params)

    def AddPolyline(self, SafeArrayOfPoints=None):
        params = [
            SafeArrayOfPoints if SafeArrayOfPoints is not None else pythoncom.Missing,
        ]
        return self.shapes.AddPolyline(*params)

    def AddShape(self, Type=None, Left=None, Top=None, Width=None, Height=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            Width if Width is not None else pythoncom.Missing,
            Height if Height is not None else pythoncom.Missing,
        ]
        return self.shapes.AddShape(*params)

    def AddSmartArt(self, Layout=None, Left=None, Top=None, Width=None, Height=None):
        params = [
            Layout if Layout is not None else pythoncom.Missing,
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            Width if Width is not None else pythoncom.Missing,
            Height if Height is not None else pythoncom.Missing,
        ]
        return self.shapes.AddSmartArt(*params)

    def AddTextbox(self, Orientation=None, Left=None, Top=None, Width=None, Height=None):
        params = [
            Orientation if Orientation is not None else pythoncom.Missing,
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            Width if Width is not None else pythoncom.Missing,
            Height if Height is not None else pythoncom.Missing,
        ]
        return self.shapes.AddTextbox(*params)

    def AddTextEffect(self, PresetTextEffect=None, Text=None, FontName=None, FontSize=None, FontBold=None, FontItalic=None, Left=None, Top=None):
        params = [
            PresetTextEffect if PresetTextEffect is not None else pythoncom.Missing,
            Text if Text is not None else pythoncom.Missing,
            FontName if FontName is not None else pythoncom.Missing,
            FontSize if FontSize is not None else pythoncom.Missing,
            FontBold if FontBold is not None else pythoncom.Missing,
            FontItalic if FontItalic is not None else pythoncom.Missing,
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
        ]
        return self.shapes.AddTextEffect(*params)

    def BuildFreeform(self, EditingType=None, X1=None, Y1=None):
        params = [
            EditingType if EditingType is not None else pythoncom.Missing,
            X1 if X1 is not None else pythoncom.Missing,
            Y1 if Y1 is not None else pythoncom.Missing,
        ]
        return self.shapes.BuildFreeform(*params)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return Shape(self.shapes.Item(*params))

    def SelectAll(self):
        self.shapes.SelectAll()


class Sheets:

    def __init__(self, sheets=None):
        self.sheets = sheets

    def __call__(self, item):
        return Sheet(self.sheets(item))

    @property
    def Application(self):
        return self.sheets.Application

    @property
    def Count(self):
        return self.sheets.Count

    @property
    def Creator(self):
        return self.sheets.Creator

    @property
    def HPageBreaks(self):
        return HPageBreaks(self.sheets.HPageBreaks)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.sheets.Item):
            return self.sheets.Item(*params)
        else:
            return self.sheets.GetItem(*params)

    @property
    def Parent(self):
        return self.sheets.Parent

    @property
    def Visible(self):
        return self.sheets.Visible

    @Visible.setter
    def Visible(self, value):
        self.sheets.Visible = value

    @property
    def VPageBreaks(self):
        return VPageBreaks(self.sheets.VPageBreaks)

    def Add(self, Before=None, After=None, Count=None, Type=None):
        params = [
            Before if Before is not None else pythoncom.Missing,
            After if After is not None else pythoncom.Missing,
            Count if Count is not None else pythoncom.Missing,
            Type if Type is not None else pythoncom.Missing,
        ]
        return Sheet(self.sheets.Add(*params))

    def Copy(self, Before=None, After=None):
        params = [
            Before if Before is not None else pythoncom.Missing,
            After if After is not None else pythoncom.Missing,
        ]
        self.sheets.Copy(*params)

    def Delete(self):
        self.sheets.Delete()

    def FillAcrossSheets(self, Range=None, Type=None):
        params = [
            Range if Range is not None else pythoncom.Missing,
            Type if Type is not None else pythoncom.Missing,
        ]
        self.sheets.FillAcrossSheets(*params)

    def Move(self, Before=None, After=None):
        params = [
            Before if Before is not None else pythoncom.Missing,
            After if After is not None else pythoncom.Missing,
        ]
        self.sheets.Move(*params)

    def PrintOut(self, From=None, To=None, Copies=None, Preview=None, ActivePrinter=None, PrintToFile=None, Collate=None, PrToFileName=None, IgnorePrintAreas=None):
        params = [
            From if From is not None else pythoncom.Missing,
            To if To is not None else pythoncom.Missing,
            Copies if Copies is not None else pythoncom.Missing,
            Preview if Preview is not None else pythoncom.Missing,
            ActivePrinter if ActivePrinter is not None else pythoncom.Missing,
            PrintToFile if PrintToFile is not None else pythoncom.Missing,
            Collate if Collate is not None else pythoncom.Missing,
            PrToFileName if PrToFileName is not None else pythoncom.Missing,
            IgnorePrintAreas if IgnorePrintAreas is not None else pythoncom.Missing,
        ]
        return self.sheets.PrintOut(*params)

    def PrintPreview(self, EnableChanges=None):
        params = [
            EnableChanges if EnableChanges is not None else pythoncom.Missing,
        ]
        self.sheets.PrintPreview(*params)

    def Select(self, Replace=None):
        params = [
            Replace if Replace is not None else pythoncom.Missing,
        ]
        self.sheets.Select(*params)


class SheetViews:

    def __init__(self, sheetviews=None):
        self.sheetviews = sheetviews

    def __call__(self, item):
        return SheetView(self.sheetviews(item))

    @property
    def Application(self):
        return self.sheetviews.Application

    @property
    def Count(self):
        return self.sheetviews.Count

    @property
    def Creator(self):
        return self.sheetviews.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.sheetviews.Item):
            return self.sheetviews.Item(*params)
        else:
            return self.sheetviews.GetItem(*params)

    @property
    def Parent(self):
        return self.sheetviews.Parent


class Slicer:

    def __init__(self, slicer=None):
        self.slicer = slicer

    @property
    def ActiveItem(self):
        return SlicerItem(self.slicer.ActiveItem)

    @property
    def Application(self):
        return self.slicer.Application

    @property
    def Caption(self):
        return self.slicer.Caption

    @Caption.setter
    def Caption(self, value):
        self.slicer.Caption = value

    @property
    def ColumnWidth(self):
        return self.slicer.ColumnWidth

    @ColumnWidth.setter
    def ColumnWidth(self, value):
        self.slicer.ColumnWidth = value

    @property
    def Creator(self):
        return self.slicer.Creator

    @property
    def DisableMoveResizeUI(self):
        return self.slicer.DisableMoveResizeUI

    @DisableMoveResizeUI.setter
    def DisableMoveResizeUI(self, value):
        self.slicer.DisableMoveResizeUI = value

    @property
    def DisplayHeader(self):
        return self.slicer.DisplayHeader

    @DisplayHeader.setter
    def DisplayHeader(self, value):
        self.slicer.DisplayHeader = value

    @property
    def Height(self):
        return self.slicer.Height

    @Height.setter
    def Height(self, value):
        self.slicer.Height = value

    @property
    def Left(self):
        return self.slicer.Left

    @Left.setter
    def Left(self, value):
        self.slicer.Left = value

    @property
    def Locked(self):
        return self.slicer.Locked

    @Locked.setter
    def Locked(self, value):
        self.slicer.Locked = value

    @property
    def Name(self):
        return self.slicer.Name

    @Name.setter
    def Name(self, value):
        self.slicer.Name = value

    @property
    def NumberOfColumns(self):
        return self.slicer.NumberOfColumns

    @NumberOfColumns.setter
    def NumberOfColumns(self, value):
        self.slicer.NumberOfColumns = value

    @property
    def Parent(self):
        return Worksheet(self.slicer.Parent)

    @property
    def RowHeight(self):
        return self.slicer.RowHeight

    @RowHeight.setter
    def RowHeight(self, value):
        self.slicer.RowHeight = value

    @property
    def Shape(self):
        return Shape(self.slicer.Shape)

    @property
    def SlicerCache(self):
        return SlicerCache(self.slicer.SlicerCache)

    @property
    def SlicerCacheLevel(self):
        return SlicerCacheLevel(self.slicer.SlicerCacheLevel)

    @property
    def Style(self):
        return self.slicer.Style

    @Style.setter
    def Style(self, value):
        self.slicer.Style = value

    @property
    def Top(self):
        return self.slicer.Top

    @Top.setter
    def Top(self, value):
        self.slicer.Top = value

    @property
    def Width(self):
        return self.slicer.Width

    @Width.setter
    def Width(self, value):
        self.slicer.Width = value

    def Copy(self):
        self.slicer.Copy()

    def Cut(self):
        self.slicer.Cut()

    def Delete(self):
        return self.slicer.Delete()


class SlicerCache:

    def __init__(self, slicercache=None):
        self.slicercache = slicercache

    @property
    def Application(self):
        return self.slicercache.Application

    @property
    def Creator(self):
        return self.slicercache.Creator

    @property
    def CrossFilterType(self):
        return self.slicercache.CrossFilterType

    @CrossFilterType.setter
    def CrossFilterType(self, value):
        self.slicercache.CrossFilterType = value

    @property
    def Index(self):
        return SlicerCaches(self.slicercache.Index)

    @property
    def Name(self):
        return self.slicercache.Name

    @Name.setter
    def Name(self, value):
        self.slicercache.Name = value

    @property
    def OLAP(self):
        return self.slicercache.OLAP

    @property
    def Parent(self):
        return SlicerCaches(self.slicercache.Parent)

    @property
    def PivotTables(self):
        return SlicerPivotTables(self.slicercache.PivotTables)

    @property
    def ShowAllItems(self):
        return self.slicercache.ShowAllItems

    @ShowAllItems.setter
    def ShowAllItems(self, value):
        self.slicercache.ShowAllItems = value

    @property
    def SlicerCacheLevels(self):
        return SlicerCacheLevel(self.slicercache.SlicerCacheLevels)

    @property
    def SlicerItems(self):
        return SlicerItems(self.slicercache.SlicerItems)

    @property
    def Slicers(self):
        return Slicers(self.slicercache.Slicers)

    @property
    def SortItems(self):
        return XlSlicerSort(self.slicercache.SortItems)

    @SortItems.setter
    def SortItems(self, value):
        self.slicercache.SortItems = value

    @property
    def SortUsingCustomLists(self):
        return self.slicercache.SortUsingCustomLists

    @SortUsingCustomLists.setter
    def SortUsingCustomLists(self, value):
        self.slicercache.SortUsingCustomLists = value

    @property
    def SourceName(self):
        return self.slicercache.SourceName

    @property
    def SourceType(self):
        return self.slicercache.SourceType

    @property
    def VisibleSlicerItems(self):
        return SlicerItems(self.slicercache.VisibleSlicerItems)

    @property
    def VisibleSlicerItemsList(self):
        return self.slicercache.VisibleSlicerItemsList

    @VisibleSlicerItemsList.setter
    def VisibleSlicerItemsList(self, value):
        self.slicercache.VisibleSlicerItemsList = value

    @property
    def WorkbookConnection(self):
        return self.slicercache.WorkbookConnection

    @WorkbookConnection.setter
    def WorkbookConnection(self, value):
        self.slicercache.WorkbookConnection = value

    def Delete(self):
        self.slicercache.Delete()


class SlicerCacheLevel:

    def __init__(self, slicercachelevel=None):
        self.slicercachelevel = slicercachelevel

    @property
    def Application(self):
        return self.slicercachelevel.Application

    @property
    def Count(self):
        return SlicerItem(self.slicercachelevel.Count)

    @property
    def Creator(self):
        return self.slicercachelevel.Creator

    @property
    def CrossFilterType(self):
        return self.slicercachelevel.CrossFilterType

    @CrossFilterType.setter
    def CrossFilterType(self, value):
        self.slicercachelevel.CrossFilterType = value

    @property
    def Name(self):
        return self.slicercachelevel.Name

    @property
    def Ordinal(self):
        return SlicerCacheLevel(self.slicercachelevel.Ordinal)

    @property
    def Parent(self):
        return SlicerCache(self.slicercachelevel.Parent)

    @property
    def SlicerItems(self):
        return SlicerItems(self.slicercachelevel.SlicerItems)

    @property
    def SortItems(self):
        return self.slicercachelevel.SortItems

    @SortItems.setter
    def SortItems(self, value):
        self.slicercachelevel.SortItems = value

    @property
    def VisibleSlicerItemsList(self):
        return self.slicercachelevel.VisibleSlicerItemsList


class SlicerCacheLevels:

    def __init__(self, slicercachelevels=None):
        self.slicercachelevels = slicercachelevels

    @property
    def Application(self):
        return self.slicercachelevels.Application

    @property
    def Count(self):
        return SlicerCache(self.slicercachelevels.Count)

    @property
    def Creator(self):
        return self.slicercachelevels.Creator

    def Item(self, Level=None):
        params = [
            Level if Level is not None else pythoncom.Missing,
        ]
        if callable(self.slicercachelevels.Item):
            return SlicerCacheLevel(self.slicercachelevels.Item(*params))
        else:
            return SlicerCacheLevel(self.slicercachelevels.GetItem(*params))

    @property
    def Parent(self):
        return SlicerCache(self.slicercachelevels.Parent)


class SlicerCaches:

    def __init__(self, slicercaches=None):
        self.slicercaches = slicercaches

    @property
    def Application(self):
        return self.slicercaches.Application

    @property
    def Count(self):
        return self.slicercaches.Count

    @property
    def Creator(self):
        return self.slicercaches.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.slicercaches.Item):
            return SlicerCache(self.slicercaches.Item(*params))
        else:
            return SlicerCache(self.slicercaches.GetItem(*params))

    @property
    def Parent(self):
        return Workbook(self.slicercaches.Parent)

    def Add(self, Source=None, SourceField=None, Name=None, SlicerCacheType=None):
        params = [
            Source if Source is not None else pythoncom.Missing,
            SourceField if SourceField is not None else pythoncom.Missing,
            Name if Name is not None else pythoncom.Missing,
            SlicerCacheType if SlicerCacheType is not None else pythoncom.Missing,
        ]
        return self.slicercaches.Add(*params)


class SlicerItem:

    def __init__(self, sliceritem=None):
        self.sliceritem = sliceritem

    @property
    def Application(self):
        return self.sliceritem.Application

    @property
    def Caption(self):
        return self.sliceritem.Caption

    @property
    def Creator(self):
        return self.sliceritem.Creator

    @property
    def HasData(self):
        return self.sliceritem.HasData

    @property
    def Name(self):
        return self.sliceritem.Name

    @property
    def Parent(self):
        return SlicerCache(self.sliceritem.Parent)

    @property
    def Selected(self):
        return self.sliceritem.Selected

    @Selected.setter
    def Selected(self, value):
        self.sliceritem.Selected = value

    @property
    def SourceName(self):
        return self.sliceritem.SourceName

    @property
    def SourceNameStandard(self):
        return self.sliceritem.SourceNameStandard

    @property
    def Value(self):
        return self.sliceritem.Value


class SlicerItems:

    def __init__(self, sliceritems=None):
        self.sliceritems = sliceritems

    @property
    def Application(self):
        return self.sliceritems.Application

    @property
    def Count(self):
        return self.sliceritems.Count

    @property
    def Creator(self):
        return self.sliceritems.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.sliceritems.Item):
            return SlicerItem(self.sliceritems.Item(*params))
        else:
            return SlicerItem(self.sliceritems.GetItem(*params))

    @property
    def Parent(self):
        return SlicerCache(self.sliceritems.Parent)


class SlicerPivotTables:

    def __init__(self, slicerpivottables=None):
        self.slicerpivottables = slicerpivottables

    @property
    def Application(self):
        return self.slicerpivottables.Application

    @property
    def Count(self):
        return self.slicerpivottables.Count

    @property
    def Creator(self):
        return self.slicerpivottables.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.slicerpivottables.Item):
            return PivotTable(self.slicerpivottables.Item(*params))
        else:
            return PivotTable(self.slicerpivottables.GetItem(*params))

    @property
    def Parent(self):
        return SlicerCache(self.slicerpivottables.Parent)

    def AddPivotTable(self, PivotTable=None):
        params = [
            PivotTable if PivotTable is not None else pythoncom.Missing,
        ]
        return self.slicerpivottables.AddPivotTable(*params)

    def RemovePivotTable(self, PivotTable=None):
        params = [
            PivotTable if PivotTable is not None else pythoncom.Missing,
        ]
        return self.slicerpivottables.RemovePivotTable(*params)


class Slicers:

    def __init__(self, slicers=None):
        self.slicers = slicers

    def __call__(self, item):
        return Slicer(self.slicers(item))

    @property
    def Application(self):
        return self.slicers.Application

    @property
    def Count(self):
        return self.slicers.Count

    @property
    def Creator(self):
        return self.slicers.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.slicers.Item):
            return Slicer(self.slicers.Item(*params))
        else:
            return Slicer(self.slicers.GetItem(*params))

    @property
    def Parent(self):
        return SlicerCache(self.slicers.Parent)

    def Add(self, SlicerDestination=None, Level=None, Name=None, Caption=None, Top=None, Left=None, Width=None, Height=None):
        params = [
            SlicerDestination if SlicerDestination is not None else pythoncom.Missing,
            Level if Level is not None else pythoncom.Missing,
            Name if Name is not None else pythoncom.Missing,
            Caption if Caption is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            Left if Left is not None else pythoncom.Missing,
            Width if Width is not None else pythoncom.Missing,
            Height if Height is not None else pythoncom.Missing,
        ]
        return Slicer(self.slicers.Add(*params))


class Sort:

    def __init__(self, sort=None):
        self.sort = sort

    @property
    def Application(self):
        return self.sort.Application

    @property
    def Creator(self):
        return self.sort.Creator

    @property
    def Header(self):
        return self.sort.Header

    @Header.setter
    def Header(self, value):
        self.sort.Header = value

    @property
    def MatchCase(self):
        return self.sort.MatchCase

    @MatchCase.setter
    def MatchCase(self, value):
        self.sort.MatchCase = value

    @property
    def Orientation(self):
        return self.sort.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.sort.Orientation = value

    @property
    def Parent(self):
        return self.sort.Parent

    @property
    def Rng(self):
        return self.sort.Rng

    @property
    def SortFields(self):
        return SortFields(self.sort.SortFields)

    @property
    def SortMethod(self):
        return self.sort.SortMethod

    @SortMethod.setter
    def SortMethod(self, value):
        self.sort.SortMethod = value

    def Apply(self):
        self.sort.Apply()

    def SetRange(self, Rng=None):
        params = [
            Rng if Rng is not None else pythoncom.Missing,
        ]
        self.sort.SetRange(*params)


class SortField:

    def __init__(self, sortfield=None):
        self.sortfield = sortfield

    @property
    def Application(self):
        return self.sortfield.Application

    @property
    def Creator(self):
        return self.sortfield.Creator

    @property
    def CustomOrder(self):
        return self.sortfield.CustomOrder

    @CustomOrder.setter
    def CustomOrder(self, value):
        self.sortfield.CustomOrder = value

    @property
    def DataOption(self):
        return self.sortfield.DataOption

    @DataOption.setter
    def DataOption(self, value):
        self.sortfield.DataOption = value

    @property
    def Key(self):
        return self.sortfield.Key

    @property
    def Order(self):
        return self.sortfield.Order

    @Order.setter
    def Order(self, value):
        self.sortfield.Order = value

    @property
    def Parent(self):
        return self.sortfield.Parent

    @property
    def Priority(self):
        return self.sortfield.Priority

    @Priority.setter
    def Priority(self, value):
        self.sortfield.Priority = value

    @property
    def SortOn(self):
        return XlSortOn(self.sortfield.SortOn)

    @SortOn.setter
    def SortOn(self, value):
        self.sortfield.SortOn = value

    @property
    def SortOnValue(self):
        return SortField(self.sortfield.SortOnValue)

    def Delete(self):
        self.sortfield.Delete()

    def ModifyKey(self, Key=None):
        params = [
            Key if Key is not None else pythoncom.Missing,
        ]
        self.sortfield.ModifyKey(*params)

    def SetIcon(self, Icon=None):
        params = [
            Icon if Icon is not None else pythoncom.Missing,
        ]
        self.sortfield.SetIcon(*params)


class SortFields:

    def __init__(self, sortfields=None):
        self.sortfields = sortfields

    @property
    def Application(self):
        return self.sortfields.Application

    @property
    def Count(self):
        return self.sortfields.Count

    @property
    def Creator(self):
        return self.sortfields.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.sortfields.Item):
            return SortField(self.sortfields.Item(*params))
        else:
            return SortField(self.sortfields.GetItem(*params))

    @property
    def Parent(self):
        return self.sortfields.Parent

    def Add(self, Key=None, SortOn=None, Order=None, CustomOrder=None, DataOption=None):
        params = [
            Key if Key is not None else pythoncom.Missing,
            SortOn if SortOn is not None else pythoncom.Missing,
            Order if Order is not None else pythoncom.Missing,
            CustomOrder if CustomOrder is not None else pythoncom.Missing,
            DataOption if DataOption is not None else pythoncom.Missing,
        ]
        return self.sortfields.Add(*params)

    def Add2(self, Key=None, SortOn=None, Order=None, CustomOrder=None, DataOption=None, SubField=None):
        params = [
            Key if Key is not None else pythoncom.Missing,
            SortOn if SortOn is not None else pythoncom.Missing,
            Order if Order is not None else pythoncom.Missing,
            CustomOrder if CustomOrder is not None else pythoncom.Missing,
            DataOption if DataOption is not None else pythoncom.Missing,
            SubField if SubField is not None else pythoncom.Missing,
        ]
        return self.sortfields.Add2(*params)

    def Clear(self):
        self.sortfields.Clear()


class SparkAxes:

    def __init__(self, sparkaxes=None):
        self.sparkaxes = sparkaxes

    @property
    def Application(self):
        return self.sparkaxes.Application

    @property
    def Creator(self):
        return self.sparkaxes.Creator

    @property
    def Horizontal(self):
        return SparkHorizontalAxis(self.sparkaxes.Horizontal)

    @property
    def Parent(self):
        return SparklineGroup(self.sparkaxes.Parent)

    @property
    def Vertical(self):
        return SparkVerticalAxis(self.sparkaxes.Vertical)


class SparkColor:

    def __init__(self, sparkcolor=None):
        self.sparkcolor = sparkcolor

    @property
    def Application(self):
        return self.sparkcolor.Application

    @property
    def Color(self):
        return FormatColor(self.sparkcolor.Color)

    @property
    def Creator(self):
        return self.sparkcolor.Creator

    @property
    def Parent(self):
        return SparklineGroup(self.sparkcolor.Parent)

    @property
    def Visible(self):
        return self.sparkcolor.Visible

    @Visible.setter
    def Visible(self, value):
        self.sparkcolor.Visible = value


class SparkHorizontalAxis:

    def __init__(self, sparkhorizontalaxis=None):
        self.sparkhorizontalaxis = sparkhorizontalaxis

    @property
    def Application(self):
        return self.sparkhorizontalaxis.Application

    @property
    def Axis(self):
        return SparkColor(self.sparkhorizontalaxis.Axis)

    @property
    def Creator(self):
        return self.sparkhorizontalaxis.Creator

    @property
    def IsDateAxis(self):
        return self.sparkhorizontalaxis.IsDateAxis

    @property
    def Parent(self):
        return SparklineGroup(self.sparkhorizontalaxis.Parent)

    @property
    def RightToLeftPlotOrder(self):
        return self.sparkhorizontalaxis.RightToLeftPlotOrder

    @RightToLeftPlotOrder.setter
    def RightToLeftPlotOrder(self, value):
        self.sparkhorizontalaxis.RightToLeftPlotOrder = value


class Sparkline:

    def __init__(self, sparkline=None):
        self.sparkline = sparkline

    @property
    def Application(self):
        return self.sparkline.Application

    @property
    def Creator(self):
        return self.sparkline.Creator

    @property
    def Location(self):
        return self.sparkline.Location

    @Location.setter
    def Location(self, value):
        self.sparkline.Location = value

    @property
    def Parent(self):
        return SparklineGroup(self.sparkline.Parent)

    @property
    def SourceData(self):
        return self.sparkline.SourceData

    @SourceData.setter
    def SourceData(self, value):
        self.sparkline.SourceData = value

    def ModifyLocation(self, Range=None):
        params = [
            Range if Range is not None else pythoncom.Missing,
        ]
        return self.sparkline.ModifyLocation(*params)

    def ModifySourceData(self, Formula=None):
        params = [
            Formula if Formula is not None else pythoncom.Missing,
        ]
        return self.sparkline.ModifySourceData(*params)


class SparklineGroup:

    def __init__(self, sparklinegroup=None):
        self.sparklinegroup = sparklinegroup

    @property
    def Application(self):
        return self.sparklinegroup.Application

    @property
    def Axes(self):
        return SparkAxes(self.sparklinegroup.Axes)

    @property
    def Count(self):
        return self.sparklinegroup.Count

    @property
    def Creator(self):
        return self.sparklinegroup.Creator

    @property
    def DateRange(self):
        return self.sparklinegroup.DateRange

    @DateRange.setter
    def DateRange(self, value):
        self.sparklinegroup.DateRange = value

    @property
    def DisplayHidden(self):
        return self.sparklinegroup.DisplayHidden

    @DisplayHidden.setter
    def DisplayHidden(self, value):
        self.sparklinegroup.DisplayHidden = value

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.sparklinegroup.Item):
            return Sparkline(self.sparklinegroup.Item(*params))
        else:
            return Sparkline(self.sparklinegroup.GetItem(*params))

    @property
    def LineWeight(self):
        return self.sparklinegroup.LineWeight

    @LineWeight.setter
    def LineWeight(self, value):
        self.sparklinegroup.LineWeight = value

    @property
    def Location(self):
        return self.sparklinegroup.Location

    @Location.setter
    def Location(self, value):
        self.sparklinegroup.Location = value

    @property
    def Parent(self):
        return self.sparklinegroup.Parent

    @property
    def SeriesColor(self):
        return FormatColor(self.sparklinegroup.SeriesColor)

    @property
    def SourceData(self):
        return self.sparklinegroup.SourceData

    @SourceData.setter
    def SourceData(self, value):
        self.sparklinegroup.SourceData = value

    @property
    def Type(self):
        return self.sparklinegroup.Type

    @Type.setter
    def Type(self, value):
        self.sparklinegroup.Type = value

    def Delete(self):
        return self.sparklinegroup.Delete()

    def Modify(self, Location=None, SourceData=None):
        params = [
            Location if Location is not None else pythoncom.Missing,
            SourceData if SourceData is not None else pythoncom.Missing,
        ]
        return self.sparklinegroup.Modify(*params)

    def ModifyDateRange(self, DateRange=None):
        params = [
            DateRange if DateRange is not None else pythoncom.Missing,
        ]
        return self.sparklinegroup.ModifyDateRange(*params)

    def ModifyLocation(self, Location=None):
        params = [
            Location if Location is not None else pythoncom.Missing,
        ]
        return self.sparklinegroup.ModifyLocation(*params)

    def ModifySourceData(self, SourceData=None):
        params = [
            SourceData if SourceData is not None else pythoncom.Missing,
        ]
        return self.sparklinegroup.ModifySourceData(*params)


class SparklineGroups:

    def __init__(self, sparklinegroups=None):
        self.sparklinegroups = sparklinegroups

    @property
    def Application(self):
        return self.sparklinegroups.Application

    @property
    def Count(self):
        return Range(self.sparklinegroups.Count)

    @property
    def Creator(self):
        return self.sparklinegroups.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.sparklinegroups.Item):
            return SparklineGroup(self.sparklinegroups.Item(*params))
        else:
            return SparklineGroup(self.sparklinegroups.GetItem(*params))

    @property
    def Parent(self):
        return Range(self.sparklinegroups.Parent)

    def Add(self, Type=None, SourceData=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            SourceData if SourceData is not None else pythoncom.Missing,
        ]
        return self.sparklinegroups.Add(*params)

    def Clear(self):
        return self.sparklinegroups.Clear()

    def ClearGroups(self):
        return self.sparklinegroups.ClearGroups()

    def Group(self, Location=None):
        params = [
            Location if Location is not None else pythoncom.Missing,
        ]
        return self.sparklinegroups.Group(*params)

    def Ungroup(self):
        return self.sparklinegroups.Ungroup()


class SparkPoints:

    def __init__(self, sparkpoints=None):
        self.sparkpoints = sparkpoints

    @property
    def Application(self):
        return self.sparkpoints.Application

    @property
    def Creator(self):
        return self.sparkpoints.Creator

    @property
    def Firstpoint(self):
        return SparkColor(self.sparkpoints.Firstpoint)

    @property
    def Highpoint(self):
        return SparkColor(self.sparkpoints.Highpoint)

    @property
    def Lastpoint(self):
        return SparkColor(self.sparkpoints.Lastpoint)

    @property
    def Lowpoint(self):
        return SparkColor(self.sparkpoints.Lowpoint)

    @property
    def Markers(self):
        return SparkColor(self.sparkpoints.Markers)

    @property
    def Negative(self):
        return SparkColor(self.sparkpoints.Negative)

    @property
    def Parent(self):
        return SparklineGroup(self.sparkpoints.Parent)


class SparkVerticalAxis:

    def __init__(self, sparkverticalaxis=None):
        self.sparkverticalaxis = sparkverticalaxis

    @property
    def Application(self):
        return self.sparkverticalaxis.Application

    @property
    def Creator(self):
        return self.sparkverticalaxis.Creator

    @property
    def CustomMaxScaleValue(self):
        return self.sparkverticalaxis.CustomMaxScaleValue

    @CustomMaxScaleValue.setter
    def CustomMaxScaleValue(self, value):
        self.sparkverticalaxis.CustomMaxScaleValue = value

    @property
    def CustomMinScaleValue(self):
        return self.sparkverticalaxis.CustomMinScaleValue

    @CustomMinScaleValue.setter
    def CustomMinScaleValue(self, value):
        self.sparkverticalaxis.CustomMinScaleValue = value

    @property
    def MaxScaleType(self):
        return self.sparkverticalaxis.MaxScaleType

    @MaxScaleType.setter
    def MaxScaleType(self, value):
        self.sparkverticalaxis.MaxScaleType = value

    @property
    def MinScaleType(self):
        return self.sparkverticalaxis.MinScaleType

    @MinScaleType.setter
    def MinScaleType(self, value):
        self.sparkverticalaxis.MinScaleType = value

    @property
    def Parent(self):
        return SparklineGroup(self.sparkverticalaxis.Parent)


class Speech:

    def __init__(self, speech=None):
        self.speech = speech

    @property
    def Direction(self):
        return XlSpeakDirection(self.speech.Direction)

    @Direction.setter
    def Direction(self, value):
        self.speech.Direction = value

    @property
    def SpeakCellOnEnter(self):
        return self.speech.SpeakCellOnEnter

    @SpeakCellOnEnter.setter
    def SpeakCellOnEnter(self, value):
        self.speech.SpeakCellOnEnter = value

    def Speak(self, Text=None, SpeakAsync=None, SpeakXML=None, Purge=None):
        params = [
            Text if Text is not None else pythoncom.Missing,
            SpeakAsync if SpeakAsync is not None else pythoncom.Missing,
            SpeakXML if SpeakXML is not None else pythoncom.Missing,
            Purge if Purge is not None else pythoncom.Missing,
        ]
        self.speech.Speak(*params)


class SpellingOptions:

    def __init__(self, spellingoptions=None):
        self.spellingoptions = spellingoptions

    @property
    def ArabicModes(self):
        return XlArabicModes(self.spellingoptions.ArabicModes)

    @ArabicModes.setter
    def ArabicModes(self, value):
        self.spellingoptions.ArabicModes = value

    @property
    def ArabicStrictAlefHamza(self):
        return self.spellingoptions.ArabicStrictAlefHamza

    @ArabicStrictAlefHamza.setter
    def ArabicStrictAlefHamza(self, value):
        self.spellingoptions.ArabicStrictAlefHamza = value

    @property
    def ArabicStrictFinalYaa(self):
        return self.spellingoptions.ArabicStrictFinalYaa

    @ArabicStrictFinalYaa.setter
    def ArabicStrictFinalYaa(self, value):
        self.spellingoptions.ArabicStrictFinalYaa = value

    @property
    def ArabicStrictTaaMarboota(self):
        return self.spellingoptions.ArabicStrictTaaMarboota

    @ArabicStrictTaaMarboota.setter
    def ArabicStrictTaaMarboota(self, value):
        self.spellingoptions.ArabicStrictTaaMarboota = value

    @property
    def BrazilReform(self):
        return self.spellingoptions.BrazilReform

    @BrazilReform.setter
    def BrazilReform(self, value):
        self.spellingoptions.BrazilReform = value

    @property
    def DictLang(self):
        return self.spellingoptions.DictLang

    @DictLang.setter
    def DictLang(self, value):
        self.spellingoptions.DictLang = value

    @property
    def GermanPostReform(self):
        return self.spellingoptions.GermanPostReform

    @GermanPostReform.setter
    def GermanPostReform(self, value):
        self.spellingoptions.GermanPostReform = value

    @property
    def HebrewModes(self):
        return XlHebrewModes(self.spellingoptions.HebrewModes)

    @HebrewModes.setter
    def HebrewModes(self, value):
        self.spellingoptions.HebrewModes = value

    @property
    def IgnoreCaps(self):
        return self.spellingoptions.IgnoreCaps

    @IgnoreCaps.setter
    def IgnoreCaps(self, value):
        self.spellingoptions.IgnoreCaps = value

    @property
    def IgnoreFileNames(self):
        return self.spellingoptions.IgnoreFileNames

    @IgnoreFileNames.setter
    def IgnoreFileNames(self, value):
        self.spellingoptions.IgnoreFileNames = value

    @property
    def IgnoreMixedDigits(self):
        return self.spellingoptions.IgnoreMixedDigits

    @IgnoreMixedDigits.setter
    def IgnoreMixedDigits(self, value):
        self.spellingoptions.IgnoreMixedDigits = value

    @property
    def KoreanCombineAux(self):
        return self.spellingoptions.KoreanCombineAux

    @KoreanCombineAux.setter
    def KoreanCombineAux(self, value):
        self.spellingoptions.KoreanCombineAux = value

    @property
    def KoreanProcessCompound(self):
        return self.spellingoptions.KoreanProcessCompound

    @KoreanProcessCompound.setter
    def KoreanProcessCompound(self, value):
        self.spellingoptions.KoreanProcessCompound = value

    @property
    def KoreanUseAutoChangeList(self):
        return self.spellingoptions.KoreanUseAutoChangeList

    @KoreanUseAutoChangeList.setter
    def KoreanUseAutoChangeList(self, value):
        self.spellingoptions.KoreanUseAutoChangeList = value

    @property
    def PortugalReform(self):
        return self.spellingoptions.PortugalReform

    @PortugalReform.setter
    def PortugalReform(self, value):
        self.spellingoptions.PortugalReform = value

    @property
    def RussianStrictE(self):
        return self.spellingoptions.RussianStrictE

    @RussianStrictE.setter
    def RussianStrictE(self, value):
        self.spellingoptions.RussianStrictE = value

    @property
    def SpanishModes(self):
        return self.spellingoptions.SpanishModes

    @SpanishModes.setter
    def SpanishModes(self, value):
        self.spellingoptions.SpanishModes = value

    @property
    def SuggestMainOnly(self):
        return self.spellingoptions.SuggestMainOnly

    @SuggestMainOnly.setter
    def SuggestMainOnly(self, value):
        self.spellingoptions.SuggestMainOnly = value

    @property
    def UserDict(self):
        return self.spellingoptions.UserDict

    @UserDict.setter
    def UserDict(self, value):
        self.spellingoptions.UserDict = value


class Style:

    def __init__(self, style=None):
        self.style = style

    @property
    def AddIndent(self):
        return self.style.AddIndent

    @AddIndent.setter
    def AddIndent(self, value):
        self.style.AddIndent = value

    @property
    def Application(self):
        return self.style.Application

    @property
    def Borders(self):
        return Borders(self.style.Borders)

    @property
    def BuiltIn(self):
        return self.style.BuiltIn

    @property
    def Creator(self):
        return self.style.Creator

    @property
    def Font(self):
        return Font(self.style.Font)

    @property
    def FormulaHidden(self):
        return self.style.FormulaHidden

    @FormulaHidden.setter
    def FormulaHidden(self, value):
        self.style.FormulaHidden = value

    @property
    def HorizontalAlignment(self):
        return XlHAlign(self.style.HorizontalAlignment)

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.style.HorizontalAlignment = value

    @property
    def IncludeAlignment(self):
        return self.style.IncludeAlignment

    @IncludeAlignment.setter
    def IncludeAlignment(self, value):
        self.style.IncludeAlignment = value

    @property
    def IncludeBorder(self):
        return self.style.IncludeBorder

    @IncludeBorder.setter
    def IncludeBorder(self, value):
        self.style.IncludeBorder = value

    @property
    def IncludeFont(self):
        return self.style.IncludeFont

    @IncludeFont.setter
    def IncludeFont(self, value):
        self.style.IncludeFont = value

    @property
    def IncludeNumber(self):
        return self.style.IncludeNumber

    @IncludeNumber.setter
    def IncludeNumber(self, value):
        self.style.IncludeNumber = value

    @property
    def IncludePatterns(self):
        return self.style.IncludePatterns

    @IncludePatterns.setter
    def IncludePatterns(self, value):
        self.style.IncludePatterns = value

    @property
    def IncludeProtection(self):
        return self.style.IncludeProtection

    @IncludeProtection.setter
    def IncludeProtection(self, value):
        self.style.IncludeProtection = value

    @property
    def IndentLevel(self):
        return self.style.IndentLevel

    @IndentLevel.setter
    def IndentLevel(self, value):
        self.style.IndentLevel = value

    @property
    def Interior(self):
        return Interior(self.style.Interior)

    @property
    def Locked(self):
        return self.style.Locked

    @Locked.setter
    def Locked(self, value):
        self.style.Locked = value

    @property
    def MergeCells(self):
        return self.style.MergeCells

    @MergeCells.setter
    def MergeCells(self, value):
        self.style.MergeCells = value

    @property
    def Name(self):
        return self.style.Name

    @property
    def NameLocal(self):
        return self.style.NameLocal

    @NameLocal.setter
    def NameLocal(self, value):
        self.style.NameLocal = value

    @property
    def NumberFormat(self):
        return self.style.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value):
        self.style.NumberFormat = value

    @property
    def NumberFormatLocal(self):
        return self.style.NumberFormatLocal

    @NumberFormatLocal.setter
    def NumberFormatLocal(self, value):
        self.style.NumberFormatLocal = value

    @property
    def Orientation(self):
        return XlOrientation(self.style.Orientation)

    @Orientation.setter
    def Orientation(self, value):
        self.style.Orientation = value

    @property
    def Parent(self):
        return self.style.Parent

    @property
    def ReadingOrder(self):
        return self.style.ReadingOrder

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.style.ReadingOrder = value

    @property
    def ShrinkToFit(self):
        return self.style.ShrinkToFit

    @ShrinkToFit.setter
    def ShrinkToFit(self, value):
        self.style.ShrinkToFit = value

    @property
    def Value(self):
        return self.style.Value

    @property
    def VerticalAlignment(self):
        return XlVAlign(self.style.VerticalAlignment)

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.style.VerticalAlignment = value

    @property
    def WrapText(self):
        return self.style.WrapText

    @WrapText.setter
    def WrapText(self, value):
        self.style.WrapText = value

    def Delete(self):
        return self.style.Delete()


class Styles:

    def __init__(self, styles=None):
        self.styles = styles

    def __call__(self, item):
        return Style(self.styles(item))

    @property
    def Application(self):
        return self.styles.Application

    @property
    def Count(self):
        return self.styles.Count

    @property
    def Creator(self):
        return self.styles.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.styles.Item):
            return self.styles.Item(*params)
        else:
            return self.styles.GetItem(*params)

    @property
    def Parent(self):
        return self.styles.Parent

    def Add(self, Name=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
        ]
        return Style(self.styles.Add(*params))

    def Merge(self, Workbook=None):
        params = [
            Workbook if Workbook is not None else pythoncom.Missing,
        ]
        return self.styles.Merge(*params)


class Tab:

    def __init__(self, tab=None):
        self.tab = tab

    @property
    def Application(self):
        return self.tab.Application

    @property
    def Color(self):
        return RGB(self.tab.Color)

    @Color.setter
    def Color(self, value):
        self.tab.Color = value

    @property
    def ColorIndex(self):
        return self.tab.ColorIndex

    @ColorIndex.setter
    def ColorIndex(self, value):
        self.tab.ColorIndex = value

    @property
    def Creator(self):
        return self.tab.Creator

    @property
    def Parent(self):
        return self.tab.Parent

    @property
    def ThemeColor(self):
        return XlThemeColor(self.tab.ThemeColor)

    @ThemeColor.setter
    def ThemeColor(self, value):
        self.tab.ThemeColor = value

    @property
    def TintAndShade(self):
        return self.tab.TintAndShade

    @TintAndShade.setter
    def TintAndShade(self, value):
        self.tab.TintAndShade = value


class TableStyle:

    def __init__(self, tablestyle=None):
        self.tablestyle = tablestyle

    @property
    def Application(self):
        return self.tablestyle.Application

    @property
    def BuiltIn(self):
        return self.tablestyle.BuiltIn

    @property
    def Creator(self):
        return self.tablestyle.Creator

    @property
    def Name(self):
        return self.tablestyle.Name

    @property
    def NameLocal(self):
        return self.tablestyle.NameLocal

    @NameLocal.setter
    def NameLocal(self, value):
        self.tablestyle.NameLocal = value

    @property
    def Parent(self):
        return self.tablestyle.Parent

    @property
    def ShowAsAvailablePivotTableStyle(self):
        return self.tablestyle.ShowAsAvailablePivotTableStyle

    @ShowAsAvailablePivotTableStyle.setter
    def ShowAsAvailablePivotTableStyle(self, value):
        self.tablestyle.ShowAsAvailablePivotTableStyle = value

    @property
    def ShowAsAvailableSlicerStyle(self):
        return self.tablestyle.ShowAsAvailableSlicerStyle

    @ShowAsAvailableSlicerStyle.setter
    def ShowAsAvailableSlicerStyle(self, value):
        self.tablestyle.ShowAsAvailableSlicerStyle = value

    @property
    def ShowAsAvailableTableStyle(self):
        return self.tablestyle.ShowAsAvailableTableStyle

    @ShowAsAvailableTableStyle.setter
    def ShowAsAvailableTableStyle(self, value):
        self.tablestyle.ShowAsAvailableTableStyle = value

    @property
    def TableStyleElements(self):
        return TableStyleElements(self.tablestyle.TableStyleElements)

    def Delete(self):
        return self.tablestyle.Delete()

    def Duplicate(self, NewTableStyleName=None):
        params = [
            NewTableStyleName if NewTableStyleName is not None else pythoncom.Missing,
        ]
        return self.tablestyle.Duplicate(*params)


class TableStyleElement:

    def __init__(self, tablestyleelement=None):
        self.tablestyleelement = tablestyleelement

    @property
    def Application(self):
        return self.tablestyleelement.Application

    @property
    def Borders(self):
        return Borders(self.tablestyleelement.Borders)

    @property
    def Creator(self):
        return self.tablestyleelement.Creator

    @property
    def Font(self):
        return Font(self.tablestyleelement.Font)

    @property
    def HasFormat(self):
        return self.tablestyleelement.HasFormat

    @property
    def Interior(self):
        return Interior(self.tablestyleelement.Interior)

    @property
    def Parent(self):
        return self.tablestyleelement.Parent

    @property
    def StripeSize(self):
        return self.tablestyleelement.StripeSize

    @StripeSize.setter
    def StripeSize(self, value):
        self.tablestyleelement.StripeSize = value

    def Clear(self):
        self.tablestyleelement.Clear()


class TableStyleElements:

    def __init__(self, tablestyleelements=None):
        self.tablestyleelements = tablestyleelements

    @property
    def Application(self):
        return self.tablestyleelements.Application

    @property
    def Count(self):
        return self.tablestyleelements.Count

    @property
    def Creator(self):
        return self.tablestyleelements.Creator

    @property
    def Parent(self):
        return self.tablestyleelements.Parent

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return TableStyleElement(self.tablestyleelements.Item(*params))


class TableStyles:

    def __init__(self, tablestyles=None):
        self.tablestyles = tablestyles

    @property
    def Application(self):
        return self.tablestyles.Application

    @property
    def Count(self):
        return self.tablestyles.Count

    @property
    def Creator(self):
        return self.tablestyles.Creator

    @property
    def Parent(self):
        return self.tablestyles.Parent

    def Add(self, TableStyleName=None):
        params = [
            TableStyleName if TableStyleName is not None else pythoncom.Missing,
        ]
        return self.tablestyles.Add(*params)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.tablestyles.Item(*params)


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
        return self.texteffectformat.Application

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
        return self.textframe.Application

    @property
    def AutoMargins(self):
        return self.textframe.AutoMargins

    @AutoMargins.setter
    def AutoMargins(self, value):
        self.textframe.AutoMargins = value

    @property
    def AutoSize(self):
        return self.textframe.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.textframe.AutoSize = value

    @property
    def Creator(self):
        return self.textframe.Creator

    @property
    def HorizontalAlignment(self):
        return XlHAlign(self.textframe.HorizontalAlignment)

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.textframe.HorizontalAlignment = value

    @property
    def HorizontalOverflow(self):
        return self.textframe.HorizontalOverflow

    @HorizontalOverflow.setter
    def HorizontalOverflow(self, value):
        self.textframe.HorizontalOverflow = value

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
    def Orientation(self):
        return self.textframe.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.textframe.Orientation = value

    @property
    def Parent(self):
        return self.textframe.Parent

    @property
    def ReadingOrder(self):
        return self.textframe.ReadingOrder

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.textframe.ReadingOrder = value

    @property
    def VerticalAlignment(self):
        return XlVAlign(self.textframe.VerticalAlignment)

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.textframe.VerticalAlignment = value

    @property
    def VerticalOverflow(self):
        return self.textframe.VerticalOverflow

    @VerticalOverflow.setter
    def VerticalOverflow(self, value):
        self.textframe.VerticalOverflow = value

    def Characters(self, Start=None, Length=None):
        params = [
            Start if Start is not None else pythoncom.Missing,
            Length if Length is not None else pythoncom.Missing,
        ]
        return self.textframe.Characters(*params)


class TextFrame2:

    def __init__(self, textframe2=None):
        self.textframe2 = textframe2

    @property
    def Application(self):
        return self.textframe2.Application

    @property
    def AutoSize(self):
        return self.textframe2.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.textframe2.AutoSize = value

    @property
    def Column(self):
        return self.textframe2.Column

    @property
    def Creator(self):
        return self.textframe2.Creator

    @property
    def HasText(self):
        return self.textframe2.HasText

    @property
    def HorizontalAnchor(self):
        return self.textframe2.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.textframe2.HorizontalAnchor = value

    @property
    def MarginBottom(self):
        return self.textframe2.MarginBottom

    @MarginBottom.setter
    def MarginBottom(self, value):
        self.textframe2.MarginBottom = value

    @property
    def MarginLeft(self):
        return self.textframe2.MarginLeft

    @MarginLeft.setter
    def MarginLeft(self, value):
        self.textframe2.MarginLeft = value

    @property
    def MarginRight(self):
        return self.textframe2.MarginRight

    @MarginRight.setter
    def MarginRight(self, value):
        self.textframe2.MarginRight = value

    @property
    def MarginTop(self):
        return self.textframe2.MarginTop

    @MarginTop.setter
    def MarginTop(self, value):
        self.textframe2.MarginTop = value

    @property
    def NoTextRotation(self):
        return self.textframe2.NoTextRotation

    @NoTextRotation.setter
    def NoTextRotation(self, value):
        self.textframe2.NoTextRotation = value

    @property
    def Orientation(self):
        return self.textframe2.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.textframe2.Orientation = value

    @property
    def Parent(self):
        return self.textframe2.Parent

    @property
    def PathFormat(self):
        return self.textframe2.PathFormat

    @PathFormat.setter
    def PathFormat(self, value):
        self.textframe2.PathFormat = value

    @property
    def Ruler(self):
        return self.textframe2.Ruler

    @property
    def TextRange(self):
        return self.textframe2.TextRange

    @property
    def ThreeD(self):
        return ThreeDFormat(self.textframe2.ThreeD)

    @property
    def VerticalAnchor(self):
        return self.textframe2.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.textframe2.VerticalAnchor = value

    @property
    def WarpFormat(self):
        return self.textframe2.WarpFormat

    @WarpFormat.setter
    def WarpFormat(self, value):
        self.textframe2.WarpFormat = value

    @property
    def WordArtformat(self):
        return self.textframe2.WordArtformat

    @WordArtformat.setter
    def WordArtformat(self, value):
        self.textframe2.WordArtformat = value

    @property
    def WordWrap(self):
        return self.textframe2.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.textframe2.WordWrap = value

    def DeleteText(self):
        self.textframe2.DeleteText()


class ThreeDFormat:

    def __init__(self, threedformat=None):
        self.threedformat = threedformat

    @property
    def Application(self):
        return self.threedformat.Application

    @property
    def BevelBottomDepth(self):
        return ThreeDFormat(self.threedformat.BevelBottomDepth)

    @BevelBottomDepth.setter
    def BevelBottomDepth(self, value):
        self.threedformat.BevelBottomDepth = value

    @property
    def BevelBottomInset(self):
        return ThreeDFormat(self.threedformat.BevelBottomInset)

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
        return ThreeDFormat(self.threedformat.BevelTopDepth)

    @BevelTopDepth.setter
    def BevelTopDepth(self, value):
        self.threedformat.BevelTopDepth = value

    @property
    def BevelTopInset(self):
        return ThreeDFormat(self.threedformat.BevelTopInset)

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

    @property
    def ContourWidth(self):
        return ThreeDFormat(self.threedformat.ContourWidth)

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
        return ThreeDFormat(self.threedformat.FieldOfView)

    @FieldOfView.setter
    def FieldOfView(self, value):
        self.threedformat.FieldOfView = value

    @property
    def LightAngle(self):
        return ThreeDFormat(self.threedformat.LightAngle)

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

    @PresetCamera.setter
    def PresetCamera(self, value):
        self.threedformat.PresetCamera = value

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

    @Visible.setter
    def Visible(self, value):
        self.threedformat.Visible = value

    @property
    def Z(self):
        return ThreeDFormat(self.threedformat.Z)

    @Z.setter
    def Z(self, value):
        self.threedformat.Z = value

    def IncrementRotationHorizontal(self, Increment=None):
        params = [
            Increment if Increment is not None else pythoncom.Missing,
        ]
        self.threedformat.IncrementRotationHorizontal(*params)

    def IncrementRotationVertical(self, Increment=None):
        params = [
            Increment if Increment is not None else pythoncom.Missing,
        ]
        self.threedformat.IncrementRotationVertical(*params)

    def IncrementRotationX(self, Increment=None):
        params = [
            Increment if Increment is not None else pythoncom.Missing,
        ]
        self.threedformat.IncrementRotationX(*params)

    def IncrementRotationY(self, Increment=None):
        params = [
            Increment if Increment is not None else pythoncom.Missing,
        ]
        self.threedformat.IncrementRotationY(*params)

    def IncrementRotationZ(self, Increment=None):
        params = [
            Increment if Increment is not None else pythoncom.Missing,
        ]
        self.threedformat.IncrementRotationZ(*params)

    def ResetRotation(self):
        self.threedformat.ResetRotation()

    def SetExtrusionDirection(self, PresetExtrusionDirection=None):
        params = [
            PresetExtrusionDirection if PresetExtrusionDirection is not None else pythoncom.Missing,
        ]
        self.threedformat.SetExtrusionDirection(*params)

    def SetPresetCamera(self, PresetCamera=None):
        params = [
            PresetCamera if PresetCamera is not None else pythoncom.Missing,
        ]
        self.threedformat.SetPresetCamera(*params)

    def SetThreeDFormat(self, PresetThreeDFormat=None):
        params = [
            PresetThreeDFormat if PresetThreeDFormat is not None else pythoncom.Missing,
        ]
        self.threedformat.SetThreeDFormat(*params)


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
        return Font(self.ticklabels.Font)

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
        return self.ticklabels.ReadingOrder

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.ticklabels.ReadingOrder = value

    def Delete(self):
        return self.ticklabels.Delete()

    def Select(self):
        return self.ticklabels.Select()


class Top10:

    def __init__(self, top10=None):
        self.top10 = top10

    @property
    def Application(self):
        return self.top10.Application

    @property
    def AppliesTo(self):
        return Range(self.top10.AppliesTo)

    @property
    def Borders(self):
        return Borders(self.top10.Borders)

    @property
    def CalcFor(self):
        return XlCalcFor(self.top10.CalcFor)

    @CalcFor.setter
    def CalcFor(self, value):
        self.top10.CalcFor = value

    @property
    def Creator(self):
        return self.top10.Creator

    @property
    def Font(self):
        return Font(self.top10.Font)

    @property
    def Interior(self):
        return Interior(self.top10.Interior)

    @property
    def NumberFormat(self):
        return self.top10.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value):
        self.top10.NumberFormat = value

    @property
    def Parent(self):
        return self.top10.Parent

    @property
    def Percent(self):
        return self.top10.Percent

    @Percent.setter
    def Percent(self, value):
        self.top10.Percent = value

    @property
    def Priority(self):
        return self.top10.Priority

    @Priority.setter
    def Priority(self, value):
        self.top10.Priority = value

    @property
    def PTCondition(self):
        return self.top10.PTCondition

    @property
    def Rank(self):
        return self.top10.Rank

    @Rank.setter
    def Rank(self, value):
        self.top10.Rank = value

    @property
    def ScopeType(self):
        return XlPivotConditionScope(self.top10.ScopeType)

    @ScopeType.setter
    def ScopeType(self, value):
        self.top10.ScopeType = value

    @property
    def StopIfTrue(self):
        return self.top10.StopIfTrue

    @StopIfTrue.setter
    def StopIfTrue(self, value):
        self.top10.StopIfTrue = value

    @property
    def TopBottom(self):
        return XlTopBottom(self.top10.TopBottom)

    @TopBottom.setter
    def TopBottom(self, value):
        self.top10.TopBottom = value

    @property
    def Type(self):
        return XlFormatConditionType(self.top10.Type)

    def Delete(self):
        self.top10.Delete()

    def ModifyAppliesToRange(self, Range=None):
        params = [
            Range if Range is not None else pythoncom.Missing,
        ]
        self.top10.ModifyAppliesToRange(*params)

    def SetFirstPriority(self):
        self.top10.SetFirstPriority()

    def SetLastPriority(self):
        self.top10.SetLastPriority()


class TreeviewControl:

    def __init__(self, treeviewcontrol=None):
        self.treeviewcontrol = treeviewcontrol

    @property
    def Application(self):
        return self.treeviewcontrol.Application

    @property
    def Creator(self):
        return self.treeviewcontrol.Creator

    @property
    def Drilled(self):
        return self.treeviewcontrol.Drilled

    @Drilled.setter
    def Drilled(self, value):
        self.treeviewcontrol.Drilled = value

    @property
    def Hidden(self):
        return self.treeviewcontrol.Hidden

    @Hidden.setter
    def Hidden(self, value):
        self.treeviewcontrol.Hidden = value

    @property
    def Parent(self):
        return self.treeviewcontrol.Parent


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
        return Border(self.trendline.Border)

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
        return XlTrendlineType(self.trendline.Order)

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
        return XlTrendlineType(self.trendline.Type)

    @Type.setter
    def Type(self, value):
        self.trendline.Type = value

    def ClearFormats(self):
        return self.trendline.ClearFormats()

    def Delete(self):
        return self.trendline.Delete()

    def Select(self):
        return self.trendline.Select()


class Trendlines:

    def __init__(self, trendlines=None):
        self.trendlines = trendlines

    def __call__(self, item):
        return Trendline(self.trendlines(item))

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
        params = [
            Type if Type is not None else pythoncom.Missing,
            Order if Order is not None else pythoncom.Missing,
            Period if Period is not None else pythoncom.Missing,
            Forward if Forward is not None else pythoncom.Missing,
            Backward if Backward is not None else pythoncom.Missing,
            Intercept if Intercept is not None else pythoncom.Missing,
            DisplayEquation if DisplayEquation is not None else pythoncom.Missing,
            DisplayRSquared if DisplayRSquared is not None else pythoncom.Missing,
            Name if Name is not None else pythoncom.Missing,
        ]
        return Trendline(self.trendlines.Add(*params))

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return Trendline(self.trendlines.Item(*params))


class UniqueValues:

    def __init__(self, uniquevalues=None):
        self.uniquevalues = uniquevalues

    @property
    def Application(self):
        return self.uniquevalues.Application

    @property
    def AppliesTo(self):
        return Range(self.uniquevalues.AppliesTo)

    @property
    def Borders(self):
        return Borders(self.uniquevalues.Borders)

    @property
    def Creator(self):
        return self.uniquevalues.Creator

    @property
    def DupeUnique(self):
        return XlDupeUnique(self.uniquevalues.DupeUnique)

    @DupeUnique.setter
    def DupeUnique(self, value):
        self.uniquevalues.DupeUnique = value

    @property
    def Font(self):
        return Font(self.uniquevalues.Font)

    @property
    def Interior(self):
        return Interior(self.uniquevalues.Interior)

    @property
    def NumberFormat(self):
        return self.uniquevalues.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value):
        self.uniquevalues.NumberFormat = value

    @property
    def Parent(self):
        return self.uniquevalues.Parent

    @property
    def Priority(self):
        return self.uniquevalues.Priority

    @Priority.setter
    def Priority(self, value):
        self.uniquevalues.Priority = value

    @property
    def PTCondition(self):
        return self.uniquevalues.PTCondition

    @property
    def ScopeType(self):
        return XlPivotConditionScope(self.uniquevalues.ScopeType)

    @ScopeType.setter
    def ScopeType(self, value):
        self.uniquevalues.ScopeType = value

    @property
    def StopIfTrue(self):
        return self.uniquevalues.StopIfTrue

    @StopIfTrue.setter
    def StopIfTrue(self, value):
        self.uniquevalues.StopIfTrue = value

    @property
    def Type(self):
        return XlFormatConditionType(self.uniquevalues.Type)

    def Delete(self):
        self.uniquevalues.Delete()

    def ModifyAppliesToRange(self, Range=None):
        params = [
            Range if Range is not None else pythoncom.Missing,
        ]
        self.uniquevalues.ModifyAppliesToRange(*params)

    def SetFirstPriority(self):
        self.uniquevalues.SetFirstPriority()

    def SetLastPriority(self):
        self.uniquevalues.SetLastPriority()


class UpBars:

    def __init__(self, upbars=None):
        self.upbars = upbars

    @property
    def Application(self):
        return self.upbars.Application

    @property
    def Creator(self):
        return self.upbars.Creator

    @property
    def Format(self):
        return ChartFormat(self.upbars.Format)

    @property
    def Name(self):
        return self.upbars.Name

    @property
    def Parent(self):
        return self.upbars.Parent

    def Delete(self):
        return self.upbars.Delete()

    def Select(self):
        return self.upbars.Select()


class UsedObjects:

    def __init__(self, usedobjects=None):
        self.usedobjects = usedobjects

    @property
    def Application(self):
        return self.usedobjects.Application

    @property
    def Count(self):
        return self.usedobjects.Count

    @property
    def Creator(self):
        return self.usedobjects.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.usedobjects.Item):
            return self.usedobjects.Item(*params)
        else:
            return self.usedobjects.GetItem(*params)

    @property
    def Parent(self):
        return self.usedobjects.Parent


class UserAccess:

    def __init__(self, useraccess=None):
        self.useraccess = useraccess

    @property
    def AllowEdit(self):
        return self.useraccess.AllowEdit

    @AllowEdit.setter
    def AllowEdit(self, value):
        self.useraccess.AllowEdit = value

    @property
    def Name(self):
        return self.useraccess.Name

    @Name.setter
    def Name(self, value):
        self.useraccess.Name = value

    def Delete(self):
        self.useraccess.Delete()


class UserAccessList:

    def __init__(self, useraccesslist=None):
        self.useraccesslist = useraccesslist

    def __call__(self, item):
        return UserAccessLis(self.useraccesslist(item))

    @property
    def Count(self):
        return self.useraccesslist.Count

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.useraccesslist.Item):
            return self.useraccesslist.Item(*params)
        else:
            return self.useraccesslist.GetItem(*params)

    def Add(self, Name=None, AllowEdit=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
            AllowEdit if AllowEdit is not None else pythoncom.Missing,
        ]
        return UserAccess(self.useraccesslist.Add(*params))

    def DeleteAll(self):
        self.useraccesslist.DeleteAll()


class Validation:

    def __init__(self, validation=None):
        self.validation = validation

    @property
    def AlertStyle(self):
        return XlDVAlertStyle(self.validation.AlertStyle)

    @property
    def Application(self):
        return self.validation.Application

    @property
    def Creator(self):
        return self.validation.Creator

    @property
    def ErrorMessage(self):
        return self.validation.ErrorMessage

    @ErrorMessage.setter
    def ErrorMessage(self, value):
        self.validation.ErrorMessage = value

    @property
    def ErrorTitle(self):
        return self.validation.ErrorTitle

    @ErrorTitle.setter
    def ErrorTitle(self, value):
        self.validation.ErrorTitle = value

    @property
    def Formula1(self):
        return self.validation.Formula1

    @property
    def Formula2(self):
        return self.validation.Formula2

    @property
    def IgnoreBlank(self):
        return self.validation.IgnoreBlank

    @IgnoreBlank.setter
    def IgnoreBlank(self, value):
        self.validation.IgnoreBlank = value

    @property
    def IMEMode(self):
        return XlIMEMode(self.validation.IMEMode)

    @IMEMode.setter
    def IMEMode(self, value):
        self.validation.IMEMode = value

    @property
    def InCellDropdown(self):
        return self.validation.InCellDropdown

    @InCellDropdown.setter
    def InCellDropdown(self, value):
        self.validation.InCellDropdown = value

    @property
    def InputMessage(self):
        return self.validation.InputMessage

    @InputMessage.setter
    def InputMessage(self, value):
        self.validation.InputMessage = value

    @property
    def InputTitle(self):
        return self.validation.InputTitle

    @InputTitle.setter
    def InputTitle(self, value):
        self.validation.InputTitle = value

    @property
    def Operator(self):
        return XlFormatConditionOperator(self.validation.Operator)

    @property
    def Parent(self):
        return self.validation.Parent

    @property
    def ShowError(self):
        return self.validation.ShowError

    @ShowError.setter
    def ShowError(self, value):
        self.validation.ShowError = value

    @property
    def ShowInput(self):
        return self.validation.ShowInput

    @ShowInput.setter
    def ShowInput(self, value):
        self.validation.ShowInput = value

    @property
    def Type(self):
        return XlDVType(self.validation.Type)

    @property
    def Value(self):
        return self.validation.Value

    def Add(self, Type=None, AlertStyle=None, Operator=None, Formula1=None, Formula2=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            AlertStyle if AlertStyle is not None else pythoncom.Missing,
            Operator if Operator is not None else pythoncom.Missing,
            Formula1 if Formula1 is not None else pythoncom.Missing,
            Formula2 if Formula2 is not None else pythoncom.Missing,
        ]
        self.validation.Add(*params)

    def Delete(self):
        self.validation.Delete()

    def Modify(self, Type=None, AlertStyle=None, Operator=None, Formula1=None, Formula2=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            AlertStyle if AlertStyle is not None else pythoncom.Missing,
            Operator if Operator is not None else pythoncom.Missing,
            Formula1 if Formula1 is not None else pythoncom.Missing,
            Formula2 if Formula2 is not None else pythoncom.Missing,
        ]
        self.validation.Modify(*params)


class ValueChange:

    def __init__(self, valuechange=None):
        self.valuechange = valuechange

    @property
    def AllocationMethod(self):
        return self.valuechange.AllocationMethod

    @property
    def AllocationValue(self):
        return self.valuechange.AllocationValue

    @property
    def AllocationWeightExpression(self):
        return self.valuechange.AllocationWeightExpression

    @property
    def Application(self):
        return self.valuechange.Application

    @property
    def Creator(self):
        return self.valuechange.Creator

    @property
    def Order(self):
        return PivotTableChangeList(self.valuechange.Order)

    @property
    def Parent(self):
        return self.valuechange.Parent

    @property
    def PivotCell(self):
        return PivotCell(self.valuechange.PivotCell)

    @property
    def Tuple(self):
        return self.valuechange.Tuple

    @property
    def Value(self):
        return self.valuechange.Value

    @property
    def VisibleInPivotTable(self):
        return self.valuechange.VisibleInPivotTable

    def Delete(self):
        self.valuechange.Delete()


class VPageBreak:

    def __init__(self, vpagebreak=None):
        self.vpagebreak = vpagebreak

    @property
    def Application(self):
        return self.vpagebreak.Application

    @property
    def Creator(self):
        return self.vpagebreak.Creator

    @property
    def Extent(self):
        return XlPageBreakExtent(self.vpagebreak.Extent)

    @property
    def Location(self):
        return Range(self.vpagebreak.Location)

    @property
    def Parent(self):
        return self.vpagebreak.Parent

    @property
    def Type(self):
        return XlPageBreak(self.vpagebreak.Type)

    @Type.setter
    def Type(self, value):
        self.vpagebreak.Type = value

    def Delete(self):
        self.vpagebreak.Delete()

    def DragOff(self, Direction=None, RegionIndex=None):
        params = [
            Direction if Direction is not None else pythoncom.Missing,
            RegionIndex if RegionIndex is not None else pythoncom.Missing,
        ]
        self.vpagebreak.DragOff(*params)


class VPageBreaks:

    def __init__(self, vpagebreaks=None):
        self.vpagebreaks = vpagebreaks

    def __call__(self, item):
        return VPageBreak(self.vpagebreaks(item))

    @property
    def Application(self):
        return self.vpagebreaks.Application

    @property
    def Count(self):
        return self.vpagebreaks.Count

    @property
    def Creator(self):
        return self.vpagebreaks.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.vpagebreaks.Item):
            return self.vpagebreaks.Item(*params)
        else:
            return self.vpagebreaks.GetItem(*params)

    @property
    def Parent(self):
        return self.vpagebreaks.Parent

    def Add(self, Before=None):
        params = [
            Before if Before is not None else pythoncom.Missing,
        ]
        return VPageBreak(self.vpagebreaks.Add(*params))


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
        return PictureType(self.walls.PictureUnit)

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
        return self.walls.ClearFormats()

    def Paste(self):
        self.walls.Paste()

    def Select(self):
        return self.walls.Select()


class Watch:

    def __init__(self, watch=None):
        self.watch = watch

    @property
    def Application(self):
        return self.watch.Application

    @property
    def Creator(self):
        return self.watch.Creator

    @property
    def Parent(self):
        return self.watch.Parent

    @property
    def Source(self):
        return self.watch.Source

    def Delete(self):
        self.watch.Delete()


class Watches:

    def __init__(self, watches=None):
        self.watches = watches

    def __call__(self, item):
        return Watche(self.watches(item))

    @property
    def Application(self):
        return self.watches.Application

    @property
    def Count(self):
        return self.watches.Count

    @property
    def Creator(self):
        return self.watches.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.watches.Item):
            return self.watches.Item(*params)
        else:
            return self.watches.GetItem(*params)

    @property
    def Parent(self):
        return self.watches.Parent

    def Add(self, Source=None):
        params = [
            Source if Source is not None else pythoncom.Missing,
        ]
        return Watch(self.watches.Add(*params))

    def Delete(self):
        self.watches.Delete()


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
        return self.weboptions.Application

    @property
    def Creator(self):
        return self.weboptions.Creator

    @property
    def DownloadComponents(self):
        return self.weboptions.DownloadComponents

    @DownloadComponents.setter
    def DownloadComponents(self, value):
        self.weboptions.DownloadComponents = value

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
    def LocationOfComponents(self):
        return self.weboptions.LocationOfComponents

    @LocationOfComponents.setter
    def LocationOfComponents(self, value):
        self.weboptions.LocationOfComponents = value

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
    def ActiveCell(self):
        return Range(self.window.ActiveCell)

    @property
    def ActiveChart(self):
        return Chart(self.window.ActiveChart)

    @property
    def ActivePane(self):
        return Pane(self.window.ActivePane)

    @property
    def ActiveSheet(self):
        return self.window.ActiveSheet

    @property
    def ActiveSheetView(self):
        return self.window.ActiveSheetView

    @property
    def Application(self):
        return self.window.Application

    @property
    def AutoFilterDateGrouping(self):
        return self.window.AutoFilterDateGrouping

    @AutoFilterDateGrouping.setter
    def AutoFilterDateGrouping(self, value):
        self.window.AutoFilterDateGrouping = value

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
    def DisplayFormulas(self):
        return self.window.DisplayFormulas

    @DisplayFormulas.setter
    def DisplayFormulas(self, value):
        self.window.DisplayFormulas = value

    @property
    def DisplayGridlines(self):
        return self.window.DisplayGridlines

    @DisplayGridlines.setter
    def DisplayGridlines(self, value):
        self.window.DisplayGridlines = value

    @property
    def DisplayHeadings(self):
        return self.window.DisplayHeadings

    @DisplayHeadings.setter
    def DisplayHeadings(self, value):
        self.window.DisplayHeadings = value

    @property
    def DisplayHorizontalScrollBar(self):
        return self.window.DisplayHorizontalScrollBar

    @DisplayHorizontalScrollBar.setter
    def DisplayHorizontalScrollBar(self, value):
        self.window.DisplayHorizontalScrollBar = value

    @property
    def DisplayOutline(self):
        return self.window.DisplayOutline

    @DisplayOutline.setter
    def DisplayOutline(self, value):
        self.window.DisplayOutline = value

    @property
    def DisplayRightToLeft(self):
        return self.window.DisplayRightToLeft

    @property
    def DisplayRuler(self):
        return self.window.DisplayRuler

    @DisplayRuler.setter
    def DisplayRuler(self, value):
        self.window.DisplayRuler = value

    @property
    def DisplayVerticalScrollBar(self):
        return self.window.DisplayVerticalScrollBar

    @DisplayVerticalScrollBar.setter
    def DisplayVerticalScrollBar(self, value):
        self.window.DisplayVerticalScrollBar = value

    @property
    def DisplayWhitespace(self):
        return self.window.DisplayWhitespace

    @DisplayWhitespace.setter
    def DisplayWhitespace(self, value):
        self.window.DisplayWhitespace = value

    @property
    def DisplayWorkbookTabs(self):
        return self.window.DisplayWorkbookTabs

    @DisplayWorkbookTabs.setter
    def DisplayWorkbookTabs(self, value):
        self.window.DisplayWorkbookTabs = value

    @property
    def DisplayZeros(self):
        return self.window.DisplayZeros

    @DisplayZeros.setter
    def DisplayZeros(self, value):
        self.window.DisplayZeros = value

    @property
    def EnableResize(self):
        return self.window.EnableResize

    @EnableResize.setter
    def EnableResize(self, value):
        self.window.EnableResize = value

    @property
    def FreezePanes(self):
        return self.window.FreezePanes

    @FreezePanes.setter
    def FreezePanes(self, value):
        self.window.FreezePanes = value

    @property
    def GridlineColor(self):
        return self.window.GridlineColor

    @GridlineColor.setter
    def GridlineColor(self, value):
        self.window.GridlineColor = value

    @property
    def GridlineColorIndex(self):
        return XlColorIndex(self.window.GridlineColorIndex)

    @GridlineColorIndex.setter
    def GridlineColorIndex(self, value):
        self.window.GridlineColorIndex = value

    @property
    def Height(self):
        return self.window.Height

    @Height.setter
    def Height(self, value):
        self.window.Height = value

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
    def OnWindow(self):
        return self.window.OnWindow

    @OnWindow.setter
    def OnWindow(self, value):
        self.window.OnWindow = value

    @property
    def Panes(self):
        return Panes(self.window.Panes)

    @property
    def Parent(self):
        return self.window.Parent

    @property
    def RangeSelection(self):
        return Range(self.window.RangeSelection)

    @property
    def ScrollColumn(self):
        return self.window.ScrollColumn

    @ScrollColumn.setter
    def ScrollColumn(self, value):
        self.window.ScrollColumn = value

    @property
    def ScrollRow(self):
        return self.window.ScrollRow

    @ScrollRow.setter
    def ScrollRow(self, value):
        self.window.ScrollRow = value

    @property
    def SelectedSheets(self):
        return Sheets(self.window.SelectedSheets)

    @property
    def Selection(self):
        return Windows(self.window.Selection)

    @property
    def SheetViews(self):
        return SheetViews(self.window.SheetViews)

    @property
    def Split(self):
        return self.window.Split

    @Split.setter
    def Split(self, value):
        self.window.Split = value

    @property
    def SplitColumn(self):
        return self.window.SplitColumn

    @SplitColumn.setter
    def SplitColumn(self, value):
        self.window.SplitColumn = value

    @property
    def SplitHorizontal(self):
        return self.window.SplitHorizontal

    @SplitHorizontal.setter
    def SplitHorizontal(self, value):
        self.window.SplitHorizontal = value

    @property
    def SplitRow(self):
        return self.window.SplitRow

    @SplitRow.setter
    def SplitRow(self, value):
        self.window.SplitRow = value

    @property
    def SplitVertical(self):
        return self.window.SplitVertical

    @SplitVertical.setter
    def SplitVertical(self, value):
        self.window.SplitVertical = value

    @property
    def TabRatio(self):
        return self.window.TabRatio

    @TabRatio.setter
    def TabRatio(self, value):
        self.window.TabRatio = value

    @property
    def Top(self):
        return self.window.Top

    @Top.setter
    def Top(self, value):
        self.window.Top = value

    @property
    def Type(self):
        return XlWindowType(self.window.Type)

    @Type.setter
    def Type(self, value):
        self.window.Type = value

    @property
    def UsableHeight(self):
        return self.window.UsableHeight

    @property
    def UsableWidth(self):
        return self.window.UsableWidth

    @property
    def View(self):
        return XlWindowView(self.window.View)

    @View.setter
    def View(self, value):
        self.window.View = value

    @property
    def Visible(self):
        return self.window.Visible

    @Visible.setter
    def Visible(self, value):
        self.window.Visible = value

    @property
    def VisibleRange(self):
        return Range(self.window.VisibleRange)

    @property
    def Width(self):
        return self.window.Width

    @Width.setter
    def Width(self, value):
        self.window.Width = value

    @property
    def WindowNumber(self):
        return self.window.WindowNumber

    @property
    def WindowState(self):
        return XlWindowState(self.window.WindowState)

    @WindowState.setter
    def WindowState(self, value):
        self.window.WindowState = value

    @property
    def Zoom(self):
        return self.window.Zoom

    @Zoom.setter
    def Zoom(self, value):
        self.window.Zoom = value

    def Activate(self):
        return self.window.Activate()

    def ActivateNext(self):
        return self.window.ActivateNext()

    def ActivatePrevious(self):
        return self.window.ActivatePrevious()

    def Close(self, SaveChanges=None, FileName=None, RouteWorkbook=None):
        params = [
            SaveChanges if SaveChanges is not None else pythoncom.Missing,
            FileName if FileName is not None else pythoncom.Missing,
            RouteWorkbook if RouteWorkbook is not None else pythoncom.Missing,
        ]
        return self.window.Close(*params)

    def LargeScroll(self, Down=None, Up=None, ToRight=None, ToLeft=None):
        params = [
            Down if Down is not None else pythoncom.Missing,
            Up if Up is not None else pythoncom.Missing,
            ToRight if ToRight is not None else pythoncom.Missing,
            ToLeft if ToLeft is not None else pythoncom.Missing,
        ]
        return self.window.LargeScroll(*params)

    def NewWindow(self):
        return self.window.NewWindow()

    def PointsToScreenPixelsX(self, Points=None):
        params = [
            Points if Points is not None else pythoncom.Missing,
        ]
        return self.window.PointsToScreenPixelsX(*params)

    def PointsToScreenPixelsY(self, Points=None):
        params = [
            Points if Points is not None else pythoncom.Missing,
        ]
        return self.window.PointsToScreenPixelsY(*params)

    def PrintOut(self, From=None, To=None, Copies=None, Preview=None, ActivePrinter=None, PrintToFile=None, Collate=None, PrToFileName=None):
        params = [
            From if From is not None else pythoncom.Missing,
            To if To is not None else pythoncom.Missing,
            Copies if Copies is not None else pythoncom.Missing,
            Preview if Preview is not None else pythoncom.Missing,
            ActivePrinter if ActivePrinter is not None else pythoncom.Missing,
            PrintToFile if PrintToFile is not None else pythoncom.Missing,
            Collate if Collate is not None else pythoncom.Missing,
            PrToFileName if PrToFileName is not None else pythoncom.Missing,
        ]
        return self.window.PrintOut(*params)

    def PrintPreview(self, EnableChanges=None):
        params = [
            EnableChanges if EnableChanges is not None else pythoncom.Missing,
        ]
        return self.window.PrintPreview(*params)

    def RangeFromPoint(self, x=None, y=None):
        params = [
            x if x is not None else pythoncom.Missing,
            y if y is not None else pythoncom.Missing,
        ]
        return self.window.RangeFromPoint(*params)

    def ScrollIntoView(self, Left=None, Top=None, Width=None, Height=None, Start=None):
        params = [
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            Width if Width is not None else pythoncom.Missing,
            Height if Height is not None else pythoncom.Missing,
            Start if Start is not None else pythoncom.Missing,
        ]
        self.window.ScrollIntoView(*params)

    def ScrollWorkbookTabs(self, Sheets=None, Position=None):
        params = [
            Sheets if Sheets is not None else pythoncom.Missing,
            Position if Position is not None else pythoncom.Missing,
        ]
        return self.window.ScrollWorkbookTabs(*params)

    def SmallScroll(self, Down=None, Up=None, ToRight=None, ToLeft=None):
        params = [
            Down if Down is not None else pythoncom.Missing,
            Up if Up is not None else pythoncom.Missing,
            ToRight if ToRight is not None else pythoncom.Missing,
            ToLeft if ToLeft is not None else pythoncom.Missing,
        ]
        return self.window.SmallScroll(*params)


class Windows:

    def __init__(self, windows=None):
        self.windows = windows

    def __call__(self, item):
        return Window(self.windows(item))

    @property
    def Application(self):
        return self.windows.Application

    @property
    def Count(self):
        return self.windows.Count

    @property
    def Creator(self):
        return self.windows.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.windows.Item):
            return self.windows.Item(*params)
        else:
            return self.windows.GetItem(*params)

    @property
    def Parent(self):
        return self.windows.Parent

    @property
    def SyncScrollingSideBySide(self):
        return self.windows.SyncScrollingSideBySide

    def Arrange(self, ArrangeStyle=None, ActiveWorkbook=None, SyncHorizontal=None, SyncVertical=None):
        params = [
            ArrangeStyle if ArrangeStyle is not None else pythoncom.Missing,
            ActiveWorkbook if ActiveWorkbook is not None else pythoncom.Missing,
            SyncHorizontal if SyncHorizontal is not None else pythoncom.Missing,
            SyncVertical if SyncVertical is not None else pythoncom.Missing,
        ]
        return self.windows.Arrange(*params)

    def BreakSideBySide(self):
        return self.windows.BreakSideBySide()

    def CompareSideBySideWith(self, WindowName=None):
        params = [
            WindowName if WindowName is not None else pythoncom.Missing,
        ]
        return self.windows.CompareSideBySideWith(*params)

    def ResetPositionsSideBySide(self):
        self.windows.ResetPositionsSideBySide()


class Workbook:

    def __init__(self, workbook=None):
        self.workbook = workbook

    @property
    def AccuracyVersion(self):
        return self.workbook.AccuracyVersion

    @AccuracyVersion.setter
    def AccuracyVersion(self, value):
        self.workbook.AccuracyVersion = value

    @property
    def ActiveChart(self):
        return Chart(self.workbook.ActiveChart)

    @property
    def ActiveSheet(self):
        return Worksheet(self.workbook.ActiveSheet)

    @property
    def ActiveSlicer(self):
        return self.workbook.ActiveSlicer

    @property
    def Application(self):
        return self.workbook.Application

    @property
    def AutoSaveOn(self):
        return self.workbook.AutoSaveOn

    @AutoSaveOn.setter
    def AutoSaveOn(self, value):
        self.workbook.AutoSaveOn = value

    @property
    def AutoUpdateFrequency(self):
        return self.workbook.AutoUpdateFrequency

    @AutoUpdateFrequency.setter
    def AutoUpdateFrequency(self, value):
        self.workbook.AutoUpdateFrequency = value

    @property
    def AutoUpdateSaveChanges(self):
        return self.workbook.AutoUpdateSaveChanges

    @AutoUpdateSaveChanges.setter
    def AutoUpdateSaveChanges(self, value):
        self.workbook.AutoUpdateSaveChanges = value

    @property
    def BuiltinDocumentProperties(self):
        return self.workbook.BuiltinDocumentProperties

    @property
    def CalculationVersion(self):
        return self.workbook.CalculationVersion

    @property
    def ChangeHistoryDuration(self):
        return self.workbook.ChangeHistoryDuration

    @ChangeHistoryDuration.setter
    def ChangeHistoryDuration(self, value):
        self.workbook.ChangeHistoryDuration = value

    @property
    def Charts(self):
        return Sheets(self.workbook.Charts)

    @property
    def CheckCompatibility(self):
        return self.workbook.CheckCompatibility

    @CheckCompatibility.setter
    def CheckCompatibility(self, value):
        self.workbook.CheckCompatibility = value

    @property
    def CodeName(self):
        return self.workbook.CodeName

    @property
    def Colors(self):
        return self.workbook.Colors

    @Colors.setter
    def Colors(self, value):
        self.workbook.Colors = value

    @property
    def CommandBars(self):
        return CommandBars(self.workbook.CommandBars)

    @property
    def ConflictResolution(self):
        return XlSaveConflictResolution(self.workbook.ConflictResolution)

    @ConflictResolution.setter
    def ConflictResolution(self, value):
        self.workbook.ConflictResolution = value

    @property
    def Connections(self):
        return self.workbook.Connections

    @property
    def ConnectionsDisabled(self):
        return self.workbook.ConnectionsDisabled

    @property
    def Container(self):
        return self.workbook.Container

    @property
    def ContentTypeProperties(self):
        return self.workbook.ContentTypeProperties

    @property
    def CreateBackup(self):
        return self.workbook.CreateBackup

    @property
    def Creator(self):
        return self.workbook.Creator

    @property
    def CustomDocumentProperties(self):
        return self.workbook.CustomDocumentProperties

    @CustomDocumentProperties.setter
    def CustomDocumentProperties(self, value):
        self.workbook.CustomDocumentProperties = value

    @property
    def CustomViews(self):
        return CustomViews(self.workbook.CustomViews)

    @property
    def CustomXMLParts(self):
        return self.workbook.CustomXMLParts

    @property
    def Date1904(self):
        return self.workbook.Date1904

    @Date1904.setter
    def Date1904(self, value):
        self.workbook.Date1904 = value

    @property
    def DefaultPivotTableStyle(self):
        return self.workbook.DefaultPivotTableStyle

    @DefaultPivotTableStyle.setter
    def DefaultPivotTableStyle(self, value):
        self.workbook.DefaultPivotTableStyle = value

    @property
    def DefaultSlicerStyle(self):
        return self.workbook.DefaultSlicerStyle

    @DefaultSlicerStyle.setter
    def DefaultSlicerStyle(self, value):
        self.workbook.DefaultSlicerStyle = value

    @property
    def DefaultTableStyle(self):
        return self.workbook.DefaultTableStyle

    @DefaultTableStyle.setter
    def DefaultTableStyle(self, value):
        self.workbook.DefaultTableStyle = value

    @property
    def DisplayDrawingObjects(self):
        return self.workbook.DisplayDrawingObjects

    @DisplayDrawingObjects.setter
    def DisplayDrawingObjects(self, value):
        self.workbook.DisplayDrawingObjects = value

    @property
    def DisplayInkComments(self):
        return self.workbook.DisplayInkComments

    @DisplayInkComments.setter
    def DisplayInkComments(self, value):
        self.workbook.DisplayInkComments = value

    @property
    def DocumentInspectors(self):
        return self.workbook.DocumentInspectors

    @property
    def DocumentLibraryVersions(self):
        return self.workbook.DocumentLibraryVersions

    @property
    def DoNotPromptForConvert(self):
        return self.workbook.DoNotPromptForConvert

    @DoNotPromptForConvert.setter
    def DoNotPromptForConvert(self, value):
        self.workbook.DoNotPromptForConvert = value

    @property
    def EnableAutoRecover(self):
        return self.workbook.EnableAutoRecover

    @EnableAutoRecover.setter
    def EnableAutoRecover(self, value):
        self.workbook.EnableAutoRecover = value

    @property
    def EncryptionProvider(self):
        return self.workbook.EncryptionProvider

    @EncryptionProvider.setter
    def EncryptionProvider(self, value):
        self.workbook.EncryptionProvider = value

    @property
    def EnvelopeVisible(self):
        return self.workbook.EnvelopeVisible

    @EnvelopeVisible.setter
    def EnvelopeVisible(self, value):
        self.workbook.EnvelopeVisible = value

    @property
    def Excel4IntlMacroSheets(self):
        return Sheets(self.workbook.Excel4IntlMacroSheets)

    @property
    def Excel4MacroSheets(self):
        return Sheets(self.workbook.Excel4MacroSheets)

    @property
    def Excel8CompatibilityMode(self):
        return self.workbook.Excel8CompatibilityMode

    @property
    def FileFormat(self):
        return XlFileFormat(self.workbook.FileFormat)

    @property
    def Final(self):
        return self.workbook.Final

    @Final.setter
    def Final(self, value):
        self.workbook.Final = value

    @property
    def ForceFullCalculation(self):
        return self.workbook.ForceFullCalculation

    @ForceFullCalculation.setter
    def ForceFullCalculation(self, value):
        self.workbook.ForceFullCalculation = value

    @property
    def FullName(self):
        return self.workbook.FullName

    @property
    def FullNameURLEncoded(self):
        return self.workbook.FullNameURLEncoded

    @property
    def HasPassword(self):
        return self.workbook.HasPassword

    @property
    def HasVBProject(self):
        return self.workbook.HasVBProject

    @property
    def HighlightChangesOnScreen(self):
        return self.workbook.HighlightChangesOnScreen

    @HighlightChangesOnScreen.setter
    def HighlightChangesOnScreen(self, value):
        self.workbook.HighlightChangesOnScreen = value

    @property
    def IconSets(self):
        return self.workbook.IconSets

    @property
    def InactiveListBorderVisible(self):
        return self.workbook.InactiveListBorderVisible

    @InactiveListBorderVisible.setter
    def InactiveListBorderVisible(self, value):
        self.workbook.InactiveListBorderVisible = value

    @property
    def IsAddin(self):
        return self.workbook.IsAddin

    @IsAddin.setter
    def IsAddin(self, value):
        self.workbook.IsAddin = value

    @property
    def IsInplace(self):
        return self.workbook.IsInplace

    @property
    def KeepChangeHistory(self):
        return self.workbook.KeepChangeHistory

    @KeepChangeHistory.setter
    def KeepChangeHistory(self, value):
        self.workbook.KeepChangeHistory = value

    @property
    def ListChangesOnNewSheet(self):
        return self.workbook.ListChangesOnNewSheet

    @ListChangesOnNewSheet.setter
    def ListChangesOnNewSheet(self, value):
        self.workbook.ListChangesOnNewSheet = value

    @property
    def Mailer(self):
        return self.workbook.Mailer

    @property
    def MultiUserEditing(self):
        return self.workbook.MultiUserEditing

    @property
    def Name(self):
        return self.workbook.Name

    @property
    def Names(self):
        return Names(self.workbook.Names)

    @property
    def Parent(self):
        return self.workbook.Parent

    @property
    def Password(self):
        return self.workbook.Password

    @Password.setter
    def Password(self, value):
        self.workbook.Password = value

    @property
    def PasswordEncryptionAlgorithm(self):
        return self.workbook.PasswordEncryptionAlgorithm

    @property
    def PasswordEncryptionFileProperties(self):
        return self.workbook.PasswordEncryptionFileProperties

    @property
    def PasswordEncryptionKeyLength(self):
        return self.workbook.PasswordEncryptionKeyLength

    @property
    def PasswordEncryptionProvider(self):
        return self.workbook.PasswordEncryptionProvider

    @property
    def Path(self):
        return self.workbook.Path

    @property
    def Permission(self):
        return self.workbook.Permission

    @property
    def PersonalViewListSettings(self):
        return self.workbook.PersonalViewListSettings

    @PersonalViewListSettings.setter
    def PersonalViewListSettings(self, value):
        self.workbook.PersonalViewListSettings = value

    @property
    def PersonalViewPrintSettings(self):
        return self.workbook.PersonalViewPrintSettings

    @PersonalViewPrintSettings.setter
    def PersonalViewPrintSettings(self, value):
        self.workbook.PersonalViewPrintSettings = value

    @property
    def PrecisionAsDisplayed(self):
        return self.workbook.PrecisionAsDisplayed

    @PrecisionAsDisplayed.setter
    def PrecisionAsDisplayed(self, value):
        self.workbook.PrecisionAsDisplayed = value

    @property
    def ProtectStructure(self):
        return self.workbook.ProtectStructure

    @property
    def ProtectWindows(self):
        return self.workbook.ProtectWindows

    @property
    def PublishObjects(self):
        return PublishObjects(self.workbook.PublishObjects)

    @property
    def ReadOnly(self):
        return self.workbook.ReadOnly

    @property
    def ReadOnlyRecommended(self):
        return self.workbook.ReadOnlyRecommended

    @property
    def RemovePersonalInformation(self):
        return self.workbook.RemovePersonalInformation

    @RemovePersonalInformation.setter
    def RemovePersonalInformation(self, value):
        self.workbook.RemovePersonalInformation = value

    @property
    def Research(self):
        return Research(self.workbook.Research)

    @property
    def RevisionNumber(self):
        return self.workbook.RevisionNumber

    @property
    def Saved(self):
        return self.workbook.Saved

    @Saved.setter
    def Saved(self, value):
        self.workbook.Saved = value

    @property
    def SaveLinkValues(self):
        return self.workbook.SaveLinkValues

    @SaveLinkValues.setter
    def SaveLinkValues(self, value):
        self.workbook.SaveLinkValues = value

    @property
    def Reply(self):
        return self.workbook.Reply

    @property
    def ServerPolicy(self):
        return self.workbook.ServerPolicy

    @property
    def ServerViewableItems(self):
        return self.workbook.ServerViewableItems

    @property
    def SharedWorkspace(self):
        return self.workbook.SharedWorkspace

    @property
    def Sheets(self):
        return Sheets(self.workbook.Sheets)

    @property
    def ShowConflictHistory(self):
        return self.workbook.ShowConflictHistory

    @ShowConflictHistory.setter
    def ShowConflictHistory(self, value):
        self.workbook.ShowConflictHistory = value

    @property
    def ShowPivotChartActiveFields(self):
        return self.workbook.ShowPivotChartActiveFields

    @ShowPivotChartActiveFields.setter
    def ShowPivotChartActiveFields(self, value):
        self.workbook.ShowPivotChartActiveFields = value

    @property
    def ShowPivotTableFieldList(self):
        return self.workbook.ShowPivotTableFieldList

    @ShowPivotTableFieldList.setter
    def ShowPivotTableFieldList(self, value):
        self.workbook.ShowPivotTableFieldList = value

    @property
    def Signatures(self):
        return self.workbook.Signatures

    @property
    def SlicerCaches(self):
        return SlicerCaches(self.workbook.SlicerCaches)

    @property
    def SmartDocument(self):
        return self.workbook.SmartDocument

    @property
    def Styles(self):
        return Styles(self.workbook.Styles)

    @property
    def Sync(self):
        return self.workbook.Sync

    @property
    def TableStyles(self):
        return TableStyles(self.workbook.TableStyles)

    @property
    def TemplateRemoveExtData(self):
        return self.workbook.TemplateRemoveExtData

    @TemplateRemoveExtData.setter
    def TemplateRemoveExtData(self, value):
        self.workbook.TemplateRemoveExtData = value

    @property
    def Theme(self):
        return self.workbook.Theme

    @property
    def UpdateLinks(self):
        return self.workbook.UpdateLinks

    @UpdateLinks.setter
    def UpdateLinks(self, value):
        self.workbook.UpdateLinks = value

    @property
    def UpdateRemoteReferences(self):
        return self.workbook.UpdateRemoteReferences

    @UpdateRemoteReferences.setter
    def UpdateRemoteReferences(self, value):
        self.workbook.UpdateRemoteReferences = value

    @property
    def UserStatus(self):
        return self.workbook.UserStatus

    @property
    def VBASigned(self):
        return self.workbook.VBASigned

    @property
    def VBProject(self):
        return self.workbook.VBProject

    @property
    def WebOptions(self):
        return WebOptions(self.workbook.WebOptions)

    @property
    def Windows(self):
        return Windows(self.workbook.Windows)

    @property
    def Worksheets(self):
        return Worksheets(self.workbook.Worksheets)

    @property
    def WritePassword(self):
        return self.workbook.WritePassword

    @WritePassword.setter
    def WritePassword(self, value):
        self.workbook.WritePassword = value

    @property
    def WriteReserved(self):
        return self.workbook.WriteReserved

    @property
    def WriteReservedBy(self):
        return self.workbook.WriteReservedBy

    @property
    def XmlMaps(self):
        return XmlMaps(self.workbook.XmlMaps)

    @property
    def XmlNamespaces(self):
        return XmlNamespaces(self.workbook.XmlNamespaces)

    def AcceptAllChanges(self, When=None, Who=None, Where=None):
        params = [
            When if When is not None else pythoncom.Missing,
            Who if Who is not None else pythoncom.Missing,
            Where if Where is not None else pythoncom.Missing,
        ]
        self.workbook.AcceptAllChanges(*params)

    def Activate(self):
        self.workbook.Activate()

    def AddToFavorites(self):
        self.workbook.AddToFavorites()

    def ApplyTheme(self, FileName=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
        ]
        self.workbook.ApplyTheme(*params)

    def BreakLink(self, Name=None, Type=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
            Type if Type is not None else pythoncom.Missing,
        ]
        self.workbook.BreakLink(*params)

    def CanCheckIn(self):
        return self.workbook.CanCheckIn()

    def ChangeFileAccess(self, Mode=None, WritePassword=None, Notify=None):
        params = [
            Mode if Mode is not None else pythoncom.Missing,
            WritePassword if WritePassword is not None else pythoncom.Missing,
            Notify if Notify is not None else pythoncom.Missing,
        ]
        self.workbook.ChangeFileAccess(*params)

    def ChangeLink(self, Name=None, NewName=None, Type=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
            NewName if NewName is not None else pythoncom.Missing,
            Type if Type is not None else pythoncom.Missing,
        ]
        self.workbook.ChangeLink(*params)

    def CheckIn(self, SaveChanges=None, Comments=None, MakePublic=None):
        params = [
            SaveChanges if SaveChanges is not None else pythoncom.Missing,
            Comments if Comments is not None else pythoncom.Missing,
            MakePublic if MakePublic is not None else pythoncom.Missing,
        ]
        self.workbook.CheckIn(*params)

    def CheckInWithVersion(self, SaveChanges=None, Comments=None, MakePublic=None, VersionType=None):
        params = [
            SaveChanges if SaveChanges is not None else pythoncom.Missing,
            Comments if Comments is not None else pythoncom.Missing,
            MakePublic if MakePublic is not None else pythoncom.Missing,
            VersionType if VersionType is not None else pythoncom.Missing,
        ]
        return self.workbook.CheckInWithVersion(*params)

    def Close(self, SaveChanges=None, FileName=None, RouteWorkbook=None):
        params = [
            SaveChanges if SaveChanges is not None else pythoncom.Missing,
            FileName if FileName is not None else pythoncom.Missing,
            RouteWorkbook if RouteWorkbook is not None else pythoncom.Missing,
        ]
        self.workbook.Close(*params)

    def ConvertComments(self):
        self.workbook.ConvertComments()

    def DeleteNumberFormat(self, NumberFormat=None):
        params = [
            NumberFormat if NumberFormat is not None else pythoncom.Missing,
        ]
        self.workbook.DeleteNumberFormat(*params)

    def EnableConnections(self):
        self.workbook.EnableConnections()

    def EndReview(self):
        self.workbook.EndReview()

    def ExclusiveAccess(self):
        return self.workbook.ExclusiveAccess()

    def ExportAsFixedFormat(self, Type=None, FileName=None, Quality=None, IncludeDocProperties=None, IgnorePrintAreas=None, From=None, To=None, OpenAfterPublish=None, FixedFormatExtClassPtr=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            FileName if FileName is not None else pythoncom.Missing,
            Quality if Quality is not None else pythoncom.Missing,
            IncludeDocProperties if IncludeDocProperties is not None else pythoncom.Missing,
            IgnorePrintAreas if IgnorePrintAreas is not None else pythoncom.Missing,
            From if From is not None else pythoncom.Missing,
            To if To is not None else pythoncom.Missing,
            OpenAfterPublish if OpenAfterPublish is not None else pythoncom.Missing,
            FixedFormatExtClassPtr if FixedFormatExtClassPtr is not None else pythoncom.Missing,
        ]
        self.workbook.ExportAsFixedFormat(*params)

    def FollowHyperlink(self, Address=None, SubAddress=None, NewWindow=None, AddHistory=None, ExtraInfo=None, Method=None, HeaderInfo=None):
        params = [
            Address if Address is not None else pythoncom.Missing,
            SubAddress if SubAddress is not None else pythoncom.Missing,
            NewWindow if NewWindow is not None else pythoncom.Missing,
            AddHistory if AddHistory is not None else pythoncom.Missing,
            ExtraInfo if ExtraInfo is not None else pythoncom.Missing,
            Method if Method is not None else pythoncom.Missing,
            HeaderInfo if HeaderInfo is not None else pythoncom.Missing,
        ]
        self.workbook.FollowHyperlink(*params)

    def ForwardMailer(self):
        self.workbook.ForwardMailer()

    def GetWorkflowTasks(self):
        return self.workbook.GetWorkflowTasks()

    def GetWorkflowTemplates(self):
        return self.workbook.GetWorkflowTemplates()

    def HighlightChangesOptions(self, When=None, Who=None, Where=None):
        params = [
            When if When is not None else pythoncom.Missing,
            Who if Who is not None else pythoncom.Missing,
            Where if Where is not None else pythoncom.Missing,
        ]
        self.workbook.HighlightChangesOptions(*params)

    def LinkInfo(self, Name=None, LinkInfo=None, Type=None, EditionRef=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
            LinkInfo if LinkInfo is not None else pythoncom.Missing,
            Type if Type is not None else pythoncom.Missing,
            EditionRef if EditionRef is not None else pythoncom.Missing,
        ]
        return self.workbook.LinkInfo(*params)

    def LinkSources(self, Type=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
        ]
        return self.workbook.LinkSources(*params)

    def LockServerFile(self):
        self.workbook.LockServerFile()

    def MergeWorkbook(self, FileName=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
        ]
        self.workbook.MergeWorkbook(*params)

    def NewWindow(self):
        return self.workbook.NewWindow()

    def OpenLinks(self, Name=None, ReadOnly=None, Type=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
            ReadOnly if ReadOnly is not None else pythoncom.Missing,
            Type if Type is not None else pythoncom.Missing,
        ]
        self.workbook.OpenLinks(*params)

    def PivotCaches(self):
        return self.workbook.PivotCaches()

    def Post(self, DestName=None):
        params = [
            DestName if DestName is not None else pythoncom.Missing,
        ]
        self.workbook.Post(*params)

    def PrintOut(self, From=None, To=None, Copies=None, Preview=None, ActivePrinter=None, PrintToFile=None, Collate=None, PrToFileName=None, IgnorePrintAreas=None):
        params = [
            From if From is not None else pythoncom.Missing,
            To if To is not None else pythoncom.Missing,
            Copies if Copies is not None else pythoncom.Missing,
            Preview if Preview is not None else pythoncom.Missing,
            ActivePrinter if ActivePrinter is not None else pythoncom.Missing,
            PrintToFile if PrintToFile is not None else pythoncom.Missing,
            Collate if Collate is not None else pythoncom.Missing,
            PrToFileName if PrToFileName is not None else pythoncom.Missing,
            IgnorePrintAreas if IgnorePrintAreas is not None else pythoncom.Missing,
        ]
        return self.workbook.PrintOut(*params)

    def PrintPreview(self, EnableChanges=None):
        params = [
            EnableChanges if EnableChanges is not None else pythoncom.Missing,
        ]
        self.workbook.PrintPreview(*params)

    def Protect(self, Password=None, Structure=None, Windows=None):
        params = [
            Password if Password is not None else pythoncom.Missing,
            Structure if Structure is not None else pythoncom.Missing,
            Windows if Windows is not None else pythoncom.Missing,
        ]
        self.workbook.Protect(*params)

    def ProtectSharing(self, FileName=None, Password=None, WriteResPassword=None, ReadOnlyRecommended=None, CreateBackup=None, SharingPassword=None, FileFormat=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            Password if Password is not None else pythoncom.Missing,
            WriteResPassword if WriteResPassword is not None else pythoncom.Missing,
            ReadOnlyRecommended if ReadOnlyRecommended is not None else pythoncom.Missing,
            CreateBackup if CreateBackup is not None else pythoncom.Missing,
            SharingPassword if SharingPassword is not None else pythoncom.Missing,
            FileFormat if FileFormat is not None else pythoncom.Missing,
        ]
        self.workbook.ProtectSharing(*params)

    def PurgeChangeHistoryNow(self, Days=None, SharingPassword=None):
        params = [
            Days if Days is not None else pythoncom.Missing,
            SharingPassword if SharingPassword is not None else pythoncom.Missing,
        ]
        self.workbook.PurgeChangeHistoryNow(*params)

    def RefreshAll(self):
        self.workbook.RefreshAll()

    def RejectAllChanges(self, When=None, Who=None, Where=None):
        params = [
            When if When is not None else pythoncom.Missing,
            Who if Who is not None else pythoncom.Missing,
            Where if Where is not None else pythoncom.Missing,
        ]
        self.workbook.RejectAllChanges(*params)

    def ReloadAs(self, Encoding=None):
        params = [
            Encoding if Encoding is not None else pythoncom.Missing,
        ]
        self.workbook.ReloadAs(*params)

    def RemoveDocumentInformation(self, RemoveDocInfoType=None):
        params = [
            RemoveDocInfoType if RemoveDocInfoType is not None else pythoncom.Missing,
        ]
        self.workbook.RemoveDocumentInformation(*params)

    def RemoveUser(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        self.workbook.RemoveUser(*params)

    def Reply(self):
        self.workbook.Reply()

    def ReplyAll(self):
        self.workbook.ReplyAll()

    def ReplyWithChanges(self, ShowMessage=None):
        params = [
            ShowMessage if ShowMessage is not None else pythoncom.Missing,
        ]
        self.workbook.ReplyWithChanges(*params)

    def ResetColors(self):
        self.workbook.ResetColors()

    def RunAutoMacros(self, Which=None):
        params = [
            Which if Which is not None else pythoncom.Missing,
        ]
        self.workbook.RunAutoMacros(*params)

    def Save(self):
        self.workbook.Save()

    def SaveAs(self, FileName=None, FileFormat=None, Password=None, WriteResPassword=None, ReadOnlyRecommended=None, CreateBackup=None, AccessMode=None, ConflictResolution=None, AddToMru=None, TextCodepage=None, TextVisualLayout=None, Local=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            FileFormat if FileFormat is not None else pythoncom.Missing,
            Password if Password is not None else pythoncom.Missing,
            WriteResPassword if WriteResPassword is not None else pythoncom.Missing,
            ReadOnlyRecommended if ReadOnlyRecommended is not None else pythoncom.Missing,
            CreateBackup if CreateBackup is not None else pythoncom.Missing,
            AccessMode if AccessMode is not None else pythoncom.Missing,
            ConflictResolution if ConflictResolution is not None else pythoncom.Missing,
            AddToMru if AddToMru is not None else pythoncom.Missing,
            TextCodepage if TextCodepage is not None else pythoncom.Missing,
            TextVisualLayout if TextVisualLayout is not None else pythoncom.Missing,
            Local if Local is not None else pythoncom.Missing,
        ]
        self.workbook.SaveAs(*params)

    def SaveAsXMLData(self, FileName=None, Map=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            Map if Map is not None else pythoncom.Missing,
        ]
        self.workbook.SaveAsXMLData(*params)

    def SaveCopyAs(self, FileName=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
        ]
        self.workbook.SaveCopyAs(*params)

    def SendFaxOverInternet(self, Recipients=None, Subject=None, ShowMessage=None):
        params = [
            Recipients if Recipients is not None else pythoncom.Missing,
            Subject if Subject is not None else pythoncom.Missing,
            ShowMessage if ShowMessage is not None else pythoncom.Missing,
        ]
        self.workbook.SendFaxOverInternet(*params)

    def SendForReview(self, Recipients=None, Subject=None, ShowMessage=None, IncludeAttachment=None):
        params = [
            Recipients if Recipients is not None else pythoncom.Missing,
            Subject if Subject is not None else pythoncom.Missing,
            ShowMessage if ShowMessage is not None else pythoncom.Missing,
            IncludeAttachment if IncludeAttachment is not None else pythoncom.Missing,
        ]
        self.workbook.SendForReview(*params)

    def SendMail(self, Recipients=None, Subject=None, ReturnReceipt=None):
        params = [
            Recipients if Recipients is not None else pythoncom.Missing,
            Subject if Subject is not None else pythoncom.Missing,
            ReturnReceipt if ReturnReceipt is not None else pythoncom.Missing,
        ]
        self.workbook.SendMail(*params)

    def SendMailer(self, FileFormat=None, Priority=None):
        params = [
            FileFormat if FileFormat is not None else pythoncom.Missing,
            Priority if Priority is not None else pythoncom.Missing,
        ]
        self.workbook.SendMailer(*params)

    def SetLinkOnData(self, Name=None, Procedure=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
            Procedure if Procedure is not None else pythoncom.Missing,
        ]
        self.workbook.SetLinkOnData(*params)

    def SetPasswordEncryptionOptions(self, PasswordEncryptionProvider=None, PasswordEncryptionAlgorithm=None, PasswordEncryptionKeyLength=None, PasswordEncryptionFileProperties=None):
        params = [
            PasswordEncryptionProvider if PasswordEncryptionProvider is not None else pythoncom.Missing,
            PasswordEncryptionAlgorithm if PasswordEncryptionAlgorithm is not None else pythoncom.Missing,
            PasswordEncryptionKeyLength if PasswordEncryptionKeyLength is not None else pythoncom.Missing,
            PasswordEncryptionFileProperties if PasswordEncryptionFileProperties is not None else pythoncom.Missing,
        ]
        self.workbook.SetPasswordEncryptionOptions(*params)

    def ToggleFormsDesign(self):
        self.workbook.ToggleFormsDesign()

    def Unprotect(self, Password=None):
        params = [
            Password if Password is not None else pythoncom.Missing,
        ]
        self.workbook.Unprotect(*params)

    def UnprotectSharing(self, SharingPassword=None):
        params = [
            SharingPassword if SharingPassword is not None else pythoncom.Missing,
        ]
        self.workbook.UnprotectSharing(*params)

    def UpdateFromFile(self):
        self.workbook.UpdateFromFile()

    def UpdateLink(self, Name=None, Type=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
            Type if Type is not None else pythoncom.Missing,
        ]
        self.workbook.UpdateLink(*params)

    def WebPagePreview(self):
        self.workbook.WebPagePreview()

    def XmlImport(self, Url=None, ImportMap=None, Overwrite=None, Destination=None):
        params = [
            Url if Url is not None else pythoncom.Missing,
            ImportMap if ImportMap is not None else pythoncom.Missing,
            Overwrite if Overwrite is not None else pythoncom.Missing,
            Destination if Destination is not None else pythoncom.Missing,
        ]
        return XlXmlImportResult(self.workbook.XmlImport(*params))

    def XmlImportXml(self, Data=None, ImportMap=None, Overwrite=None, Destination=None):
        params = [
            Data if Data is not None else pythoncom.Missing,
            ImportMap if ImportMap is not None else pythoncom.Missing,
            Overwrite if Overwrite is not None else pythoncom.Missing,
            Destination if Destination is not None else pythoncom.Missing,
        ]
        return XlXmlImportResult(self.workbook.XmlImportXml(*params))


class Workbooks:

    def __init__(self, workbooks=None):
        self.workbooks = workbooks

    def __call__(self, item):
        return Workbook(self.workbooks(item))

    @property
    def Application(self):
        return self.workbooks.Application

    @property
    def Count(self):
        return self.workbooks.Count

    @property
    def Creator(self):
        return self.workbooks.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.workbooks.Item):
            return self.workbooks.Item(*params)
        else:
            return self.workbooks.GetItem(*params)

    @property
    def Parent(self):
        return self.workbooks.Parent

    def Add(self, Template=None):
        params = [
            Template if Template is not None else pythoncom.Missing,
        ]
        return Workbook(self.workbooks.Add(*params))

    def CanCheckOut(self, FileName=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
        ]
        return self.workbooks.CanCheckOut(*params)

    def CheckOut(self, FileName=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
        ]
        self.workbooks.CheckOut(*params)

    def Close(self):
        self.workbooks.Close()

    def Open(self, FileName=None, UpdateLinks=None, ReadOnly=None, Format=None, Password=None, WriteResPassword=None, IgnoreReadOnlyRecommended=None, Origin=None, Delimiter=None, Editable=None, Notify=None, Converter=None, AddToMru=None, Local=None, CorruptLoad=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            UpdateLinks if UpdateLinks is not None else pythoncom.Missing,
            ReadOnly if ReadOnly is not None else pythoncom.Missing,
            Format if Format is not None else pythoncom.Missing,
            Password if Password is not None else pythoncom.Missing,
            WriteResPassword if WriteResPassword is not None else pythoncom.Missing,
            IgnoreReadOnlyRecommended if IgnoreReadOnlyRecommended is not None else pythoncom.Missing,
            Origin if Origin is not None else pythoncom.Missing,
            Delimiter if Delimiter is not None else pythoncom.Missing,
            Editable if Editable is not None else pythoncom.Missing,
            Notify if Notify is not None else pythoncom.Missing,
            Converter if Converter is not None else pythoncom.Missing,
            AddToMru if AddToMru is not None else pythoncom.Missing,
            Local if Local is not None else pythoncom.Missing,
            CorruptLoad if CorruptLoad is not None else pythoncom.Missing,
        ]
        return Workbook(self.workbooks.Open(*params))

    def OpenDatabase(self, FileName=None, CommandText=None, CommandType=None, BackgroundQuery=None, ImportDataAs=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            CommandText if CommandText is not None else pythoncom.Missing,
            CommandType if CommandType is not None else pythoncom.Missing,
            BackgroundQuery if BackgroundQuery is not None else pythoncom.Missing,
            ImportDataAs if ImportDataAs is not None else pythoncom.Missing,
        ]
        return Workbook(self.workbooks.OpenDatabase(*params))

    def OpenText(self, FileName=None, Origin=None, StartRow=None, DataType=None, TextQualifier=None, ConsecutiveDelimiter=None, Tab=None, Semicolon=None, Comma=None, Space=None, Other=None, OtherChar=None, FieldInfo=None, TextVisualLayout=None, DecimalSeparator=None, ThousandsSeparator=None, TrailingMinusNumbers=None, Local=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            Origin if Origin is not None else pythoncom.Missing,
            StartRow if StartRow is not None else pythoncom.Missing,
            DataType if DataType is not None else pythoncom.Missing,
            TextQualifier if TextQualifier is not None else pythoncom.Missing,
            ConsecutiveDelimiter if ConsecutiveDelimiter is not None else pythoncom.Missing,
            Tab if Tab is not None else pythoncom.Missing,
            Semicolon if Semicolon is not None else pythoncom.Missing,
            Comma if Comma is not None else pythoncom.Missing,
            Space if Space is not None else pythoncom.Missing,
            Other if Other is not None else pythoncom.Missing,
            OtherChar if OtherChar is not None else pythoncom.Missing,
            FieldInfo if FieldInfo is not None else pythoncom.Missing,
            TextVisualLayout if TextVisualLayout is not None else pythoncom.Missing,
            DecimalSeparator if DecimalSeparator is not None else pythoncom.Missing,
            ThousandsSeparator if ThousandsSeparator is not None else pythoncom.Missing,
            TrailingMinusNumbers if TrailingMinusNumbers is not None else pythoncom.Missing,
            Local if Local is not None else pythoncom.Missing,
        ]
        self.workbooks.OpenText(*params)

    def OpenXML(self, FileName=None, Stylesheets=None, LoadOption=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            Stylesheets if Stylesheets is not None else pythoncom.Missing,
            LoadOption if LoadOption is not None else pythoncom.Missing,
        ]
        return Workbook(self.workbooks.OpenXML(*params))


class Worksheet:

    def __init__(self, worksheet=None):
        self.worksheet = worksheet

    @property
    def Application(self):
        return self.worksheet.Application

    @property
    def AutoFilter(self):
        return AutoFilter(self.worksheet.AutoFilter)

    @property
    def AutoFilterMode(self):
        return self.worksheet.AutoFilterMode

    @AutoFilterMode.setter
    def AutoFilterMode(self, value):
        self.worksheet.AutoFilterMode = value

    def Cells(self, RowIndex=None, ColumnIndex=None):
        params = [
            RowIndex if RowIndex is not None else pythoncom.Missing,
            ColumnIndex if ColumnIndex is not None else pythoncom.Missing,
        ]
        if callable(self.worksheet.Cells):
            return Range(self.worksheet.Cells(*params))
        else:
            return Range(self.worksheet.GetCells(*params))

    @property
    def CircularReference(self):
        return Range(self.worksheet.CircularReference)

    @property
    def CodeName(self):
        return self.worksheet.CodeName

    @property
    def Columns(self):
        return Range(self.worksheet.Columns)

    @property
    def Comments(self):
        return Comments(self.worksheet.Comments)

    @property
    def CommentsThreaded(self):
        return CommentsThreaded(self.worksheet.CommentsThreaded)

    @property
    def ConsolidationFunction(self):
        return XlConsolidationFunction(self.worksheet.ConsolidationFunction)

    @property
    def ConsolidationOptions(self):
        return self.worksheet.ConsolidationOptions

    @property
    def ConsolidationSources(self):
        return self.worksheet.ConsolidationSources

    @property
    def Creator(self):
        return self.worksheet.Creator

    @property
    def CustomProperties(self):
        return CustomProperties(self.worksheet.CustomProperties)

    @property
    def DisplayPageBreaks(self):
        return self.worksheet.DisplayPageBreaks

    @DisplayPageBreaks.setter
    def DisplayPageBreaks(self, value):
        self.worksheet.DisplayPageBreaks = value

    @property
    def DisplayRightToLeft(self):
        return self.worksheet.DisplayRightToLeft

    @property
    def EnableAutoFilter(self):
        return self.worksheet.EnableAutoFilter

    @EnableAutoFilter.setter
    def EnableAutoFilter(self, value):
        self.worksheet.EnableAutoFilter = value

    @property
    def EnableCalculation(self):
        return self.worksheet.EnableCalculation

    @EnableCalculation.setter
    def EnableCalculation(self, value):
        self.worksheet.EnableCalculation = value

    @property
    def EnableFormatConditionsCalculation(self):
        return self.worksheet.EnableFormatConditionsCalculation

    @EnableFormatConditionsCalculation.setter
    def EnableFormatConditionsCalculation(self, value):
        self.worksheet.EnableFormatConditionsCalculation = value

    @property
    def EnableOutlining(self):
        return self.worksheet.EnableOutlining

    @EnableOutlining.setter
    def EnableOutlining(self, value):
        self.worksheet.EnableOutlining = value

    @property
    def EnablePivotTable(self):
        return self.worksheet.EnablePivotTable

    @EnablePivotTable.setter
    def EnablePivotTable(self, value):
        self.worksheet.EnablePivotTable = value

    @property
    def EnableSelection(self):
        return XlEnableSelection(self.worksheet.EnableSelection)

    @EnableSelection.setter
    def EnableSelection(self, value):
        self.worksheet.EnableSelection = value

    @property
    def FilterMode(self):
        return self.worksheet.FilterMode

    @property
    def HPageBreaks(self):
        return HPageBreaks(self.worksheet.HPageBreaks)

    @property
    def Hyperlinks(self):
        return Hyperlinks(self.worksheet.Hyperlinks)

    @property
    def Index(self):
        return self.worksheet.Index

    @property
    def ListObjects(self):
        return ListObject(self.worksheet.ListObjects)

    @property
    def MailEnvelope(self):
        return self.worksheet.MailEnvelope

    @property
    def Name(self):
        return self.worksheet.Name

    @Name.setter
    def Name(self, value):
        self.worksheet.Name = value

    @property
    def Names(self):
        return Names(self.worksheet.Names)

    @property
    def Next(self):
        return Worksheet(self.worksheet.Next)

    @property
    def Outline(self):
        return Outline(self.worksheet.Outline)

    @property
    def PageSetup(self):
        return PageSetup(self.worksheet.PageSetup)

    @property
    def Parent(self):
        return self.worksheet.Parent

    @property
    def Previous(self):
        return Worksheet(self.worksheet.Previous)

    @property
    def PrintedCommentPages(self):
        return self.worksheet.PrintedCommentPages

    @property
    def ProtectContents(self):
        return self.worksheet.ProtectContents

    @property
    def ProtectDrawingObjects(self):
        return self.worksheet.ProtectDrawingObjects

    @property
    def Protection(self):
        return Protection(self.worksheet.Protection)

    @property
    def ProtectionMode(self):
        return self.worksheet.ProtectionMode

    @property
    def ProtectScenarios(self):
        return self.worksheet.ProtectScenarios

    @property
    def QueryTables(self):
        return QueryTables(self.worksheet.QueryTables)

    def Range(self, Cell1=None, Cell2=None):
        params = [
            Cell1 if Cell1 is not None else pythoncom.Missing,
            Cell2 if Cell2 is not None else pythoncom.Missing,
        ]
        if callable(self.worksheet.Range):
            return Range(self.worksheet.Range(*params))
        else:
            return Range(self.worksheet.GetRange(*params))

    @property
    def Rows(self):
        return Range(self.worksheet.Rows)

    @property
    def ScrollArea(self):
        return self.worksheet.ScrollArea

    @ScrollArea.setter
    def ScrollArea(self, value):
        self.worksheet.ScrollArea = value

    @property
    def Shapes(self):
        return Shapes(self.worksheet.Shapes)

    @property
    def Sort(self):
        return Sort(self.worksheet.Sort)

    @property
    def StandardHeight(self):
        return self.worksheet.StandardHeight

    @property
    def StandardWidth(self):
        return self.worksheet.StandardWidth

    @StandardWidth.setter
    def StandardWidth(self, value):
        self.worksheet.StandardWidth = value

    @property
    def Tab(self):
        return Tab(self.worksheet.Tab)

    @property
    def TransitionExpEval(self):
        return self.worksheet.TransitionExpEval

    @TransitionExpEval.setter
    def TransitionExpEval(self, value):
        self.worksheet.TransitionExpEval = value

    @property
    def TransitionFormEntry(self):
        return self.worksheet.TransitionFormEntry

    @TransitionFormEntry.setter
    def TransitionFormEntry(self, value):
        self.worksheet.TransitionFormEntry = value

    @property
    def Type(self):
        return XlSheetType(self.worksheet.Type)

    @property
    def UsedRange(self):
        return Range(self.worksheet.UsedRange)

    @property
    def Visible(self):
        return XlSheetVisibility(self.worksheet.Visible)

    @Visible.setter
    def Visible(self, value):
        self.worksheet.Visible = value

    @property
    def VPageBreaks(self):
        return VPageBreaks(self.worksheet.VPageBreaks)

    def Activate(self):
        self.worksheet.Activate()

    def Calculate(self):
        self.worksheet.Calculate()

    def ChartObjects(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.worksheet.ChartObjects(*params)

    def CheckSpelling(self, CustomDictionary=None, IgnoreUppercase=None, AlwaysSuggest=None, SpellLang=None):
        params = [
            CustomDictionary if CustomDictionary is not None else pythoncom.Missing,
            IgnoreUppercase if IgnoreUppercase is not None else pythoncom.Missing,
            AlwaysSuggest if AlwaysSuggest is not None else pythoncom.Missing,
            SpellLang if SpellLang is not None else pythoncom.Missing,
        ]
        self.worksheet.CheckSpelling(*params)

    def CircleInvalid(self):
        self.worksheet.CircleInvalid()

    def ClearArrows(self):
        self.worksheet.ClearArrows()

    def ClearCircles(self):
        self.worksheet.ClearCircles()

    def Copy(self, Before=None, After=None):
        params = [
            Before if Before is not None else pythoncom.Missing,
            After if After is not None else pythoncom.Missing,
        ]
        self.worksheet.Copy(*params)

    def Delete(self):
        return self.worksheet.Delete()

    def Evaluate(self, Name=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
        ]
        return self.worksheet.Evaluate(*params)

    def ExportAsFixedFormat(self, Type=None, FileName=None, Quality=None, IncludeDocProperties=None, IgnorePrintAreas=None, From=None, To=None, OpenAfterPublish=None, FixedFormatExtClassPtr=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            FileName if FileName is not None else pythoncom.Missing,
            Quality if Quality is not None else pythoncom.Missing,
            IncludeDocProperties if IncludeDocProperties is not None else pythoncom.Missing,
            IgnorePrintAreas if IgnorePrintAreas is not None else pythoncom.Missing,
            From if From is not None else pythoncom.Missing,
            To if To is not None else pythoncom.Missing,
            OpenAfterPublish if OpenAfterPublish is not None else pythoncom.Missing,
            FixedFormatExtClassPtr if FixedFormatExtClassPtr is not None else pythoncom.Missing,
        ]
        self.worksheet.ExportAsFixedFormat(*params)

    def Move(self, Before=None, After=None):
        params = [
            Before if Before is not None else pythoncom.Missing,
            After if After is not None else pythoncom.Missing,
        ]
        self.worksheet.Move(*params)

    def OLEObjects(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.worksheet.OLEObjects(*params)

    def Paste(self, Destination=None, Link=None):
        params = [
            Destination if Destination is not None else pythoncom.Missing,
            Link if Link is not None else pythoncom.Missing,
        ]
        self.worksheet.Paste(*params)

    def PasteSpecial(self, Format=None, Link=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, NoHTMLFormatting=None):
        params = [
            Format if Format is not None else pythoncom.Missing,
            Link if Link is not None else pythoncom.Missing,
            DisplayAsIcon if DisplayAsIcon is not None else pythoncom.Missing,
            IconFileName if IconFileName is not None else pythoncom.Missing,
            IconIndex if IconIndex is not None else pythoncom.Missing,
            IconLabel if IconLabel is not None else pythoncom.Missing,
            NoHTMLFormatting if NoHTMLFormatting is not None else pythoncom.Missing,
        ]
        self.worksheet.PasteSpecial(*params)

    def PivotTables(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return PivotTable(self.worksheet.PivotTables(*params))

    def PivotTableWizard(self, SourceType=None, SourceData=None, TableDestination=None, TableName=None, RowGrand=None, ColumnGrand=None, SaveData=None, HasAutoFormat=None, AutoPage=None, Reserved=None, BackgroundQuery=None, OptimizeCache=None, PageFieldOrder=None, PageFieldWrapCount=None, ReadData=None, Connection=None):
        params = [
            SourceType if SourceType is not None else pythoncom.Missing,
            SourceData if SourceData is not None else pythoncom.Missing,
            TableDestination if TableDestination is not None else pythoncom.Missing,
            TableName if TableName is not None else pythoncom.Missing,
            RowGrand if RowGrand is not None else pythoncom.Missing,
            ColumnGrand if ColumnGrand is not None else pythoncom.Missing,
            SaveData if SaveData is not None else pythoncom.Missing,
            HasAutoFormat if HasAutoFormat is not None else pythoncom.Missing,
            AutoPage if AutoPage is not None else pythoncom.Missing,
            Reserved if Reserved is not None else pythoncom.Missing,
            BackgroundQuery if BackgroundQuery is not None else pythoncom.Missing,
            OptimizeCache if OptimizeCache is not None else pythoncom.Missing,
            PageFieldOrder if PageFieldOrder is not None else pythoncom.Missing,
            PageFieldWrapCount if PageFieldWrapCount is not None else pythoncom.Missing,
            ReadData if ReadData is not None else pythoncom.Missing,
            Connection if Connection is not None else pythoncom.Missing,
        ]
        return PivotTable(self.worksheet.PivotTableWizard(*params))

    def PrintOut(self, From=None, To=None, Copies=None, Preview=None, ActivePrinter=None, PrintToFile=None, Collate=None, PrToFileName=None, IgnorePrintAreas=None):
        params = [
            From if From is not None else pythoncom.Missing,
            To if To is not None else pythoncom.Missing,
            Copies if Copies is not None else pythoncom.Missing,
            Preview if Preview is not None else pythoncom.Missing,
            ActivePrinter if ActivePrinter is not None else pythoncom.Missing,
            PrintToFile if PrintToFile is not None else pythoncom.Missing,
            Collate if Collate is not None else pythoncom.Missing,
            PrToFileName if PrToFileName is not None else pythoncom.Missing,
            IgnorePrintAreas if IgnorePrintAreas is not None else pythoncom.Missing,
        ]
        return self.worksheet.PrintOut(*params)

    def PrintPreview(self, EnableChanges=None):
        params = [
            EnableChanges if EnableChanges is not None else pythoncom.Missing,
        ]
        self.worksheet.PrintPreview(*params)

    def Protect(self, Password=None, DrawingObjects=None, Contents=None, Scenarios=None, UserInterfaceOnly=None, AllowFormattingCells=None, AllowFormattingColumns=None, AllowFormattingRows=None, AllowInsertingColumns=None, AllowInsertingRows=None, AllowInsertingHyperlinks=None, AllowDeletingColumns=None, AllowDeletingRows=None, AllowSorting=None, AllowFiltering=None, AllowUsingPivotTables=None):
        params = [
            Password if Password is not None else pythoncom.Missing,
            DrawingObjects if DrawingObjects is not None else pythoncom.Missing,
            Contents if Contents is not None else pythoncom.Missing,
            Scenarios if Scenarios is not None else pythoncom.Missing,
            UserInterfaceOnly if UserInterfaceOnly is not None else pythoncom.Missing,
            AllowFormattingCells if AllowFormattingCells is not None else pythoncom.Missing,
            AllowFormattingColumns if AllowFormattingColumns is not None else pythoncom.Missing,
            AllowFormattingRows if AllowFormattingRows is not None else pythoncom.Missing,
            AllowInsertingColumns if AllowInsertingColumns is not None else pythoncom.Missing,
            AllowInsertingRows if AllowInsertingRows is not None else pythoncom.Missing,
            AllowInsertingHyperlinks if AllowInsertingHyperlinks is not None else pythoncom.Missing,
            AllowDeletingColumns if AllowDeletingColumns is not None else pythoncom.Missing,
            AllowDeletingRows if AllowDeletingRows is not None else pythoncom.Missing,
            AllowSorting if AllowSorting is not None else pythoncom.Missing,
            AllowFiltering if AllowFiltering is not None else pythoncom.Missing,
            AllowUsingPivotTables if AllowUsingPivotTables is not None else pythoncom.Missing,
        ]
        self.worksheet.Protect(*params)

    def ResetAllPageBreaks(self):
        self.worksheet.ResetAllPageBreaks()

    def SaveAs(self, FileName=None, FileFormat=None, Password=None, WriteResPassword=None, ReadOnlyRecommended=None, CreateBackup=None, AddToMru=None, TextCodepage=None, TextVisualLayout=None, Local=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            FileFormat if FileFormat is not None else pythoncom.Missing,
            Password if Password is not None else pythoncom.Missing,
            WriteResPassword if WriteResPassword is not None else pythoncom.Missing,
            ReadOnlyRecommended if ReadOnlyRecommended is not None else pythoncom.Missing,
            CreateBackup if CreateBackup is not None else pythoncom.Missing,
            AddToMru if AddToMru is not None else pythoncom.Missing,
            TextCodepage if TextCodepage is not None else pythoncom.Missing,
            TextVisualLayout if TextVisualLayout is not None else pythoncom.Missing,
            Local if Local is not None else pythoncom.Missing,
        ]
        self.worksheet.SaveAs(*params)

    def Scenarios(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.worksheet.Scenarios(*params)

    def Select(self, Replace=None):
        params = [
            Replace if Replace is not None else pythoncom.Missing,
        ]
        self.worksheet.Select(*params)

    def SetBackgroundPicture(self, FileName=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
        ]
        self.worksheet.SetBackgroundPicture(*params)

    def ShowAllData(self):
        self.worksheet.ShowAllData()

    def ShowDataForm(self):
        self.worksheet.ShowDataForm()

    def Unprotect(self, Password=None):
        params = [
            Password if Password is not None else pythoncom.Missing,
        ]
        self.worksheet.Unprotect(*params)

    def XmlDataQuery(self, XPath=None, SelectionNamespaces=None, Map=None):
        params = [
            XPath if XPath is not None else pythoncom.Missing,
            SelectionNamespaces if SelectionNamespaces is not None else pythoncom.Missing,
            Map if Map is not None else pythoncom.Missing,
        ]
        return self.worksheet.XmlDataQuery(*params)

    def XmlMapQuery(self, XPath=None, SelectionNamespaces=None, Map=None):
        params = [
            XPath if XPath is not None else pythoncom.Missing,
            SelectionNamespaces if SelectionNamespaces is not None else pythoncom.Missing,
            Map if Map is not None else pythoncom.Missing,
        ]
        return self.worksheet.XmlMapQuery(*params)


class WorksheetFunction:

    def __init__(self, worksheetfunction=None):
        self.worksheetfunction = worksheetfunction

    @property
    def Application(self):
        return self.worksheetfunction.Application

    @property
    def Creator(self):
        return self.worksheetfunction.Creator

    @property
    def Parent(self):
        return self.worksheetfunction.Parent

    def AccrInt(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.AccrInt(*params)

    def AccrIntM(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.AccrIntM(*params)

    def Acos(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Acos(*params)

    def Acosh(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Acosh(*params)

    def Aggregate(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Aggregate(*params)

    def AmorDegrc(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.AmorDegrc(*params)

    def AmorLinc(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.AmorLinc(*params)

    def And(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.And(*params)

    def Asc(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Asc(*params)

    def Asin(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Asin(*params)

    def Asinh(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Asinh(*params)

    def Atan2(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Atan2(*params)

    def Atanh(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Atanh(*params)

    def AveDev(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.AveDev(*params)

    def Average(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Average(*params)

    def AverageIf(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.AverageIf(*params)

    def AverageIfs(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.AverageIfs(*params)

    def BahtText(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.BahtText(*params)

    def BesselI(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.BesselI(*params)

    def BesselJ(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.BesselJ(*params)

    def BesselK(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.BesselK(*params)

    def BesselY(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.BesselY(*params)

    def BetaDist(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.BetaDist(*params)

    def BetaInv(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.BetaInv(*params)

    def Beta_Dist(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Beta_Dist(*params)

    def Beta_Inv(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Beta_Inv(*params)

    def Bin2Dec(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Bin2Dec(*params)

    def Bin2Hex(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Bin2Hex(*params)

    def Bin2Oct(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Bin2Oct(*params)

    def BinomDist(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.BinomDist(*params)

    def Binom_Dist(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Binom_Dist(*params)

    def Binom_Inv(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Binom_Inv(*params)

    def Ceiling(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Ceiling(*params)

    def Ceiling_Precise(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Ceiling_Precise(*params)

    def ChiDist(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ChiDist(*params)

    def ChiInv(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ChiInv(*params)

    def ChiSq_Dist(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ChiSq_Dist(*params)

    def ChiSq_Dist_RT(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ChiSq_Dist_RT(*params)

    def ChiSq_Inv(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ChiSq_Inv(*params)

    def ChiSq_Inv_RT(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ChiSq_Inv_RT(*params)

    def ChiSq_Test(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ChiSq_Test(*params)

    def ChiTest(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ChiTest(*params)

    def Choose(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Choose(*params)

    def Clean(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Clean(*params)

    def Combin(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Combin(*params)

    def Complex(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Complex(*params)

    def Confidence(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Confidence(*params)

    def Confidence_Norm(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Confidence_Norm(*params)

    def Confidence_T(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Confidence_T(*params)

    def Convert(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Convert(*params)

    def Correl(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Correl(*params)

    def Cosh(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Cosh(*params)

    def Count(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Count(*params)

    def CountA(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.CountA(*params)

    def CountBlank(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.CountBlank(*params)

    def CountIf(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.CountIf(*params)

    def CountIfs(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.CountIfs(*params)

    def CoupDayBs(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.CoupDayBs(*params)

    def CoupDays(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.CoupDays(*params)

    def CoupDaysNc(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.CoupDaysNc(*params)

    def CoupNcd(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.CoupNcd(*params)

    def CoupNum(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.CoupNum(*params)

    def Covar(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Covar(*params)

    def Covariance_P(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Covariance_P(*params)

    def Covariance_S(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Covariance_S(*params)

    def CritBinom(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.CritBinom(*params)

    def CumIPmt(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.CumIPmt(*params)

    def CumPrinc(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.CumPrinc(*params)

    def DAverage(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.DAverage(*params)

    def Days360(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Days360(*params)

    def Db(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Db(*params)

    def Dbcs(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Dbcs(*params)

    def DCount(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.DCount(*params)

    def DCountA(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.DCountA(*params)

    def Ddb(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Ddb(*params)

    def Dec2Bin(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Dec2Bin(*params)

    def Dec2Hex(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Dec2Hex(*params)

    def Dec2Oct(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Dec2Oct(*params)

    def Degrees(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Degrees(*params)

    def Delta(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Delta(*params)

    def DevSq(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.DevSq(*params)

    def DGet(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.DGet(*params)

    def Disc(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Disc(*params)

    def DMax(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.DMax(*params)

    def DMin(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.DMin(*params)

    def Dollar(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Dollar(*params)

    def DollarDe(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.DollarDe(*params)

    def DollarFr(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.DollarFr(*params)

    def DProduct(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.DProduct(*params)

    def DStDev(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.DStDev(*params)

    def DStDevP(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.DStDevP(*params)

    def DSum(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.DSum(*params)

    def Duration(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Duration(*params)

    def DVar(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.DVar(*params)

    def DVarP(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.DVarP(*params)

    def EDate(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.EDate(*params)

    def Effect(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Effect(*params)

    def EoMonth(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.EoMonth(*params)

    def Erf(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Erf(*params)

    def ErfC(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ErfC(*params)

    def ErfC_Precise(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ErfC_Precise(*params)

    def Erf_Precise(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Erf_Precise(*params)

    def Even(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Even(*params)

    def ExponDist(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ExponDist(*params)

    def Expon_Dist(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Expon_Dist(*params)

    def Fact(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Fact(*params)

    def FactDouble(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.FactDouble(*params)

    def FDist(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.FDist(*params)

    def Find(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Find(*params)

    def FindB(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.FindB(*params)

    def FInv(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.FInv(*params)

    def Fisher(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Fisher(*params)

    def FisherInv(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.FisherInv(*params)

    def Fixed(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Fixed(*params)

    def Floor(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Floor(*params)

    def Floor_Precise(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Floor_Precise(*params)

    def Forecast(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Forecast(*params)

    def Frequency(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Frequency(*params)

    def FTest(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.FTest(*params)

    def Fv(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Fv(*params)

    def FVSchedule(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.FVSchedule(*params)

    def F_Dist(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.F_Dist(*params)

    def F_Dist_RT(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.F_Dist_RT(*params)

    def F_Inv(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.F_Inv(*params)

    def F_Inv_RT(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.F_Inv_RT(*params)

    def F_Test(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.F_Test(*params)

    def GammaDist(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.GammaDist(*params)

    def GammaInv(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.GammaInv(*params)

    def GammaLn(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.GammaLn(*params)

    def GammaLn_Precise(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.GammaLn_Precise(*params)

    def Gamma_Dist(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Gamma_Dist(*params)

    def Gamma_Inv(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Gamma_Inv(*params)

    def Gcd(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Gcd(*params)

    def GeoMean(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.GeoMean(*params)

    def GeStep(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.GeStep(*params)

    def Growth(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Growth(*params)

    def HarMean(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.HarMean(*params)

    def Hex2Bin(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Hex2Bin(*params)

    def Hex2Dec(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Hex2Dec(*params)

    def Hex2Oct(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Hex2Oct(*params)

    def HLookup(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.HLookup(*params)

    def HypGeomDist(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.HypGeomDist(*params)

    def HypGeom_Dist(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.HypGeom_Dist(*params)

    def IfError(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.IfError(*params)

    def ImAbs(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ImAbs(*params)

    def Imaginary(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Imaginary(*params)

    def ImArgument(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ImArgument(*params)

    def ImConjugate(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ImConjugate(*params)

    def ImCos(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ImCos(*params)

    def ImDiv(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ImDiv(*params)

    def ImExp(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ImExp(*params)

    def ImLn(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ImLn(*params)

    def ImLog10(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ImLog10(*params)

    def ImLog2(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ImLog2(*params)

    def ImPower(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ImPower(*params)

    def ImProduct(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ImProduct(*params)

    def ImReal(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ImReal(*params)

    def ImSin(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ImSin(*params)

    def ImSqrt(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ImSqrt(*params)

    def ImSub(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ImSub(*params)

    def ImSum(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ImSum(*params)

    def Intercept(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Intercept(*params)

    def IntRate(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.IntRate(*params)

    def Ipmt(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Ipmt(*params)

    def Irr(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Irr(*params)

    def IsErr(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.IsErr(*params)

    def IsError(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.IsError(*params)

    def IsEven(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.IsEven(*params)

    def IsLogical(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.IsLogical(*params)

    def IsNA(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.IsNA(*params)

    def IsNonText(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.IsNonText(*params)

    def IsNumber(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.IsNumber(*params)

    def IsOdd(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.IsOdd(*params)

    def ISO_Ceiling(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ISO_Ceiling(*params)

    def Ispmt(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Ispmt(*params)

    def IsText(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.IsText(*params)

    def Kurt(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Kurt(*params)

    def Large(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Large(*params)

    def Lcm(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Lcm(*params)

    def LinEst(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.LinEst(*params)

    def Ln(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Ln(*params)

    def Log(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Log(*params)

    def Log10(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Log10(*params)

    def LogEst(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.LogEst(*params)

    def LogInv(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.LogInv(*params)

    def LogNormDist(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.LogNormDist(*params)

    def LogNorm_Dist(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.LogNorm_Dist(*params)

    def LogNorm_Inv(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.LogNorm_Inv(*params)

    def Lookup(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Lookup(*params)

    def Match(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Match(*params)

    def Max(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Max(*params)

    def MDeterm(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.MDeterm(*params)

    def MDuration(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.MDuration(*params)

    def Median(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Median(*params)

    def Min(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Min(*params)

    def MInverse(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.MInverse(*params)

    def MIrr(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.MIrr(*params)

    def MMult(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.MMult(*params)

    def Mode(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Mode(*params)

    def Mode_Mult(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Mode_Mult(*params)

    def Mode_Sngl(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Mode_Sngl(*params)

    def MRound(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.MRound(*params)

    def MultiNomial(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.MultiNomial(*params)

    def NegBinomDist(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.NegBinomDist(*params)

    def NegBinom_Dist(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.NegBinom_Dist(*params)

    def NetworkDays(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.NetworkDays(*params)

    def NetworkDays_Intl(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.NetworkDays_Intl(*params)

    def Nominal(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Nominal(*params)

    def NormDist(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.NormDist(*params)

    def NormInv(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.NormInv(*params)

    def NormSDist(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.NormSDist(*params)

    def NormSInv(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.NormSInv(*params)

    def Norm_Dist(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Norm_Dist(*params)

    def Norm_Inv(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Norm_Inv(*params)

    def Norm_S_Dist(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Norm_S_Dist(*params)

    def Norm_S_Inv(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Norm_S_Inv(*params)

    def NPer(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.NPer(*params)

    def Npv(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Npv(*params)

    def Oct2Bin(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Oct2Bin(*params)

    def Oct2Dec(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Oct2Dec(*params)

    def Oct2Hex(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Oct2Hex(*params)

    def Odd(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Odd(*params)

    def OddFPrice(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.OddFPrice(*params)

    def OddFYield(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.OddFYield(*params)

    def OddLPrice(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.OddLPrice(*params)

    def OddLYield(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.OddLYield(*params)

    def Or(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Or(*params)

    def Pearson(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Pearson(*params)

    def Percentile(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Percentile(*params)

    def Percentile_Exc(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Percentile_Exc(*params)

    def Percentile_Inc(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Percentile_Inc(*params)

    def PercentRank(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.PercentRank(*params)

    def PercentRank_Exc(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.PercentRank_Exc(*params)

    def PercentRank_Inc(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.PercentRank_Inc(*params)

    def Permut(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Permut(*params)

    def Phonetic(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Phonetic(*params)

    def Pi(self):
        return self.worksheetfunction.Pi()

    def Pmt(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Pmt(*params)

    def Poisson(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Poisson(*params)

    def Poisson_Dist(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Poisson_Dist(*params)

    def Power(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Power(*params)

    def Ppmt(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Ppmt(*params)

    def Price(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Price(*params)

    def PriceDisc(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.PriceDisc(*params)

    def PriceMat(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.PriceMat(*params)

    def Prob(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Prob(*params)

    def Product(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Product(*params)

    def Proper(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Proper(*params)

    def Pv(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Pv(*params)

    def Quartile(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Quartile(*params)

    def Quartile_Exc(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Quartile_Exc(*params)

    def Quartile_Inc(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Quartile_Inc(*params)

    def Quotient(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Quotient(*params)

    def Radians(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Radians(*params)

    def RandBetween(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.RandBetween(*params)

    def Rank(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Rank(*params)

    def Rank_Avg(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Rank_Avg(*params)

    def Rank_Eq(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Rank_Eq(*params)

    def Rate(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Rate(*params)

    def Received(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Received(*params)

    def Replace(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Replace(*params)

    def ReplaceB(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ReplaceB(*params)

    def Rept(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Rept(*params)

    def Roman(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Roman(*params)

    def Round(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Round(*params)

    def RoundDown(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.RoundDown(*params)

    def RoundUp(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.RoundUp(*params)

    def RSq(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.RSq(*params)

    def RTD(self, progID=None, server=None, topic1=None, topic2=None, topic3=None, topic4=None, topic5=None, topic6=None, topic7=None, topic8=None, topic9=None, topic10=None, topic11=None, topic12=None, topic13=None, topic14=None, topic15=None, topic16=None, topic17=None, topic18=None, topic19=None, topic20=None, topic21=None, topic22=None, topic23=None, topic24=None, topic25=None, topic26=None, topic27=None, topic28=None):
        params = [
            progID if progID is not None else pythoncom.Missing,
            server if server is not None else pythoncom.Missing,
            topic1 if topic1 is not None else pythoncom.Missing,
            topic2 if topic2 is not None else pythoncom.Missing,
            topic3 if topic3 is not None else pythoncom.Missing,
            topic4 if topic4 is not None else pythoncom.Missing,
            topic5 if topic5 is not None else pythoncom.Missing,
            topic6 if topic6 is not None else pythoncom.Missing,
            topic7 if topic7 is not None else pythoncom.Missing,
            topic8 if topic8 is not None else pythoncom.Missing,
            topic9 if topic9 is not None else pythoncom.Missing,
            topic10 if topic10 is not None else pythoncom.Missing,
            topic11 if topic11 is not None else pythoncom.Missing,
            topic12 if topic12 is not None else pythoncom.Missing,
            topic13 if topic13 is not None else pythoncom.Missing,
            topic14 if topic14 is not None else pythoncom.Missing,
            topic15 if topic15 is not None else pythoncom.Missing,
            topic16 if topic16 is not None else pythoncom.Missing,
            topic17 if topic17 is not None else pythoncom.Missing,
            topic18 if topic18 is not None else pythoncom.Missing,
            topic19 if topic19 is not None else pythoncom.Missing,
            topic20 if topic20 is not None else pythoncom.Missing,
            topic21 if topic21 is not None else pythoncom.Missing,
            topic22 if topic22 is not None else pythoncom.Missing,
            topic23 if topic23 is not None else pythoncom.Missing,
            topic24 if topic24 is not None else pythoncom.Missing,
            topic25 if topic25 is not None else pythoncom.Missing,
            topic26 if topic26 is not None else pythoncom.Missing,
            topic27 if topic27 is not None else pythoncom.Missing,
            topic28 if topic28 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.RTD(*params)

    def Search(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Search(*params)

    def SearchB(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.SearchB(*params)

    def SeriesSum(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.SeriesSum(*params)

    def Sinh(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Sinh(*params)

    def Skew(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Skew(*params)

    def Sln(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Sln(*params)

    def Slope(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Slope(*params)

    def Small(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Small(*params)

    def SqrtPi(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.SqrtPi(*params)

    def Standardize(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Standardize(*params)

    def StDev(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.StDev(*params)

    def StDevP(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.StDevP(*params)

    def StDev_P(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.StDev_P(*params)

    def StDev_S(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.StDev_S(*params)

    def StEyx(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.StEyx(*params)

    def Substitute(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Substitute(*params)

    def Subtotal(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Subtotal(*params)

    def Sum(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Sum(*params)

    def SumIf(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.SumIf(*params)

    def SumIfs(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.SumIfs(*params)

    def SumProduct(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.SumProduct(*params)

    def SumSq(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.SumSq(*params)

    def SumX2MY2(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.SumX2MY2(*params)

    def SumX2PY2(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.SumX2PY2(*params)

    def SumXMY2(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.SumXMY2(*params)

    def Syd(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Syd(*params)

    def Tanh(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Tanh(*params)

    def TBillEq(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.TBillEq(*params)

    def TBillPrice(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.TBillPrice(*params)

    def TBillYield(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.TBillYield(*params)

    def TDist(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.TDist(*params)

    def Text(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Text(*params)

    def TInv(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.TInv(*params)

    def Transpose(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Transpose(*params)

    def Trend(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Trend(*params)

    def Trim(self, Arg1=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Trim(*params)

    def TrimMean(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.TrimMean(*params)

    def TTest(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.TTest(*params)

    def T_Dist(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.T_Dist(*params)

    def T_Dist_2T(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.T_Dist_2T(*params)

    def T_Dist_RT(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.T_Dist_RT(*params)

    def T_Inv(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.T_Inv(*params)

    def T_Inv_2T(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.T_Inv_2T(*params)

    def T_Test(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.T_Test(*params)

    def USDollar(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.USDollar(*params)

    def Var(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Var(*params)

    def VarP(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.VarP(*params)

    def Var_P(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Var_P(*params)

    def Var_S(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
            Arg8 if Arg8 is not None else pythoncom.Missing,
            Arg9 if Arg9 is not None else pythoncom.Missing,
            Arg10 if Arg10 is not None else pythoncom.Missing,
            Arg11 if Arg11 is not None else pythoncom.Missing,
            Arg12 if Arg12 is not None else pythoncom.Missing,
            Arg13 if Arg13 is not None else pythoncom.Missing,
            Arg14 if Arg14 is not None else pythoncom.Missing,
            Arg15 if Arg15 is not None else pythoncom.Missing,
            Arg16 if Arg16 is not None else pythoncom.Missing,
            Arg17 if Arg17 is not None else pythoncom.Missing,
            Arg18 if Arg18 is not None else pythoncom.Missing,
            Arg19 if Arg19 is not None else pythoncom.Missing,
            Arg20 if Arg20 is not None else pythoncom.Missing,
            Arg21 if Arg21 is not None else pythoncom.Missing,
            Arg22 if Arg22 is not None else pythoncom.Missing,
            Arg23 if Arg23 is not None else pythoncom.Missing,
            Arg24 if Arg24 is not None else pythoncom.Missing,
            Arg25 if Arg25 is not None else pythoncom.Missing,
            Arg26 if Arg26 is not None else pythoncom.Missing,
            Arg27 if Arg27 is not None else pythoncom.Missing,
            Arg28 if Arg28 is not None else pythoncom.Missing,
            Arg29 if Arg29 is not None else pythoncom.Missing,
            Arg30 if Arg30 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Var_S(*params)

    def Vdb(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
            Arg7 if Arg7 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Vdb(*params)

    def VLookup(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.VLookup(*params)

    def Weekday(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Weekday(*params)

    def WeekNum(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.WeekNum(*params)

    def Weibull(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Weibull(*params)

    def Weibull_Dist(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Weibull_Dist(*params)

    def WorkDay(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.WorkDay(*params)

    def WorkDay_Intl(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.WorkDay_Intl(*params)

    def Xirr(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Xirr(*params)

    def Xnpv(self, Arg1=None, Arg2=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Xnpv(*params)

    def YearFrac(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.YearFrac(*params)

    def YieldDisc(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.YieldDisc(*params)

    def YieldMat(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
            Arg4 if Arg4 is not None else pythoncom.Missing,
            Arg5 if Arg5 is not None else pythoncom.Missing,
            Arg6 if Arg6 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.YieldMat(*params)

    def ZTest(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.ZTest(*params)

    def Z_Test(self, Arg1=None, Arg2=None, Arg3=None):
        params = [
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
            Arg3 if Arg3 is not None else pythoncom.Missing,
        ]
        return self.worksheetfunction.Z_Test(*params)


class Worksheets:

    def __init__(self, worksheets=None):
        self.worksheets = worksheets

    def __call__(self, item):
        return Worksheet(self.worksheets(item))

    @property
    def Application(self):
        return self.worksheets.Application

    @property
    def Count(self):
        return self.worksheets.Count

    @property
    def Creator(self):
        return self.worksheets.Creator

    @property
    def HPageBreaks(self):
        return HPageBreaks(self.worksheets.HPageBreaks)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.worksheets.Item):
            return self.worksheets.Item(*params)
        else:
            return self.worksheets.GetItem(*params)

    @property
    def Parent(self):
        return self.worksheets.Parent

    @property
    def Visible(self):
        return self.worksheets.Visible

    @Visible.setter
    def Visible(self, value):
        self.worksheets.Visible = value

    @property
    def VPageBreaks(self):
        return VPageBreaks(self.worksheets.VPageBreaks)

    def Add(self, Before=None, After=None, Count=None, Type=None):
        params = [
            Before if Before is not None else pythoncom.Missing,
            After if After is not None else pythoncom.Missing,
            Count if Count is not None else pythoncom.Missing,
            Type if Type is not None else pythoncom.Missing,
        ]
        return Worksheet(self.worksheets.Add(*params))

    def Copy(self, Before=None, After=None):
        params = [
            Before if Before is not None else pythoncom.Missing,
            After if After is not None else pythoncom.Missing,
        ]
        self.worksheets.Copy(*params)

    def Delete(self):
        self.worksheets.Delete()

    def FillAcrossSheets(self, Range=None, Type=None):
        params = [
            Range if Range is not None else pythoncom.Missing,
            Type if Type is not None else pythoncom.Missing,
        ]
        self.worksheets.FillAcrossSheets(*params)

    def Move(self, Before=None, After=None):
        params = [
            Before if Before is not None else pythoncom.Missing,
            After if After is not None else pythoncom.Missing,
        ]
        self.worksheets.Move(*params)

    def PrintOut(self, From=None, To=None, Copies=None, Preview=None, ActivePrinter=None, PrintToFile=None, Collate=None, PrToFileName=None, IgnorePrintAreas=None):
        params = [
            From if From is not None else pythoncom.Missing,
            To if To is not None else pythoncom.Missing,
            Copies if Copies is not None else pythoncom.Missing,
            Preview if Preview is not None else pythoncom.Missing,
            ActivePrinter if ActivePrinter is not None else pythoncom.Missing,
            PrintToFile if PrintToFile is not None else pythoncom.Missing,
            Collate if Collate is not None else pythoncom.Missing,
            PrToFileName if PrToFileName is not None else pythoncom.Missing,
            IgnorePrintAreas if IgnorePrintAreas is not None else pythoncom.Missing,
        ]
        return self.worksheets.PrintOut(*params)

    def PrintPreview(self, EnableChanges=None):
        params = [
            EnableChanges if EnableChanges is not None else pythoncom.Missing,
        ]
        self.worksheets.PrintPreview(*params)

    def Select(self, Replace=None):
        params = [
            Replace if Replace is not None else pythoncom.Missing,
        ]
        self.worksheets.Select(*params)


class WorksheetView:

    def __init__(self, worksheetview=None):
        self.worksheetview = worksheetview

    @property
    def Application(self):
        return self.worksheetview.Application

    @property
    def Creator(self):
        return self.worksheetview.Creator

    @property
    def DisplayFormulas(self):
        return self.worksheetview.DisplayFormulas

    @DisplayFormulas.setter
    def DisplayFormulas(self, value):
        self.worksheetview.DisplayFormulas = value

    @property
    def DisplayGridlines(self):
        return self.worksheetview.DisplayGridlines

    @DisplayGridlines.setter
    def DisplayGridlines(self, value):
        self.worksheetview.DisplayGridlines = value

    @property
    def DisplayHeadings(self):
        return self.worksheetview.DisplayHeadings

    @DisplayHeadings.setter
    def DisplayHeadings(self, value):
        self.worksheetview.DisplayHeadings = value

    @property
    def DisplayOutline(self):
        return self.worksheetview.DisplayOutline

    @DisplayOutline.setter
    def DisplayOutline(self, value):
        self.worksheetview.DisplayOutline = value

    @property
    def DisplayZeros(self):
        return self.worksheetview.DisplayZeros

    @DisplayZeros.setter
    def DisplayZeros(self, value):
        self.worksheetview.DisplayZeros = value

    @property
    def Parent(self):
        return self.worksheetview.Parent

    @property
    def Sheet(self):
        return WorksheetView(self.worksheetview.Sheet)


# xlAboveBelow enumeration
xlAboveAverage = 0
xlAboveStdDev = 4
xlBelowAverage = 1
xlBelowStdDev = 5
xlEqualAboveAverage = 2
xlEqualBelowAverage = 3

# xlActionType enumeration
xlActionTypeDrillthrough = 256
xlActionTypeReport = 128
xlActionTypeRowset = 16
xlActionTypeUrl = 1

# xlAllocation enumeration
xlAutomaticAllocation = 2
xlManualAllocation = 1

# xlAllocationMethod enumeration
xlEqualAllocation = 1
xlWeightedAllocation = 2

# xlAllocationValue enumeration
xlAllocateIncrement = 2
xlAllocateValue = 1

# xlApplicationInternational enumeration
xl24HourClock = 33
xl4DigitYears = 43
xlAlternateArraySeparator = 16
xlColumnSeparator = 14
xlCountryCode = 1
xlCountrySetting = 2
xlCurrencyBefore = 37
xlCurrencyCode = 25
xlCurrencyDigits = 27
xlCurrencyLeadingZeros = 40
xlCurrencyMinusSign = 38
xlCurrencyNegative = 28
xlCurrencySpaceBefore = 36
xlCurrencyTrailingZeros = 39
xlDateOrder = 32
xlDateSeparator = 17
xlDayCode = 21
xlDayLeadingZero = 42
xlDecimalSeparator = 3
xlGeneralFormatName = 26
xlHourCode = 22
xlLeftBrace = 12
xlLeftBracket = 10
xlListSeparator = 5
xlLowerCaseColumnLetter = 9
xlLowerCaseRowLetter = 8
xlMDY = 44
xlMetric = 35
xlMinuteCode = 23
xlMonthCode = 20
xlMonthLeadingZero = 41
xlMonthNameChars = 30
xlNoncurrencyDigits = 29
xlNonEnglishFunctions = 34
xlRightBrace = 13
xlRightBracket = 11
xlRowSeparator = 15
xlSecondCode = 24
xlThousandsSeparator = 4
xlTimeLeadingZero = 45
xlTimeSeparator = 18
xlUpperCaseColumnLetter = 7
xlUpperCaseRowLetter = 6
xlWeekdayNameChars = 31
xlYearCode = 19

# xlApplyNamesOrder enumeration
xlColumnThenRow = 2
xlRowThenColumn = 1

# xlArabicModes enumeration
xlArabicBothStrict = 3
xlArabicNone = 0
xlArabicStrictAlefHamza = 1
xlArabicStrictFinalYaa = 2

# xlArrangeStyle enumeration
xlArrangeStyleCascade = 7
xlArrangeStyleHorizontal = -4128
xlArrangeStyleTiled = 1
xlArrangeStyleVertical = -4166

# xlArrowHeadLength enumeration
xlArrowHeadLengthLong = 3
xlArrowHeadLengthMedium = -4138
xlArrowHeadLengthShort = 1

# xlArrowHeadStyle enumeration
xlArrowHeadStyleClosed = 3
xlArrowHeadStyleDoubleClosed = 5
xlArrowHeadStyleDoubleOpen = 4
xlArrowHeadStyleNone = -4142
xlArrowHeadStyleOpen = 2

# xlArrowHeadWidth enumeration
xlArrowHeadWidthMedium = -4138
xlArrowHeadWidthNarrow = 1
xlArrowHeadWidthWide = 3

# xlAutoFillType enumeration
xlFillCopy = 1
xlFillDays = 5
xlFillDefault = 0
xlFillFormats = 3
xlFillMonths = 7
xlFillSeries = 2
xlFillValues = 4
xlFillWeekdays = 6
xlFillYears = 8
xlGrowthTrend = 10
xlLinearTrend = 9
xlFlashFill = 11

# xlAutoFilterOperator enumeration
xlAnd = 1
xlBottom10Items = 4
xlBottom10Percent = 6
xlFilterCellColor = 8
xlFilterDynamic = 11
xlFilterFontColor = 9
xlFilterIcon = 10
xlFilterValues = 7
xlOr = 2
xlTop10Items = 3
xlTop10Percent = 5

# xlAxisCrosses enumeration
xlAxisCrossesAutomatic = -4105
xlAxisCrossesCustom = -4114
xlAxisCrossesMaximum = 2
xlAxisCrossesMinimum = 4

# xlAxisGroup enumeration
xlPrimary = 1
xlSecondary = 2

# xlAxisType enumeration
xlCategory = 1
xlSeriesAxis = 3
xlValue = 2

# xlBackground enumeration
xlBackgroundAutomatic = -4105
xlBackgroundOpaque = 3
xlBackgroundTransparent = 2

# xlBarShape enumeration
xlBox = 0
xlConeToMax = 5
xlConeToPoint = 4
xlCylinder = 3
xlPyramidToMax = 2
xlPyramidToPoint = 1

# xlBordersIndex enumeration
xlDiagonalDown = 5
xlDiagonalUp = 6
xlEdgeBottom = 9
xlEdgeLeft = 7
xlEdgeRight = 10
xlEdgeTop = 8
xlInsideHorizontal = 12
xlInsideVertical = 11

# xlBorderWeight enumeration
xlHairline = 1
xlMedium = -4138
xlThick = 4
xlThin = 2

# xlBuiltInDialog enumeration
xlDialogActivate = 103
xlDialogActiveCellFont = 476
xlDialogAddChartAutoformat = 390
xlDialogAddinManager = 321
xlDialogAlignment = 43
xlDialogApplyNames = 133
xlDialogApplyStyle = 212
xlDialogAppMove = 170
xlDialogAppSize = 171
xlDialogArrangeAll = 12
xlDialogAssignToObject = 213
xlDialogAssignToTool = 293
xlDialogAttachText = 80
xlDialogAttachToolbars = 323
xlDialogAutoCorrect = 485
xlDialogAxes = 78
xlDialogBorder = 45
xlDialogCalculation = 32
xlDialogCellProtection = 46
xlDialogChangeLink = 166
xlDialogChartAddData = 392
xlDialogChartLocation = 527
xlDialogChartOptionsDataLabelMultiple = 724
xlDialogChartOptionsDataLabels = 505
xlDialogChartOptionsDataTable = 506
xlDialogChartSourceData = 540
xlDialogChartTrend = 350
xlDialogChartType = 526
xlDialogChartWizard = 288
xlDialogCheckboxProperties = 435
xlDialogClear = 52
xlDialogColorPalette = 161
xlDialogColumnWidth = 47
xlDialogCombination = 73
xlDialogConditionalFormatting = 583
xlDialogConsolidate = 191
xlDialogCopyChart = 147
xlDialogCopyPicture = 108
xlDialogCreateList = 796
xlDialogCreateNames = 62
xlDialogCreatePublisher = 217
xlDialogCreateRelationship = 1272
xlDialogCustomizeToolbar = 276
xlDialogCustomViews = 493
xlDialogDataDelete = 36
xlDialogDataLabel = 379
xlDialogDataLabelMultiple = 723
xlDialogDataSeries = 40
xlDialogDataValidation = 525
xlDialogDefineName = 61
xlDialogDefineStyle = 229
xlDialogDeleteFormat = 111
xlDialogDeleteName = 110
xlDialogDemote = 203
xlDialogDisplay = 27
xlDialogDocumentInspector = 862
xlDialogEditboxProperties = 438
xlDialogEditColor = 223
xlDialogEditDelete = 54
xlDialogEditionOptions = 251
xlDialogEditSeries = 228
xlDialogErrorbarX = 463
xlDialogErrorbarY = 464
xlDialogErrorChecking = 732
xlDialogEvaluateFormula = 709
xlDialogExternalDataProperties = 530
xlDialogExtract = 35
xlDialogFileDelete = 6
xlDialogFileSharing = 481
xlDialogFillGroup = 200
xlDialogFillWorkgroup = 301
xlDialogFilter = 447
xlDialogFilterAdvanced = 370
xlDialogFindFile = 475
xlDialogFont = 26
xlDialogFontProperties = 381
xlDialogFormatAuto = 269
xlDialogFormatChart = 465
xlDialogFormatCharttype = 423
xlDialogFormatFont = 150
xlDialogFormatLegend = 88
xlDialogFormatMain = 225
xlDialogFormatMove = 128
xlDialogFormatNumber = 42
xlDialogFormatOverlay = 226
xlDialogFormatSize = 129
xlDialogFormatText = 89
xlDialogFormulaFind = 64
xlDialogFormulaGoto = 63
xlDialogFormulaReplace = 130
xlDialogFunctionWizard = 450
xlDialogGallery3dArea = 193
xlDialogGallery3dBar = 272
xlDialogGallery3dColumn = 194
xlDialogGallery3dLine = 195
xlDialogGallery3dPie = 196
xlDialogGallery3dSurface = 273
xlDialogGalleryArea = 67
xlDialogGalleryBar = 68
xlDialogGalleryColumn = 69
xlDialogGalleryCustom = 388
xlDialogGalleryDoughnut = 344
xlDialogGalleryLine = 70
xlDialogGalleryPie = 71
xlDialogGalleryRadar = 249
xlDialogGalleryScatter = 72
xlDialogGoalSeek = 198
xlDialogGridlines = 76
xlDialogImportTextFile = 666
xlDialogInsert = 55
xlDialogInsertHyperlink = 596
xlDialogInsertObject = 259
xlDialogInsertPicture = 342
xlDialogInsertTitle = 380
xlDialogLabelProperties = 436
xlDialogListboxProperties = 437
xlDialogMacroOptions = 382
xlDialogMailEditMailer = 470
xlDialogMailLogon = 339
xlDialogMailNextLetter = 378
xlDialogMainChart = 85
xlDialogMainChartType = 185
xlDialogManageRelationships = 1271
xlDialogMenuEditor = 322
xlDialogMove = 262
xlDialogMyPermission = 834
xlDialogNameManager = 977
xlDialogNew = 119
xlDialogNewName = 978
xlDialogNewWebQuery = 667
xlDialogNote = 154
xlDialogObjectProperties = 207
xlDialogObjectProtection = 214
xlDialogOpen = 1
xlDialogOpenLinks = 2
xlDialogOpenMail = 188
xlDialogOpenText = 441
xlDialogOptionsCalculation = 318
xlDialogOptionsChart = 325
xlDialogOptionsEdit = 319
xlDialogOptionsGeneral = 356
xlDialogOptionsListsAdd = 458
xlDialogOptionsME = 647
xlDialogOptionsTransition = 355
xlDialogOptionsView = 320
xlDialogOutline = 142
xlDialogOverlay = 86
xlDialogOverlayChartType = 186
xlDialogPageSetup = 7
xlDialogParse = 91
xlDialogPasteNames = 58
xlDialogPasteSpecial = 53
xlDialogPatterns = 84
xlDialogPermission = 832
xlDialogPhonetic = 656
xlDialogPivotCalculatedField = 570
xlDialogPivotCalculatedItem = 572
xlDialogPivotClientServerSet = 689
xlDialogPivotFieldGroup = 433
xlDialogPivotFieldProperties = 313
xlDialogPivotFieldUngroup = 434
xlDialogPivotShowPages = 421
xlDialogPivotSolveOrder = 568
xlDialogPivotTableOptions = 567
xlDialogPivotTableSlicerConnections = 1183
xlDialogPivotTableWhatIfAnalysisSettings = 1153
xlDialogPivotTableWizard = 312
xlDialogPlacement = 300
xlDialogPrint = 8
xlDialogPrinterSetup = 9
xlDialogPrintPreview = 222
xlDialogPromote = 202
xlDialogProperties = 474
xlDialogPropertyFields = 754
xlDialogProtectDocument = 28
xlDialogProtectSharing = 620
xlDialogPublishAsWebPage = 653
xlDialogPushbuttonProperties = 445
xlDialogRecommendedPivotTables = 1258
xlDialogReplaceFont = 134
xlDialogRoutingSlip = 336
xlDialogRowHeight = 127
xlDialogRun = 17
xlDialogSaveAs = 5
xlDialogSaveCopyAs = 456
xlDialogSaveNewObject = 208
xlDialogSaveWorkbook = 145
xlDialogSaveWorkspace = 285
xlDialogScale = 87
xlDialogScenarioAdd = 307
xlDialogScenarioCells = 305
xlDialogScenarioEdit = 308
xlDialogScenarioMerge = 473
xlDialogScenarioSummary = 311
xlDialogScrollbarProperties = 420
xlDialogSearch = 731
xlDialogSelectSpecial = 132
xlDialogSendMail = 189
xlDialogSeriesAxes = 460
xlDialogSeriesOptions = 557
xlDialogSeriesOrder = 466
xlDialogSeriesShape = 504
xlDialogSeriesX = 461
xlDialogSeriesY = 462
xlDialogSetBackgroundPicture = 509
xlDialogSetManager = 1109
xlDialogSetMDXEditor = 1208
xlDialogSetPrintTitles = 23
xlDialogSetTupleEditorOnColumns = 1108
xlDialogSetTupleEditorOnRows = 1107
xlDialogSetUpdateStatus = 159
xlDialogShowDetail = 204
xlDialogShowToolbar = 220
xlDialogSize = 261
xlDialogSlicerCreation = 1182
xlDialogSlicerPivotTableConnections = 1184
xlDialogSlicerSettings = 1179
xlDialogSort = 39
xlDialogSortSpecial = 192
xlDialogSparklineInsertColumn = 1134
xlDialogSparklineInsertLine = 1133
xlDialogSparklineInsertWinLoss = 1135
xlDialogSplit = 137
xlDialogStandardFont = 190
xlDialogStandardWidth = 472
xlDialogStyle = 44
xlDialogSubscribeTo = 218
xlDialogSubtotalCreate = 398
xlDialogSummaryInfo = 474
xlDialogTable = 41
xlDialogTabOrder = 394
xlDialogTextToColumns = 422
xlDialogUnhide = 94
xlDialogUpdateLink = 201
xlDialogVbaInsertFile = 328
xlDialogVbaMakeAddin = 478
xlDialogVbaProcedureDefinition = 330
xlDialogView3d = 197
xlDialogWebOptionsBrowsers = 773
xlDialogWebOptionsEncoding = 686
xlDialogWebOptionsFiles = 684
xlDialogWebOptionsFonts = 687
xlDialogWebOptionsGeneral = 683
xlDialogWebOptionsPictures = 685
xlDialogWindowMove = 14
xlDialogWindowSize = 13
xlDialogWorkbookAdd = 281
xlDialogWorkbookCopy = 283
xlDialogWorkbookInsert = 354
xlDialogWorkbookMove = 282
xlDialogWorkbookName = 386
xlDialogWorkbookNew = 302
xlDialogWorkbookOptions = 284
xlDialogWorkbookProtect = 417
xlDialogWorkbookTabSplit = 415
xlDialogWorkbookUnhide = 384
xlDialogWorkgroup = 199
xlDialogWorkspace = 95
xlDialogZoom = 256

# xlCalcFor enumeration
xlAllValues = 0
xlColGroups = 2
xlRowGroups = 1

# xlCalculatedMemberType enumeration
xlCalculatedMeasure = 2
xlCalculatedMember = 0
xlCalculatedSet = 1

# xlCalculation enumeration
xlCalculationAutomatic = -4105
xlCalculationManual = -4135
xlCalculationSemiautomatic = 2

# xlCalculationInterruptKey enumeration
xlAnyKey = 2
xlEscKey = 1
xlNoKey = 0

# xlCalculationState enumeration
xlCalculating = 1
xlDone = 0
xlPending = 2

# xlCategoryType enumeration
xlAutomaticScale = -4105
xlCategoryScale = 2
xlTimeScale = 3

# xlCellChangedState enumeration
xlCellChangeApplied = 3
xlCellChanged = 2
xlCellNotChanged = 1

# xlCellInsertionMode enumeration
xlInsertDeleteCells = 1
xlInsertEntireRows = 2
xlOverwriteCells = 0

# xlCellType enumeration
xlCellTypeAllFormatConditions = -4172
xlCellTypeAllValidation = -4174
xlCellTypeBlanks = 4
xlCellTypeComments = -4144
xlCellTypeConstants = 2
xlCellTypeFormulas = -4123
xlCellTypeLastCell = 11
xlCellTypeSameFormatConditions = -4173
xlCellTypeSameValidation = -4175
xlCellTypeVisible = 12

# xlChartElementPosition enumeration
xlChartElementPositionAutomatic = -4105
xlChartElementPositionCustom = -4114

# xlChartGallery enumeration
xlAnyGallery = 23
xlBuiltIn = 21
xlUserDefined = 22

# xlChartItem enumeration
xlAxis = 21
xlAxisTitle = 17
xlChartArea = 2
xlChartTitle = 4
xlCorners = 6
xlDataLabel = 0
xlDataTable = 7
xlDisplayUnitLabel = 30
xlDownBars = 20
xlDropLines = 26
xlErrorBars = 9
xlFloor = 23
xlHiLoLines = 25
xlLeaderLines = 29
xlLegend = 24
xlLegendEntry = 12
xlLegendKey = 13
xlMajorGridlines = 15
xlMinorGridlines = 16
xlNothing = 28
xlPivotChartDropZone = 32
xlPivotChartFieldButton = 31
xlPlotArea = 19
xlRadarAxisLabels = 27
xlSeries = 3
xlSeriesLines = 22
xlShape = 14
xlTrendline = 8
xlUpBars = 18
xlWalls = 5
xlXErrorBars = 10
xlYErrorBars = 11

# XlChartLocation enumeration
xlLocationAsNewSheet = 1
xlLocationAsObject = 2
xlLocationAutomatic = 3

# XlChartPicturePlacement enumeration
xlAllFaces = 7
xlEnd = 2
xlEndSides = 3
xlFront = 4
xlFrontEnd = 6
xlFrontSides = 5
xlSides = 1

# XlChartPictureType enumeration
xlStack = 2
xlStackScale = 3
xlStretch = 1

# XlChartSplitType enumeration
xlSplitByCustomSplit = 4
xlSplitByPercentValue = 3
xlSplitByPosition = 1
xlSplitByValue = 2

# XlChartType enumeration
xl3DArea = -4098
xl3DAreaStacked = 78
xl3DAreaStacked100 = 79
xl3DBarClustered = 60
xl3DBarStacked = 61
xl3DBarStacked100 = 62
xl3DColumn = -4100
xl3DColumnClustered = 54
xl3DColumnStacked = 55
xl3DColumnStacked100 = 56
xl3DLine = -4101
xl3DPie = -4102
xl3DPieExploded = 70
xlArea = 1
xlAreaEx = 135
xlAreaStacked = 76
xlAreaStacked100 = 77
xlAreaStacked100Ex = 137
xlAreaStackedEx = 136
xlBarClustered = 57
xlBarClusteredEx = 132
xlBarOfPie = 71
xlBarStacked = 58
xlBarStacked100 = 59
xlBarStacked100Ex = 134
xlBarStackedEx = 133
xlBoxwhisker = 121
xlBubble = 15
xlBubble3DEffect = 87
xlBubbleEx = 139
xlColumnClustered = 51
xlColumnClusteredEx = 124
xlColumnStacked = 52
xlColumnStacked100 = 53
xlColumnStacked100Ex = 126
xlColumnStackedEx = 125
xlCombo = -4152
xlComboAreaStackedColumnClustered = 115
xlComboColumnClusteredLine = 113
xlComboColumnClusteredLineSecondaryAxis = 114
xlConeBarClustered = 102
xlConeBarStacked = 103
xlConeBarStacked100 = 104
xlConeCol = 105
xlConeColClustered = 99
xlConeColStacked = 100
xlConeColStacked100 = 101
xlCylinderBarClustered = 95
xlCylinderBarStacked = 96
xlCylinderBarStacked100 = 97
xlCylinderCol = 98
xlCylinderColClustered = 92
xlCylinderColStacked = 93
xlCylinderColStacked100 = 94
xlDoughnut = -4120
xlDoughnutEx = 131
xlDoughnutExploded = 80
xlFunnel = 123
xlHistogram = 118
xlLine = 4
xlLineEx = 127
xlLineMarkers = 65
xlLineMarkersStacked = 66
xlLineMarkersStacked100 = 67
xlLineStacked = 63
xlLineStacked100 = 64
xlLineStacked100Ex = 129
xlLineStackedEx = 128
xlOtherCombinations = 116
xlPareto = 122
xlPie = 5
xlPieEx = 130
xlPieExploded = 69
xlPieOfPie = 68
xlPyramidBarClustered = 109
xlPyramidBarStacked = 110
xlPyramidBarStacked100 = 111
xlPyramidCol = 112
xlPyramidColClustered = 106
xlPyramidColStacked = 107
xlPyramidColStacked100 = 108
xlRadar = -4151
xlRadarFilled = 82
xlRadarMarkers = 81
xlRegionMap = 140
xlStockHLC = 88
xlStockOHLC = 89
xlStockVHLC = 90
xlStockVOHLC = 91
xlSuggestedChart = -2
xlSunburst = 120
xlSurface = 83
xlSurfaceTopView = 85
xlSurfaceTopViewWireframe = 86
xlSurfaceWireframe = 84
xlTreemap = 117
xlWaterfall = 119
xlXYScatter = -4169
xlXYScatterEx = 138
xlXYScatterLines = 74
xlXYScatterLinesNoMarkers = 75
xlXYScatterSmooth = 72
xlXYScatterSmoothNoMarkers = 73

# XlCheckInVersionType enumeration
xlCheckInMajorVersion = 1
xlCheckInMinorVersion = 0
xlCheckInOverwriteVersion = 2

# XlClipboardFormat enumeration
xlClipboardFormatBIFF = 8
xlClipboardFormatBIFF12 = 63
xlClipboardFormatBIFF2 = 18
xlClipboardFormatBIFF3 = 20
xlClipboardFormatBIFF4 = 30
xlClipboardFormatBinary = 15
xlClipboardFormatBitmap = 9
xlClipboardFormatCGM = 13
xlClipboardFormatCSV = 5
xlClipboardFormatDIF = 4
xlClipboardFormatDspText = 12
xlClipboardFormatEmbeddedObject = 21
xlClipboardFormatEmbedSource = 22
xlClipboardFormatLink = 11
xlClipboardFormatLinkSource = 23
xlClipboardFormatLinkSourceDesc = 32
xlClipboardFormatMovie = 24
xlClipboardFormatNative = 14
xlClipboardFormatObjectDesc = 31
xlClipboardFormatObjectLink = 19
xlClipboardFormatOwnerLink = 17
xlClipboardFormatPICT = 2
xlClipboardFormatPrintPICT = 3
xlClipboardFormatRTF = 7
xlClipboardFormatScreenPICT = 29
xlClipboardFormatStandardFont = 28
xlClipboardFormatStandardScale = 27
xlClipboardFormatSYLK = 6
xlClipboardFormatTable = 16
xlClipboardFormatText = 0
xlClipboardFormatToolFace = 25
xlClipboardFormatToolFacePICT = 26
xlClipboardFormatVALU = 1
xlClipboardFormatWK1 = 10

# XlCmdType enumeration
xlCmdCube = 1
xlCmdDAX = 8
xlCmdDefault = 4
xlCmdExcel = 7
xlCmdList = 5
xlCmdSql = 2
xlCmdTable = 3
xlCmdTableCollection = 6

# XlColorIndex enumeration
xlColorIndexAutomatic = -4105
xlColorIndexNone = -4142

# XlColumnDataType enumeration
xlDMYFormat = 4
xlDYMFormat = 7
xlEMDFormat = 10
xlGeneralFormat = 1
xlMDYFormat = 3
xlMYDFormat = 6
xlSkipColumn = 9
xlTextFormat = 2
xlYDMFormat = 8
xlYMDFormat = 5

# XlCommandUnderlines enumeration
xlCommandUnderlinesAutomatic = -4105
xlCommandUnderlinesOff = -4146
xlCommandUnderlinesOn = 1

# XlCommentDisplayMode enumeration
xlCommentAndIndicator = 1
xlCommentIndicatorOnly = -1
xlNoIndicator = 0

# XlConditionValueTypes enumeration
xlConditionValueAutomaticMax = 7
xlConditionValueAutomaticMin = 6
xlConditionValueFormula = 4
xlConditionValueHighestValue = 2
xlConditionValueLowestValue = 1
xlConditionValueNone = -1
xlConditionValueNumber = 0
xlConditionValuePercent = 3
xlConditionValuePercentile = 5

# XlConnectionType enumeration
xlConnectionTypeDATAFEED = 6
xlConnectionTypeMODEL = 7
xlConnectionTypeNOSOURCE = 9
xlConnectionTypeODBC = 2
xlConnectionTypeOLEDB = 1
xlConnectionTypeTEXT = 4
xlConnectionTypeWEB = 5
xlConnectionTypeWORKSHEET = 8
xlConnectionTypeXMLMAP = 3

# XlConsolidationFunction enumeration
xlAverage = -4106
xlCount = -4112
xlCountNums = -4113
xlDistinctCount = 11
xlMax = -4136
xlMin = -4139
xlProduct = -4149
xlStDev = -4155
xlStDevP = -4156
xlSum = -4157
xlUnknown = 1000
xlVar = -4164
xlVarP = -4165

# XlContainsOperator enumeration
xlBeginsWith = 2
xlContains = 0
xlDoesNotContain = 1
xlEndsWith = 3

# XlCopyPictureFormat enumeration
xlBitmap = 2
xlPicture = -4147

# XlCorruptLoad enumeration
xlExtractData = 2
xlNormalLoad = 0
xlRepairFile = 1

# XlCreator enumeration
xlCreatorCode = 1480803660

# XlCredentialsMethod enumeration
CredentialsMethodIntegrated = 0
CredentialsMethodNone = 1
CredentialsMethodStored = 2

# XlCubeFieldSubType enumeration
xlCubeAttribute = 4
xlCubeCalculatedMeasure = 5
xlCubeHierarchy = 1
xlCubeImplicitMeasure = 11
xlCubeKPIGoal = 7
xlCubeKPIStatus = 8
xlCubeKPITrend = 9
xlCubeKPIValue = 6
xlCubeKPIWeight = 10
xlCubeMeasure = 2
xlCubeSet = 3

# XlCubeFieldType enumeration
xlHierarchy = 1
xlMeasure = 2
xlSet = 3

# XlCutCopyMode enumeration
xlCopy = 1
xlCut = 2

# XlCVError enumeration
xlErrDiv0 = 2007
xlErrNA = 2042
xlErrName = 2029
xlErrNull = 2000
xlErrNum = 2036
xlErrRef = 2023
xlErrValue = 2015
xlErrSpill = 2045

# XlDataBarAxisPosition enumeration
xlDataBarAxisAutomatic = 0
xlDataBarAxisMidpoint = 1
xlDataBarAxisNone = 2

# XlDataBarBorderType enumeration
xlDataBarBorderNone = 0
xlDataBarBorderSolid = 1

# XlDataBarFillType enumeration
xlDataBarFillGradient = 1
xlDataBarFillSolid = 0

# XlDataBarNegativeColorType enumeration
xlDataBarColor = 0
xlDataBarSameAsPositive = 1

# XlDataLabelPosition enumeration
xlLabelPositionAbove = 0
xlLabelPositionBelow = 1
xlLabelPositionBestFit = 5
xlLabelPositionCenter = -4108
xlLabelPositionCustom = 7
xlLabelPositionInsideBase = 4
xlLabelPositionInsideEnd = 3
xlLabelPositionLeft = -4131
xlLabelPositionMixed = 6
xlLabelPositionOutsideEnd = 2
xlLabelPositionRight = -4152

# XlDataLabelSeparator enumeration
xlDataLabelSeparatorDefault = 1

# XlDataLabelsType enumeration
xlDataLabelsShowBubbleSizes = 6
xlDataLabelsShowLabel = 4
xlDataLabelsShowLabelAndPercent = 5
xlDataLabelsShowNone = -4142
xlDataLabelsShowPercent = 3
xlDataLabelsShowValue = 2

# XlDataSeriesDate enumeration
xlDay = 1
xlMonth = 3
xlWeekday = 2
xlYear = 4

# XlDataSeriesType enumeration
xlAutoFill = 4
xlChronological = 3
xlDataSeriesLinear = -4132
xlGrowth = 2

# XlDeleteShiftDirection enumeration
xlShiftToLeft = -4159
xlShiftUp = -4162

# XlDirection enumeration
xlDown = -4121
xlToLeft = -4159
xlToRight = -4161
xlUp = -4162

# XlDisplayBlanksAs enumeration
xlInterpolated = 3
xlNotPlotted = 1
xlZero = 2

# XlDisplayDrawingObjects enumeration
xlDisplayShapes = -4104
xlHide = 3
xlPlaceholders = 2

# XlDisplayUnit enumeration
xlHundredMillions = -8
xlHundreds = -2
xlHundredThousands = -5
xlMillionMillions = -10
xlMillions = -6
xlTenMillions = -7
xlTenThousands = -4
xlThousandMillions = -9
xlThousands = -3

# XlDupeUnique enumeration
xlDuplicate = 1
xlUnique = 0

# XlDVAlertStyle enumeration
xlValidAlertInformation = 3
xlValidAlertStop = 1
xlValidAlertWarning = 2

# XlDVType enumeration
xlValidateCustom = 7
xlValidateDate = 4
xlValidateDecimal = 2
xlValidateInputOnly = 0
xlValidateList = 3
xlValidateTextLength = 6
xlValidateTime = 5
xlValidateWholeNumber = 1

# XlDynamicFilterCriteria enumeration
xlFilterAboveAverage = 33
xlFilterAllDatesInPeriodApril = 24
xlFilterAllDatesInPeriodAugust = 28
xlFilterAllDatesInPeriodDecember = 32
xlFilterAllDatesInPeriodFebruary = 22
xlFilterAllDatesInPeriodJanuary = 21
xlFilterAllDatesInPeriodJuly = 27
xlFilterAllDatesInPeriodJune = 26
xlFilterAllDatesInPeriodMarch = 23
xlFilterAllDatesInPeriodMay = 25
xlFilterAllDatesInPeriodNovember = 31
xlFilterAllDatesInPeriodOctober = 30
xlFilterAllDatesInPeriodQuarter1 = 17
xlFilterAllDatesInPeriodQuarter2 = 18
xlFilterAllDatesInPeriodQuarter3 = 19
xlFilterAllDatesInPeriodQuarter4 = 20
xlFilterAllDatesInPeriodSeptember = 29
xlFilterBelowAverage = 34
xlFilterLastMonth = 8
xlFilterLastQuarter = 11
xlFilterLastWeek = 5
xlFilterLastYear = 14
xlFilterNextMonth = 9
xlFilterNextQuarter = 12
xlFilterNextWeek = 6
xlFilterNextYear = 15
xlFilterThisMonth = 7
xlFilterThisQuarter = 10
xlFilterThisWeek = 4
xlFilterThisYear = 13
xlFilterToday = 1
xlFilterTomorrow = 3
xlFilterYearToDate = 16
xlFilterYesterday = 2

# XlEditionFormat enumeration
xlBIFF = 2
xlPICT = 1
xlRTF = 4
xlVALU = 8

# XlEditionOptionsOption enumeration
xlAutomaticUpdate = 4
xlCancel = 1
xlChangeAttributes = 6
xlManualUpdate = 5
xlOpenSource = 3
xlSelect = 3
xlSendPublisher = 2
xlUpdateSubscriber = 2

# XlEditionType enumeration
xlPublisher = 1
xlSubscriber = 2

# XlEnableCancelKey enumeration
xlDisabled = 0
xlErrorHandler = 2
xlInterrupt = 1

# XlEnableSelection enumeration
xlNoRestrictions = 0
xlNoSelection = -4142
xlUnlockedCells = 1

# XlEndStyleCap enumeration
xlCap = 1
xlNoCap = 2

# XlErrorBarDirection enumeration
xlX = -4168
xlY = 1

# XlErrorBarInclude enumeration
xlErrorBarIncludeBoth = 1
xlErrorBarIncludeMinusValues = 3
xlErrorBarIncludeNone = -4142
xlErrorBarIncludePlusValues = 2

# XlErrorBarType enumeration
xlErrorBarTypeCustom = -4114
xlErrorBarTypeFixedValue = 1
xlErrorBarTypePercent = 2
xlErrorBarTypeStDev = -4155
xlErrorBarTypeStError = 4

# XlErrorChecks enumeration
xlEmptyCellReferences = 7
xlEvaluateToError = 1
xlInconsistentFormula = 4
xlInconsistentListFormula = 9
xlListDataValidation = 8
xlNumberAsText = 3
xlOmittedCells = 5
xlStaleValue = 12
xlTextDate = 2
xlUnlockedFormulaCells = 6

# XlFileAccess enumeration
xlReadOnly = 3
xlReadWrite = 2

# XlFileFormat enumeration
xlAddIn = 18
xlAddIn8 = 18
xlCSV = 6
xlCSVMac = 22
xlCSVMSDOS = 24
xlCSVUTF8 = 62
xlCSVWindows = 23
xlCurrentPlatformText = -4158
xlDBF2 = 7
xlDBF3 = 8
xlDBF4 = 11
xlDIF = 9
xlExcel12 = 50
xlExcel2 = 16
xlExcel2FarEast = 27
xlExcel3 = 29
xlExcel4 = 33
xlExcel4Workbook = 35
xlExcel5 = 39
xlExcel7 = 39
xlExcel8 = 56
xlExcel9795 = 43
xlHtml = 44
xlIntlAddIn = 26
xlIntlMacro = 25
xlOpenDocumentSpreadsheet = 60
xlOpenXMLAddIn = 55
xlOpenXMLStrictWorkbook = "61 (&H3D)"
xlOpenXMLTemplate = 54
xlOpenXMLTemplateMacroEnabled = 53
xlOpenXMLWorkbook = 51
xlOpenXMLWorkbookMacroEnabled = 52
xlSYLK = 2
xlTemplate = 17
xlTemplate8 = 17
xlTextMac = 19
xlTextMSDOS = 21
xlTextPrinter = 36
xlTextWindows = 20
xlUnicodeText = 42
xlWebArchive = 45
xlWJ2WD1 = 14
xlWJ3 = 40
xlWJ3FJ3 = 41
xlWK1 = 5
xlWK1ALL = 31
xlWK1FMT = 30
xlWK3 = 15
xlWK3FM3 = 32
xlWK4 = 38
xlWKS = 4
xlWorkbookDefault = 51
xlWorkbookNormal = -4143
xlWorks2FarEast = 28
xlWQ1 = 34
xlXMLSpreadsheet = 46

# XlFileValidationPivotMode enumeration
xlFileValidationPivotDefault = 0
xlFileValidationPivotRun = 1
xlFileValidationPivotSkip = 2

# XlFillWith enumeration
xlFillWithAll = -4104
xlFillWithContents = 2
xlFillWithFormats = -4122

# XlFilterAction enumeration
xlFilterCopy = 2
xlFilterInPlace = 1

# XlFilterAllDatesInPeriod enumeration
xlFilterAllDatesInPeriodDay = 2
xlFilterAllDatesInPeriodHour = 3
xlFilterAllDatesInPeriodMinute = 4
xlFilterAllDatesInPeriodMonth = 1
xlFilterAllDatesInPeriodSecond = 5
xlFilterAllDatesInPeriodYear = 0

# XlFindLookIn enumeration
xlComments = -4144
xlCommentsThreaded = -4184
xlFormulas = -4123
xlValues = -4163

# XlFixedFormatQuality enumeration
xlQualityMinimum = 1
xlQualityStandard = 0

# XlFixedFormatType enumeration
xlTypePDF = 0
xlTypeXPS = 1

# XlFormatConditionOperator enumeration
xlBetween = 1
xlEqual = 3
xlGreater = 5
xlGreaterEqual = 7
xlLess = 6
xlLessEqual = 8
xlNotBetween = 2
xlNotEqual = 4

# XlFormatConditionType enumeration
xlAboveAverageCondition = 12
xlBlanksCondition = 10
xlCellValue = 1
xlColorScale = 3
xlDataBar = 4
xlErrorsCondition = 16
xlExpression = 2
xlIconSet = 6
xlNoBlanksCondition = 13
xlNoErrorsCondition = 17
xlTextString = 9
xlTimePeriod = 11
xlTop10 = 5
xlUniqueValues = 8

# XlFormatFilterTypes enumeration
FilterBottom = 0
FilterBottomPercent = 2
FilterTop = 1
FilterTopPercent = 3

# XlFormControl enumeration
xlButtonControl = 0
xlCheckBox = 1
xlDropDown = 2
xlEditBox = 3
xlGroupBox = 4
xlLabel = 5
xlListBox = 6
xlOptionButton = 7
xlScrollBar = 8
xlSpinner = 9

# XlFormulaLabel enumeration
xlColumnLabels = 2
xlMixedLabels = 3
xlNoLabels = -4142
xlRowLabels = 1

# XlGenerateTableRefs enumeration
xlA1TableRefs = 0
xlTableNames = 1

# XlGradientFillType enumeration
GradientFillLinear = 0
GradientFillPath = 1

# XlHAlign enumeration
xlHAlignCenter = -4108
xlHAlignCenterAcrossSelection = 7
xlHAlignDistributed = -4117
xlHAlignFill = 5
xlHAlignGeneral = 1
xlHAlignJustify = -4130
xlHAlignLeft = -4131
xlHAlignRight = -4152

# XlHebrewModes enumeration
xlHebrewFullScript = 0
xlHebrewMixedAuthorizedScript = 3
xlHebrewMixedScript = 2
xlHebrewPartialScript = 1

# XlHighlightChangesTime enumeration
xlAllChanges = 2
xlNotYetReviewed = 3
xlSinceMyLastSave = 1

# XlHtmlType enumeration
xlHtmlCalc = 1
xlHtmlChart = 3
xlHtmlList = 2
xlHtmlStatic = 0

# XlIcon enumeration
xlIcon0Bars = 37
xlIcon0FilledBoxes = 52
xlIcon1Bar = 38
xlIcon1FilledBox = 51
xlIcon2Bars = 39
xlIcon2FilledBoxes = 50
xlIcon3Bars = 40
xlIcon3FilledBoxes = 49
xlIcon4Bars = 41
xlIcon4FilledBoxes = 48
xlIconBlackCircle = 32
xlIconBlackCircleWithBorder = 13
xlIconCircleWithOneWhiteQuarter = 33
xlIconCircleWithThreeWhiteQuarters = 35
xlIconCircleWithTwoWhiteQuarters = 34
xlIconGoldStar = 42
xlIconGrayCircle = 31
xlIconGrayDownArrow = 6
xlIconGrayDownInclineArrow = 28
xlIconGraySideArrow = 5
xlIconGrayUpArrow = 4
xlIconGrayUpInclineArrow = 27
xlIconGreenCheck = 22
xlIconGreenCheckSymbol = 19
xlIconGreenCircle = 10
xlIconGreenFlag = 7
xlIconGreenTrafficLight = 14
xlIconGreenUpArrow = 1
xlIconGreenUpTriangle = 45
xlIconHalfGoldStar = 43
xlIconNoCellIcon = -1
xlIconPinkCircle = 30
xlIconRedCircle = 29
xlIconRedCircleWithBorder = 12
xlIconRedCross = 24
xlIconRedCrossSymbol = 21
xlIconRedDiamond = 18
xlIconRedDownArrow = 3
xlIconRedDownTriangle = 47
xlIconRedFlag = 9
xlIconRedTrafficLight = 16
xlIconSilverStar = 44
xlIconWhiteCircleAllWhiteQuarters = 36
xlIconYellowCircle = 11
xlIconYellowDash = 46
xlIconYellowDownInclineArrow = 26
xlIconYellowExclamation = 23
xlIconYellowExclamationSymbol = 20
xlIconYellowFlag = 8
xlIconYellowSideArrow = 2
xlIconYellowTrafficLight = 15
xlIconYellowTriangle = 17
xlIconYellowUpInclineArrow = 25

# XlIconSet enumeration
xl3Arrows = 1
xl3ArrowsGray = 2
xl3Flags = 3
xl3Signs = 6
xl3Symbols = 7
xl3TrafficLights1 = 4
xl3TrafficLights2 = 5
xl4Arrows = 8
xl4ArrowsGray = 9
xl4CRV = 11
xl4RedToBlack = 10
xl4TrafficLights = 12
xl5Arrows = 13
xl5ArrowsGray = 14
xl5CRV = 15
xl5Quarters = 16

# XlIMEMode enumeration
xlIMEModeAlpha = 8
xlIMEModeAlphaFull = 7
xlIMEModeDisable = 3
xlIMEModeHangul = 10
xlIMEModeHangulFull = 9
xlIMEModeHiragana = 4
xlIMEModeKatakana = 5
xlIMEModeKatakanaHalf = 6
xlIMEModeNoControl = 0
xlIMEModeOff = 2
xlIMEModeOn = 1

# XlImportDataAs enumeration
xlPivotTableReport = 1
xlQueryTable = 0

# XlInsertFormatOrigin enumeration
xlFormatFromLeftOrAbove = 0
xlFormatFromRightOrBelow = 1

# XlInsertShiftDirection enumeration
xlShiftDown = -4121
xlShiftToRight = -4161

# XlLayoutFormType enumeration
xlOutline = 1
xlTabular = 0

# XlLayoutRowType enumeration
xlCompactRow = 0
xlOutlineRow = 2
xlTabularRow = 1

# XlLegendPosition enumeration
xlLegendPositionBottom = -4107
xlLegendPositionCorner = 2
xlLegendPositionCustom = -4161
xlLegendPositionLeft = -4131
xlLegendPositionRight = -4152
xlLegendPositionTop = -4160

# XlLineStyle enumeration
xlContinuous = 1
xlDash = -4115
xlDashDot = 4
xlDashDotDot = 5
xlDot = -4118
xlDouble = -4119
xlLineStyleNone = -4142
xlSlantDashDot = 13

# XlLink enumeration
xlExcelLinks = 1
xlOLELinks = 2
xlPublishers = 5
xlSubscribers = 6

# XlLinkedDataTypeState enumeration
xlLinkedDataTypeStateNone = 0
xlLinkedDataTypeStateValidLinkedData = 1
xlLinkedDataTypeStateDisambiguationNeeded = 2
xlLinkedDataTypeStateBrokenLinkedData = 3
xlLinkedDataTypeStateFetchingData = 4

# XlLinkInfo enumeration
xlEditionDate = 2
xlLinkInfoStatus = 3
xlUpdateState = 1

# XlLinkInfoType enumeration
xlLinkInfoOLELinks = 2
xlLinkInfoPublishers = 5
xlLinkInfoSubscribers = 6

# XlLinkStatus enumeration
xlLinkStatusCopiedValues = 10
xlLinkStatusIndeterminate = 5
xlLinkStatusInvalidName = 7
xlLinkStatusMissingFile = 1
xlLinkStatusMissingSheet = 2
xlLinkStatusNotStarted = 6
xlLinkStatusOK = 0
xlLinkStatusOld = 3
xlLinkStatusSourceNotCalculated = 4
xlLinkStatusSourceNotOpen = 8
xlLinkStatusSourceOpen = 9

# XlLinkType enumeration
xlLinkTypeExcelLinks = 1
xlLinkTypeOLELinks = 2

# XlListConflict enumeration
xlListConflictDialog = 0
xlListConflictDiscardAllConflicts = 2
xlListConflictError = 3
xlListConflictRetryAllConflicts = 1

# XlListDataType enumeration
xlListDataTypeCheckbox = 9
xlListDataTypeChoice = 6
xlListDataTypeChoiceMulti = 7
xlListDataTypeCounter = 11
xlListDataTypeCurrency = 4
xlListDataTypeDateTime = 5
xlListDataTypeHyperLink = 10
xlListDataTypeListLookup = 8
xlListDataTypeMultiLineRichText = 12
xlListDataTypeMultiLineText = 2
xlListDataTypeNone = 0
xlListDataTypeNumber = 3
xlListDataTypeText = 1

# XlListObjectSourceType enumeration
xlSrcExternal = 0
xlSrcModel = 4
xlSrcQuery = 3
xlSrcRange = 1
xlSrcXml = 2

# XlLocationInTable enumeration
xlColumnHeader = -4110
xlColumnItem = 5
xlDataHeader = 3
xlDataItem = 7
xlPageHeader = 2
xlPageItem = 6
xlRowHeader = -4153
xlRowItem = 4
xlTableBody = 8

# XlLookAt enumeration
xlPart = 2
xlWhole = 1

# XlLookFor enumeration
LookForBlanks = 0
LookForErrors = 1
LookForFormulas = 2

# XlMailSystem enumeration
xlMAPI = 1
xlNoMailSystem = 0
xlPowerTalk = 2

# XlMarkerStyle enumeration
xlMarkerStyleAutomatic = -4105
xlMarkerStyleCircle = 8
xlMarkerStyleDash = -4115
xlMarkerStyleDiamond = 2
xlMarkerStyleDot = -4118
xlMarkerStyleNone = -4142
xlMarkerStylePicture = -4147
xlMarkerStylePlus = 9
xlMarkerStyleSquare = 1
xlMarkerStyleStar = 5
xlMarkerStyleTriangle = 3
xlMarkerStyleX = -4168

# XlMeasurementUnits enumeration
xlCentimeters = 1
xlInches = 0
xlMillimeters = 2

# XlMouseButton enumeration
xlNoButton = 0
xlPrimaryButton = 1
xlSecondaryButton = 2

# XlMousePointer enumeration
xlDefault = -4143
xlIBeam = 3
xlNorthwestArrow = 1
xlWait = 2

# XlMSApplication enumeration
xlMicrosoftAccess = 4
xlMicrosoftFoxPro = 5
xlMicrosoftMail = 3
xlMicrosoftPowerPoint = 2
xlMicrosoftProject = 6
xlMicrosoftSchedulePlus = 7
xlMicrosoftWord = 1

# XlOartHorizontalOverflow enumeration
xlOartHorizontalOverflowClip = 1
xlOartHorizontalOverflowOverflow = 0

# XlOartVerticalOverflow enumeration
xlOartVerticalOverflowClip = 1
xlOartVerticalOverflowEllipsis = 2
xlOartVerticalOverflowOverflow = 0

# XlObjectSize enumeration
xlFitToPage = 2
xlFullPage = 3
xlScreenSize = 1

# XlOLEType enumeration
xlOLEControl = 2
xlOLEEmbed = 1
xlOLELink = 0

# XlOLEVerb enumeration
xlVerbOpen = 2
xlVerbPrimary = 1

# XlOrder enumeration
xlDownThenOver = 1
xlOverThenDown = 2

# XlOrientation enumeration
xlDownward = -4170
xlHorizontal = -4128
xlUpward = -4171
xlVertical = -4166

# XlPageBreak enumeration
xlPageBreakAutomatic = -4105
xlPageBreakManual = -4135
xlPageBreakNone = -4142

# XlPageBreakExtent enumeration
xlPageBreakFull = 1
xlPageBreakPartial = 2

# XlPageOrientation enumeration
xlLandscape = 2
xlPortrait = 1

# XlPaperSize enumeration
xlPaper10x14 = 16
xlPaper11x17 = 17
xlPaperA3 = 8
xlPaperA4 = 9
xlPaperA4Small = 10
xlPaperA5 = 11
xlPaperB4 = 12
xlPaperB5 = 13
xlPaperCsheet = 24
xlPaperDsheet = 25
xlPaperEnvelope10 = 20
xlPaperEnvelope11 = 21
xlPaperEnvelope12 = 22
xlPaperEnvelope14 = 23
xlPaperEnvelope9 = 19
xlPaperEnvelopeB4 = 33
xlPaperEnvelopeB5 = 34
xlPaperEnvelopeB6 = 35
xlPaperEnvelopeC3 = 29
xlPaperEnvelopeC4 = 30
xlPaperEnvelopeC5 = 28
xlPaperEnvelopeC6 = 31
xlPaperEnvelopeC65 = 32
xlPaperEnvelopeDL = 27
xlPaperEnvelopeItaly = 36
xlPaperEnvelopeMonarch = 37
xlPaperEnvelopePersonal = 38
xlPaperEsheet = 26
xlPaperExecutive = 7
xlPaperFanfoldLegalGerman = 41
xlPaperFanfoldStdGerman = 40
xlPaperFanfoldUS = 39
xlPaperFolio = 14
xlPaperLedger = 4
xlPaperLegal = 5
xlPaperLetter = 1
xlPaperLetterSmall = 2
xlPaperNote = 18
xlPaperQuarto = 15
xlPaperStatement = 6
xlPaperTabloid = 3
xlPaperUser = 256

# XlParameterDataType enumeration
xlParamTypeBigInt = -5
xlParamTypeBinary = -2
xlParamTypeBit = -7
xlParamTypeChar = 1
xlParamTypeDate = 9
xlParamTypeDecimal = 3
xlParamTypeDouble = 8
xlParamTypeFloat = 6
xlParamTypeInteger = 4
xlParamTypeLongVarBinary = -4
xlParamTypeLongVarChar = -1
xlParamTypeNumeric = 2
xlParamTypeReal = 7
xlParamTypeSmallInt = 5
xlParamTypeTime = 10
xlParamTypeTimestamp = 11
xlParamTypeTinyInt = -6
xlParamTypeUnknown = 0
xlParamTypeVarBinary = -3
xlParamTypeVarChar = 12
xlParamTypeWChar = -8

# XlParameterType enumeration
xlConstant = 1
xlPrompt = 0
xlRange = 2

# XlPasteSpecialOperation enumeration
xlPasteSpecialOperationAdd = 2
xlPasteSpecialOperationDivide = 5
xlPasteSpecialOperationMultiply = 4
xlPasteSpecialOperationNone = -4142
xlPasteSpecialOperationSubtract = 3

# XlPasteType enumeration
xlPasteAll = -4104
xlPasteAllExceptBorders = 7
xlPasteAllMergingConditionalFormats = 14
xlPasteAllUsingSourceTheme = 13
xlPasteColumnWidths = 8
xlPasteComments = -4144
xlPasteFormats = -4122
xlPasteFormulas = -4123
xlPasteFormulasAndNumberFormats = 11
xlPasteValidation = 6
xlPasteValues = -4163
xlPasteValuesAndNumberFormats = 12

# XlPattern enumeration
xlPatternAutomatic = -4105
xlPatternChecker = 9
xlPatternCrissCross = 16
xlPatternDown = -4121
xlPatternGray16 = 17
xlPatternGray25 = -4124
xlPatternGray50 = -4125
xlPatternGray75 = -4126
xlPatternGray8 = 18
xlPatternGrid = 15
xlPatternHorizontal = -4128
xlPatternLightDown = 13
xlPatternLightHorizontal = 11
xlPatternLightUp = 14
xlPatternLightVertical = 12
xlPatternNone = -4142
xlPatternSemiGray75 = 10
xlPatternSolid = 1
xlPatternUp = -4162
xlPatternVertical = -4166

# XlPhoneticAlignment enumeration
xlPhoneticAlignCenter = 2
xlPhoneticAlignDistributed = 3
xlPhoneticAlignLeft = 1
xlPhoneticAlignNoControl = 0

# XlPhoneticCharacterType enumeration
xlHiragana = 2
xlKatakana = 1
xlKatakanaHalf = 0
xlNoConversion = 3

# XlPictureAppearance enumeration
xlPrinter = 2
xlScreen = 1

# XlPictureConvertorType enumeration
xlBMP = 1
xlCGM = 7
xlDRW = 4
xlDXF = 5
xlEPS = 8
xlHGL = 6
xlPCT = 13
xlPCX = 10
xlPIC = 11
xlPLT = 12
xlTIF = 9
xlWMF = 2
xlWPG = 3

# XlPieSliceIndex enumeration
xlCenterPoint = 5
xlInnerCenterPoint = 8
xlInnerClockwisePoint = 7
xlInnerCounterClockwisePoint = 9
xlMidClockwiseRadiusPoint = 4
xlMidCounterClockwiseRadiusPoint = 6
xlOuterCenterPoint = 2
xlOuterClockwisePoint = 3
xlOuterCounterClockwisePoint = 1

# XlPieSliceLocation enumeration
xlHorizontalCoordinate = 1
xlVerticalCoordinate = 2

# XlPivotCellType enumeration
xlPivotCellBlankCell = 9
xlPivotCellCustomSubtotal = 7
xlPivotCellDataField = 4
xlPivotCellDataPivotField = 8
xlPivotCellGrandTotal = 3
xlPivotCellPageFieldItem = 6
xlPivotCellPivotField = 5
xlPivotCellPivotItem = 1
xlPivotCellSubtotal = 2
xlPivotCellValue = 0

# XlPivotConditionScope enumeration
xlDataFieldScope = 2
xlFieldsScope = 1
xlSelectionScope = 0

# XlPivotFieldCalculation enumeration
xlDifferenceFrom = 2
xlIndex = 9
xlNoAdditionalCalculation = -4143
xlPercentDifferenceFrom = 4
xlPercentOf = 3
xlPercentOfColumn = 7
xlPercentOfParent = 12
xlPercentOfParentColumn = 11
xlPercentOfParentRow = 10
xlPercentOfRow = 6
xlPercentOfTotal = 8
xlPercentRunningTotal = 13
xlRankAscending = 14
xlRankDecending = 15
xlRunningTotal = 5

# XlPivotFieldDataType enumeration
xlDate = 2
xlNumber = -4145
xlText = -4158

# XlPivotFieldOrientation enumeration
xlColumnField = 2
xlDataField = 4
xlHidden = 0
xlPageField = 3
xlRowField = 1

# XlPivotFieldRepeatLabels enumeration
xlDoNotRepeatLabels = 1
xlRepeatLabels = 2

# XlPivotFilterType enumeration
xlBefore = 31
xlBeforeOrEqualTo = 32
xlAfter = 33
xlAfterOrEqualTo = 34
xlAllDatesInPeriodJanuary = 57
xlAllDatesInPeriodFebruary = 58
xlAllDatesInPeriodMarch = 59
xlAllDatesInPeriodApril = 60
xlAllDatesInPeriodMay = 61
xlAllDatesInPeriodJune = 62
xlAllDatesInPeriodJuly = 63
xlAllDatesInPeriodAugust = 64
xlAllDatesInPeriodSeptember = 65
xlAllDatesInPeriodOctober = 66
xlAllDatesInPeriodNovember = 67
xlAllDatesInPeriodDecember = 68
xlAllDatesInPeriodQuarter1 = 53
xlAllDatesInPeriodQuarter2 = 54
xlAllDatesInPeriodQuarter3 = 55
xlAllDatesInPeriodQuarter4 = 56
xlBottomCount = 2
xlBottomPercent = 4
xlBottomSum = 6
xlCaptionBeginsWith = 17
xlCaptionContains = 21
xlCaptionDoesNotBeginWith = 18
xlCaptionDoesNotContain = 22
xlCaptionDoesNotEndWith = 20
xlCaptionDoesNotEqual = 16
xlCaptionEndsWith = 19
xlCaptionEquals = 15
xlCaptionIsBetween = 27
xlCaptionIsGreaterThan = 23
xlCaptionIsGreaterThanOrEqualTo = 24
xlCaptionIsLessThan = 25
xlCaptionIsLessThanOrEqualTo = 26
xlCaptionIsNotBetween = 28
xlDateBetween = 35
xlDateLastMonth = 45
xlDateLastQuarter = 48
xlDateLastWeek = 42
xlDateLastYear = 51
xlDateNextMonth = 43
xlDateNextQuarter = 46
xlDateNextWeek = 40
xlDateNextYear = 49
xlDateThisMonth = 44
xlDateThisQuarter = 47
xlDateThisWeek = 41
xlDateThisYear = 50
xlDateToday = 38
xlDateTomorrow = 37
xlDateYesterday = 39
xlNotSpecificDate = 30
xlSpecificDate = 29
xlTopCount = 1
xlTopPercent = 3
xlTopSum = 5
xlValueDoesNotEqual = 8
xlValueEquals = 7
xlValueIsBetween = 13
xlValueIsGreaterThan = 9
xlValueIsGreaterThanOrEqualTo = 10
xlValueIsLessThan = 11
xlValueIsLessThanOrEqualTo = 12
xlValueIsNotBetween = 14
xlYearToDate = 52

# XlPivotFormatType enumeration
xlPTClassic = 20
xlPTNone = 21
xlReport1 = 0
xlReport10 = 9
xlReport2 = 1
xlReport3 = 2
xlReport4 = 3
xlReport5 = 4
xlReport6 = 5
xlReport7 = 6
xlReport8 = 7
xlReport9 = 8
xlTable1 = 10
xlTable10 = 19
xlTable2 = 11
xlTable3 = 12
xlTable4 = 13
xlTable5 = 14
xlTable6 = 15
xlTable7 = 16
xlTable8 = 17
xlTable9 = 18

# XlPivotLineType enumeration
xlPivotLineBlank = 3
xlPivotLineGrandTotal = 2
xlPivotLineRegular = 0
xlPivotLineSubtotal = 1

# XlPivotTableMissingItems enumeration
xlMissingItemsDefault = -1
xlMissingItemsMax = 32500
xlMissingItemsMax2 = 1048576
xlMissingItemsNone = 0

# XlPivotTableSourceType enumeration
xlConsolidation = 3
xlDatabase = 1
xlExternal = 2
xlPivotTable = -4148
xlScenario = 4

# XlPivotTableVersionList enumeration
xlPivotTableVersion2000 = 0
xlPivotTableVersion10 = 1
xlPivotTableVersion11 = 2
xlPivotTableVersion12 = 3
xlPivotTableVersion14 = 4
xlPivotTableVersion15 = 5
xlPivotTableVersionCurrent = -1

# XlPlacement enumeration
xlFreeFloating = 3
xlMove = 2
xlMoveAndSize = 1

# XlPlatform enumeration
xlMacintosh = 1
xlMSDOS = 3
xlWindows = 2

# XlPortugueseReform enumeration
xlPortugueseBoth = 3
xlPortuguesePostReform = 2
xlPortuguesePreReform = 1

# XlPrintErrors enumeration
xlPrintErrorsBlank = 1
xlPrintErrorsDash = 2
xlPrintErrorsDisplayed = 0
xlPrintErrorsNA = 3

# XlPrintLocation enumeration
xlPrintInPlace = 16
xlPrintNoComments = -4142
xlPrintSheetEnd = 1

# XlPriority enumeration
xlPriorityHigh = -4127
xlPriorityLow = -4134
xlPriorityNormal = -4143

# XlPropertyDisplayedIn enumeration
xlDisplayPropertyInPivotTable = 1
xlDisplayPropertyInPivotTableAndTooltip = 3
xlDisplayPropertyInTooltip = 2

# XlProtectedViewCloseReason enumeration
xlProtectedViewCloseEdit = 1
xlProtectedViewCloseForced = 2
xlProtectedViewCloseNormal = 0

# XlProtectedViewWindowState enumeration
xlProtectedViewWindowMaximized = 2
xlProtectedViewWindowMinimized = 1
xlProtectedViewWindowNormal = 0

# XlPTSelectionMode enumeration
xlBlanks = 4
xlButton = 15
xlDataAndLabel = 0
xlDataOnly = 2
xlFirstRow = 256
xlLabelOnly = 1
xlOrigin = 3

# XlQueryType enumeration
xlADORecordset = 7
xlDAORecordset = 2
xlODBCQuery = 1
xlOLEDBQuery = 5
xlTextImport = 6
xlWebQuery = 4

# XlRangeAutoFormat enumeration
xlRangeAutoFormat3DEffects1 = 13
xlRangeAutoFormat3DEffects2 = 14
xlRangeAutoFormatAccounting1 = 4
xlRangeAutoFormatAccounting2 = 5
xlRangeAutoFormatAccounting3 = 6
xlRangeAutoFormatAccounting4 = 17
xlRangeAutoFormatClassic1 = 1
xlRangeAutoFormatClassic2 = 2
xlRangeAutoFormatClassic3 = 3
xlRangeAutoFormatClassicPivotTable = 31
xlRangeAutoFormatColor1 = 7
xlRangeAutoFormatColor2 = 8
xlRangeAutoFormatColor3 = 9
xlRangeAutoFormatList1 = 10
xlRangeAutoFormatList2 = 11
xlRangeAutoFormatList3 = 12
xlRangeAutoFormatLocalFormat1 = 15
xlRangeAutoFormatLocalFormat2 = 16
xlRangeAutoFormatLocalFormat3 = 19
xlRangeAutoFormatLocalFormat4 = 20
xlRangeAutoFormatNone = -4142
xlRangeAutoFormatPTNone = 42
xlRangeAutoFormatReport1 = 21
xlRangeAutoFormatReport10 = 30
xlRangeAutoFormatReport2 = 22
xlRangeAutoFormatReport3 = 23
xlRangeAutoFormatReport4 = 24
xlRangeAutoFormatReport5 = 25
xlRangeAutoFormatReport6 = 26
xlRangeAutoFormatReport7 = 27
xlRangeAutoFormatReport8 = 28
xlRangeAutoFormatReport9 = 29
xlRangeAutoFormatSimple = -4154
xlRangeAutoFormatTable1 = 32
xlRangeAutoFormatTable10 = 41
xlRangeAutoFormatTable2 = 33
xlRangeAutoFormatTable3 = 34
xlRangeAutoFormatTable4 = 35
xlRangeAutoFormatTable5 = 36
xlRangeAutoFormatTable6 = 37
xlRangeAutoFormatTable7 = 38
xlRangeAutoFormatTable8 = 39
xlRangeAutoFormatTable9 = 40

# XlRangeValueDataType enumeration
xlRangeValueDefault = 10
xlRangeValueMSPersistXML = 12
xlRangeValueXMLSpreadsheet = 11

# XlReferenceStyle enumeration
xlA1 = 1
xlR1C1 = -4150

# XlReferenceType enumeration
xlAbsolute = 1
xlAbsRowRelColumn = 2
xlRelative = 4
xlRelRowAbsColumn = 3

# XlRemoveDocInfoType enumeration
xlRDIAll = 99
xlRDIComments = 1
xlRDIContentType = 16
xlRDIDefinedNameComments = 18
xlRDIDocumentManagementPolicy = 15
xlRDIDocumentProperties = 8
xlRDIDocumentServerProperties = 14
xlRDIDocumentWorkspace = 10
xlRDIEmailHeader = 5
xlRDIExcelDataModel = 23
xlRDIInactiveDataConnections = 19
xlRDIInkAnnotations = 11
xlRDIInlineWebExtensions = 21
xlRDIPrinterPath = 20
xlRDIPublishInfo = 13
xlRDIRemovePersonalInformation = 4
xlRDIRoutingSlip = 6
xlRDIScenarioComments = 12
xlRDISendForReview = 7
xlRDITaskpaneWebExtensions = 22

# XlRgbColor enumeration
rgbAliceBlue = 16775408
rgbAntiqueWhite = 14150650
rgbAqua = 16776960
rgbAquamarine = 13959039
rgbAzure = 16777200
rgbBeige = 14480885
rgbBisque = 12903679
rgbBlack = 0
rgbBlanchedAlmond = 13495295
rgbBlue = 16711680
rgbBlueViolet = 14822282
rgbBrown = 2763429
rgbBurlyWood = 8894686
rgbCadetBlue = 10526303
rgbChartreuse = 65407
rgbCoral = 5275647
rgbCornflowerBlue = 15570276
rgbCornsilk = 14481663
rgbCrimson = 3937500
rgbDarkBlue = 9109504
rgbDarkCyan = 9145088
rgbDarkGoldenrod = 755384
rgbDarkGray = 11119017
rgbDarkGreen = 25600
rgbDarkGrey = 11119017
rgbDarkKhaki = 7059389
rgbDarkMagenta = 9109643
rgbDarkOliveGreen = 3107669
rgbDarkOrange = 36095
rgbDarkOrchid = 13382297
rgbDarkRed = 139
rgbDarkSalmon = 8034025
rgbDarkSeaGreen = 9419919
rgbDarkSlateBlue = 9125192
rgbDarkSlateGray = 5197615
rgbDarkSlateGrey = 5197615
rgbDarkTurquoise = 13749760
rgbDarkViolet = 13828244
rgbDeepPink = 9639167
rgbDeepSkyBlue = 16760576
rgbDimGray = 6908265
rgbDimGrey = 6908265
rgbDodgerBlue = 16748574
rgbFireBrick = 2237106
rgbFloralWhite = 15792895
rgbForestGreen = 2263842
rgbFuchsia = 16711935
rgbGainsboro = 14474460
rgbGhostWhite = 16775416
rgbGold = 55295
rgbGoldenrod = 2139610
rgbGray = 8421504
rgbGreen = 32768
rgbGreenYellow = 3145645
rgbGrey = 8421504
rgbHoneydew = 15794160
rgbHotPink = 11823615
rgbIndianRed = 6053069
rgbIndigo = 8519755
rgbIvory = 15794175
rgbKhaki = 9234160
rgbLavender = 16443110
rgbLavenderBlush = 16118015
rgbLawnGreen = 64636
rgbLemonChiffon = 13499135
rgbLightBlue = 15128749
rgbLightCoral = 8421616
rgbLightCyan = 9145088
rgbLightGoldenrodYellow = 13826810
rgbLightGray = 13882323
rgbLightGreen = 9498256
rgbLightGrey = 13882323
rgbLightPink = 12695295
rgbLightSalmon = 8036607
rgbLightSeaGreen = 11186720
rgbLightSkyBlue = 16436871
rgbLightSlateGray = 10061943
rgbLightSteelBlue = 14599344
rgbLightYellow = 14745599
rgbLime = 65280
rgbLimeGreen = 3329330
rgbLinen = 15134970
rgbMaroon = 128
rgbMediumAquamarine = 11206502
rgbMediumBlue = 13434880
rgbMediumOrchid = 13850042
rgbMediumPurple = 14381203
rgbMediumSeaGreen = 7451452
rgbMediumSlateBlue = 15624315
rgbMediumSpringGreen = 10156544
rgbMediumTurquoise = 13422920
rgbMediumVioletRed = 8721863
rgbMidnightBlue = 7346457
rgbMintCream = 16449525
rgbMistyRose = 14804223
rgbMoccasin = 11920639
rgbNavajoWhite = 11394815
rgbNavy = 8388608
rgbNavyBlue = 8388608
rgbOldLace = 15136253
rgbOlive = 32896
rgbOliveDrab = 2330219
rgbOrange = 42495
rgbOrangeRed = 17919
rgbOrchid = 14053594
rgbPaleGoldenrod = 7071982
rgbPaleGreen = 10025880
rgbPaleTurquoise = 15658671
rgbPaleVioletRed = 9662683
rgbPapayaWhip = 14020607
rgbPeachPuff = 12180223
rgbPeru = 4163021
rgbPink = 13353215
rgbPlum = 14524637
rgbPowderBlue = 15130800
rgbPurple = 8388736
rgbRed = 255
rgbRosyBrown = 9408444
rgbRoyalBlue = 14772545
rgbSalmon = 7504122
rgbSandyBrown = 6333684
rgbSeaGreen = 5737262
rgbSeashell = 15660543
rgbSienna = 2970272
rgbSilver = 12632256
rgbSkyBlue = 15453831
rgbSlateBlue = 13458026
rgbSlateGray = 9470064
rgbSnow = 16448255
rgbSpringGreen = 8388352
rgbSteelBlue = 11829830
rgbTan = 9221330
rgbTeal = 8421376
rgbThistle = 14204888
rgbTomato = 4678655
rgbTurquoise = 13688896
rgbViolet = 15631086
rgbWheat = 11788021
rgbWhite = 16777215
rgbWhiteSmoke = 16119285
rgbYellow = 65535
rgbYellowGreen = 3329434

# XlRobustConnect enumeration
xlAlways = 1
xlAsRequired = 0
xlNever = 2

# XlRowCol enumeration
xlColumns = 2
xlRows = 1

# XlRunAutoMacro enumeration
xlAutoActivate = 3
xlAutoClose = 2
xlAutoDeactivate = 4
xlAutoOpen = 1

# XlSaveAction enumeration
xlDoNotSaveChanges = 2
xlSaveChanges = 1

# XlSaveAsAccessMode enumeration
xlExclusive = 3
xlNoChange = 1
xlShared = 2

# XlSaveConflictResolution enumeration
xlLocalSessionChanges = 2
xlOtherSessionChanges = 3
xlUserResolution = 1

# XlScaleType enumeration
xlScaleLinear = -4132
xlScaleLogarithmic = -4133

# XlSearchDirection enumeration
xlNext = 1
xlPrevious = 2

# XlSearchOrder enumeration
xlByColumns = 2
xlByRows = 1

# XlSearchWithin enumeration
xlWithinSheet = 1
xlWithinWorkbook = 2

# XlSheetType enumeration
xlChart = -4109
xlDialogSheet = -4116
xlExcel4IntlMacroSheet = 4
xlExcel4MacroSheet = 3
xlWorksheet = -4167

# XlSheetVisibility enumeration
xlSheetHidden = 0
xlSheetVeryHidden = 2
xlSheetVisible = -1

# XlSizeRepresents enumeration
xlSizeIsArea = 1
xlSizeIsWidth = 2

# XlSlicerCrossFilterType enumeration
xlSlicerCrossFilterHideButtonsWithNoData = 4
xlSlicerCrossFilterShowItemsWithDataAtTop = 2
xlSlicerCrossFilterShowItemsWithNoData = 3
xlSlicerNoCrossFilter = 1

# XlSlicerSort enumeration
xlSlicerSortAscending = 2
xlSlicerSortDataSourceOrder = 1
xlSlicerSortDescending = 3

# XlSortDataOption enumeration
xlSortNormal = 0
xlSortTextAsNumbers = 1

# XlSortMethod enumeration
xlPinYin = 1
xlStroke = 2

# XlSortMethodOld enumeration
xlCodePage = 2
xlSyllabary = 1

# XlSortOn enumeration
xlSortOnCellColor = 1
xlSortOnFontColor = 2
xlSortOnIcon = 3
xlSortOnValues = 0

# XlSortOrder enumeration
xlAscending = 1
xlDescending = 2
xlManual = -4135

# XlSortOrientation enumeration
xlSortColumns = 1
xlSortRows = 2

# XlSortType enumeration
xlSortLabels = 2
xlSortValues = 1

# XlSourceType enumeration
xlSourceAutoFilter = 3
xlSourceChart = 5
xlSourcePivotTable = 6
xlSourcePrintArea = 2
xlSourceQuery = 7
xlSourceRange = 4
xlSourceSheet = 1
xlSourceWorkbook = 0

# XlSpanishModes enumeration
xlSpanishTuteoAndVoseo = 1
xlSpanishTuteoOnly = 0
xlSpanishVoseoOnly = 2

# XlSparklineRowCol enumeration
xlSparklineColumnsSquare = 2
xlSparklineNonSquare = 0
xlSparklineRowsSquare = 1

# XlSparkScale enumeration
xlSparkScaleCustom = 3
xlSparkScaleGroup = 1
xlSparkScaleSingle = 2

# XlSparkType enumeration
xlSparkColumn = 2
xlSparkColumnStacked100 = 3
xlSparkLine = 1

# XlSpeakDirection enumeration
xlSpeakByColumns = 1
xlSpeakByRows = 0

# XlSpecialCellsValue enumeration
xlErrors = 16
xlLogical = 4
xlNumbers = 1
xlTextValues = 2

# XlStdColorScale enumeration
ColorScaleBlackWhite = 3
ColorScaleGYR = 2
ColorScaleRYG = 1
ColorScaleWhiteBlack = 4

# XlSubscribeToFormat enumeration
xlSubscribeToPicture = -4147
xlSubscribeToText = -4158

# XlSubtotalLocationType enumeration
xlAtBottom = 2
xlAtTop = 1

# XlSummaryColumn enumeration
xlSummaryOnLeft = -4131
xlSummaryOnRight = -4152

# XlSummaryReportType enumeration
xlStandardSummary = 1
xlSummaryPivotTable = -4148

# XlSummaryRow enumeration
xlSummaryAbove = 0
xlSummaryBelow = 1

# XlTableStyleElementType enumeration
xlBlankRow = 19
xlColumnStripe1 = 7
xlColumnStripe2 = 8
xlColumnSubheading1 = 20
xlColumnSubheading2 = 21
xlColumnSubheading3 = 22
xlFirstColumn = 3
xlFirstHeaderCell = 9
xlFirstTotalCell = 11
xlGrandTotalColumn = 4
xlGrandTotalRow = 2
xlHeaderRow = 1
xlLastColumn = 4
xlLastHeaderCell = 10
xlLastTotalCell = 12
xlPageFieldLabels = 26
xlPageFieldValues = 27
xlRowStripe1 = 5
xlRowStripe2 = 6
xlRowSubheading1 = 23
xlRowSubheading2 = 24
xlRowSubheading3 = 25
xlSlicerHoveredSelectedItemWithData = 33
xlSlicerHoveredSelectedItemWithNoData = 35
xlSlicerHoveredUnselectedItemWithData = 32
xlSlicerHoveredUnselectedItemWithNoData = 34
xlSlicerSelectedItemWithData = 30
xlSlicerSelectedItemWithNoData = 31
xlSlicerUnselectedItemWithData = 28
xlSlicerUnselectedItemWithNoData = 29
xlSubtotalColumn1 = 13
xlSubtotalColumn2 = 14
xlSubtotalColumn3 = 15
xlSubtotalRow1 = 16
xlSubtotalRow2 = 17
xlSubtotalRow3 = 18
xlTimelinePeriodLabels1 = 38
xlTimelinePeriodLabels2 = 39
xlTimelineSelectedTimeBlock = 40
xlTimelineSelectedTimeBlockSpace = 42
xlTimelineSelectionLabel = 36
xlTimelineTimeLevel = 37
xlTimelineUnselectedTimeBlock = 41
xlTotalRow = 2
xlWholeTable = 0

# XlTabPosition enumeration
xlTabPositionFirst = 0
xlTabPositionLast = 1

# XlTextParsingType enumeration
xlDelimited = 1
xlFixedWidth = 2

# XlTextQualifier enumeration
xlTextQualifierDoubleQuote = 1
xlTextQualifierNone = -4142
xlTextQualifierSingleQuote = 2

# XlTextVisualLayoutType enumeration
xlTextVisualLTR = 1
xlTextVisualRTL = 2

# XlThemeColor enumeration
xlThemeColorAccent1 = 5
xlThemeColorAccent2 = 6
xlThemeColorAccent3 = 7
xlThemeColorAccent4 = 8
xlThemeColorAccent5 = 9
xlThemeColorAccent6 = 10
xlThemeColorDark1 = 1
xlThemeColorDark2 = 3
xlThemeColorFollowedHyperlink = 12
xlThemeColorHyperlink = 11
xlThemeColorLight1 = 2
xlThemeColorLight2 = 4

# XlThemeFont enumeration
xlThemeFontMajor = 2
xlThemeFontMinor = 1
xlThemeFontNone = 0

# XlThreadMode enumeration
xlThreadModeAutomatic = 0
xlThreadModeManual = 1

# XlTickLabelOrientation enumeration
xlTickLabelOrientationAutomatic = -4105
xlTickLabelOrientationDownward = -4170
xlTickLabelOrientationHorizontal = -4128
xlTickLabelOrientationUpward = -4171
xlTickLabelOrientationVertical = -4166

# XlTickLabelPosition enumeration
xlTickLabelPositionHigh = -4127
xlTickLabelPositionLow = -4134
xlTickLabelPositionNextToAxis = 4
xlTickLabelPositionNone = -4142

# XlTickMark enumeration
xlTickMarkCross = 4
xlTickMarkInside = 2
xlTickMarkNone = -4142
xlTickMarkOutside = 3

# XlTimePeriods enumeration
xlLast7Days = 2
xlLastMonth = 5
xlLastWeek = 4
xlNextMonth = 8
xlNextWeek = 7
xlThisMonth = 9
xlThisWeek = 3
xlToday = 0
xlTomorrow = 6
xlYesterday = 1

# XlTimeUnit enumeration
xlDays = 0
xlMonths = 1
xlYears = 2

# XlToolbarProtection enumeration
xlNoButtonChanges = 1
xlNoChanges = 4
xlNoDockingChanges = 3
xlNoShapeChanges = 2
xlToolbarProtectionNone = -4143

# XlTopBottom enumeration
xlTop10Bottom = 0
xlTop10Top = 1

# XlTotalsCalculation enumeration
xlTotalsCalculationAverage = 2
xlTotalsCalculationCount = 3
xlTotalsCalculationCountNums = 4
xlTotalsCalculationCustom = 9
xlTotalsCalculationMax = 6
xlTotalsCalculationMin = 5
xlTotalsCalculationNone = 0
xlTotalsCalculationStdDev = 7
xlTotalsCalculationSum = 1
xlTotalsCalculationVar = 8

# XlTrendlineType enumeration
xlExponential = 5
xlLinear = -4132
xlLogarithmic = -4133
xlMovingAvg = 6
xlPolynomial = 3
xlPower = 4

# XlUnderlineStyle enumeration
xlUnderlineStyleDouble = -4119
xlUnderlineStyleDoubleAccounting = 5
xlUnderlineStyleNone = -4142
xlUnderlineStyleSingle = 2
xlUnderlineStyleSingleAccounting = 4

# XlUpdateLinks enumeration
xlUpdateLinksAlways = 3
xlUpdateLinksNever = 2
xlUpdateLinksUserSetting = 1

# XlVAlign enumeration
xlVAlignBottom = -4107
xlVAlignCenter = -4108
xlVAlignDistributed = -4117
xlVAlignJustify = -4130
xlVAlignTop = -4160

# XlWBATemplate enumeration
xlWBATChart = -4109
xlWBATExcel4IntlMacroSheet = 4
xlWBATExcel4MacroSheet = 3
xlWBATWorksheet = -4167

# XlWebFormatting enumeration
xlWebFormattingAll = 1
xlWebFormattingNone = 3
xlWebFormattingRTF = 2

# XlWebSelectionType enumeration
xlAllTables = 2
xlEntirePage = 1
xlSpecifiedTables = 3

# XlWindowState enumeration
xlMaximized = -4137
xlMinimized = -4140
xlNormal = -4143

# XlWindowType enumeration
xlChartAsWindow = 5
xlChartInPlace = 4
xlClipboard = 3
xlInfo = -4129
xlWorkbook = 1

# XlWindowView enumeration
xlNormalView = 1
xlPageBreakPreview = 2
xlPageLayoutView = 3

# XlXLMMacroType enumeration
xlCommand = 2
xlFunction = 1
xlNotXLM = 3

# XlXmlExportResult enumeration
xlXmlExportSuccess = 0
xlXmlExportValidationFailed = 1

# XlXmlImportResult enumeration
xlXmlImportElementsTruncated = 1
xlXmlImportSuccess = 0
xlXmlImportValidationFailed = 2

# XlXmlLoadOption enumeration
xlXmlLoadImportToList = 2
xlXmlLoadMapXml = 3
xlXmlLoadOpenXml = 1
xlXmlLoadPromptUser = 0

# XlYesNoGuess enumeration
xlGuess = 0
xlNo = 2
xlYes = 1

class XmlDataBinding:

    def __init__(self, xmldatabinding=None):
        self.xmldatabinding = xmldatabinding

    @property
    def Application(self):
        return self.xmldatabinding.Application

    @property
    def Creator(self):
        return self.xmldatabinding.Creator

    @property
    def Parent(self):
        return self.xmldatabinding.Parent

    @property
    def SourceUrl(self):
        return self.xmldatabinding.SourceUrl

    def ClearSettings(self):
        self.xmldatabinding.ClearSettings()

    def LoadSettings(self, Url=None):
        params = [
            Url if Url is not None else pythoncom.Missing,
        ]
        self.xmldatabinding.LoadSettings(*params)

    def Refresh(self):
        return XlXmlImportResult(self.xmldatabinding.Refresh())


class XmlMap:

    def __init__(self, xmlmap=None):
        self.xmlmap = xmlmap

    @property
    def AdjustColumnWidth(self):
        return self.xmlmap.AdjustColumnWidth

    @property
    def AppendOnImport(self):
        return self.xmlmap.AppendOnImport

    @property
    def Application(self):
        return self.xmlmap.Application

    @property
    def Creator(self):
        return self.xmlmap.Creator

    @property
    def DataBinding(self):
        return XmlDataBinding(self.xmlmap.DataBinding)

    @property
    def IsExportable(self):
        return XPath(self.xmlmap.IsExportable)

    @property
    def Name(self):
        return self.xmlmap.Name

    @Name.setter
    def Name(self, value):
        self.xmlmap.Name = value

    @property
    def Parent(self):
        return self.xmlmap.Parent

    @property
    def PreserveColumnFilter(self):
        return self.xmlmap.PreserveColumnFilter

    @PreserveColumnFilter.setter
    def PreserveColumnFilter(self, value):
        self.xmlmap.PreserveColumnFilter = value

    @property
    def PreserveNumberFormatting(self):
        return self.xmlmap.PreserveNumberFormatting

    @PreserveNumberFormatting.setter
    def PreserveNumberFormatting(self, value):
        self.xmlmap.PreserveNumberFormatting = value

    @property
    def RootElementName(self):
        return self.xmlmap.RootElementName

    @property
    def RootElementNamespace(self):
        return XmlNamespace(self.xmlmap.RootElementNamespace)

    @property
    def SaveDataSourceDefinition(self):
        return self.xmlmap.SaveDataSourceDefinition

    @SaveDataSourceDefinition.setter
    def SaveDataSourceDefinition(self, value):
        self.xmlmap.SaveDataSourceDefinition = value

    @property
    def Schemas(self):
        return XmlSchemas(self.xmlmap.Schemas)

    @property
    def ShowImportExportValidationErrors(self):
        return self.xmlmap.ShowImportExportValidationErrors

    @ShowImportExportValidationErrors.setter
    def ShowImportExportValidationErrors(self, value):
        self.xmlmap.ShowImportExportValidationErrors = value

    @property
    def WorkbookConnection(self):
        return XMLMap(self.xmlmap.WorkbookConnection)

    def Delete(self):
        self.xmlmap.Delete()

    def Export(self, Url=None, Overwrite=None):
        params = [
            Url if Url is not None else pythoncom.Missing,
            Overwrite if Overwrite is not None else pythoncom.Missing,
        ]
        return XlXmlExportResult(self.xmlmap.Export(*params))

    def ExportXml(self, Data=None):
        params = [
            Data if Data is not None else pythoncom.Missing,
        ]
        return XlXmlExportResult(self.xmlmap.ExportXml(*params))

    def Import(self, Url=None, Overwrite=None):
        params = [
            Url if Url is not None else pythoncom.Missing,
            Overwrite if Overwrite is not None else pythoncom.Missing,
        ]
        return XlXmlImportResult(self.xmlmap.Import(*params))

    def ImportXml(self, XmlData=None, Overwrite=None):
        params = [
            XmlData if XmlData is not None else pythoncom.Missing,
            Overwrite if Overwrite is not None else pythoncom.Missing,
        ]
        return XlXmlImportResult(self.xmlmap.ImportXml(*params))


class XmlMaps:

    def __init__(self, xmlmaps=None):
        self.xmlmaps = xmlmaps

    @property
    def Application(self):
        return self.xmlmaps.Application

    @property
    def Count(self):
        return self.xmlmaps.Count

    @property
    def Creator(self):
        return self.xmlmaps.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.xmlmaps.Item):
            return self.xmlmaps.Item(*params)
        else:
            return self.xmlmaps.GetItem(*params)

    @property
    def Parent(self):
        return self.xmlmaps.Parent

    def Add(self, Schema=None, RootElementName=None):
        params = [
            Schema if Schema is not None else pythoncom.Missing,
            RootElementName if RootElementName is not None else pythoncom.Missing,
        ]
        return XmlMap(self.xmlmaps.Add(*params))


class XmlNamespace:

    def __init__(self, xmlnamespace=None):
        self.xmlnamespace = xmlnamespace

    @property
    def Application(self):
        return self.xmlnamespace.Application

    @property
    def Creator(self):
        return self.xmlnamespace.Creator

    @property
    def Parent(self):
        return self.xmlnamespace.Parent

    @property
    def Prefix(self):
        return self.xmlnamespace.Prefix

    @property
    def Uri(self):
        return self.xmlnamespace.Uri


class XmlNamespaces:

    def __init__(self, xmlnamespaces=None):
        self.xmlnamespaces = xmlnamespaces

    @property
    def Application(self):
        return self.xmlnamespaces.Application

    @property
    def Count(self):
        return self.xmlnamespaces.Count

    @property
    def Creator(self):
        return self.xmlnamespaces.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.xmlnamespaces.Item):
            return self.xmlnamespaces.Item(*params)
        else:
            return self.xmlnamespaces.GetItem(*params)

    @property
    def Parent(self):
        return self.xmlnamespaces.Parent

    @property
    def Value(self):
        return self.xmlnamespaces.Value

    def InstallManifest(self, Path=None, InstallForAllUsers=None):
        params = [
            Path if Path is not None else pythoncom.Missing,
            InstallForAllUsers if InstallForAllUsers is not None else pythoncom.Missing,
        ]
        self.xmlnamespaces.InstallManifest(*params)


class XmlSchema:

    def __init__(self, xmlschema=None):
        self.xmlschema = xmlschema

    @property
    def Application(self):
        return self.xmlschema.Application

    @property
    def Creator(self):
        return self.xmlschema.Creator

    @property
    def Name(self):
        return XmlMap(self.xmlschema.Name)

    @property
    def Namespace(self):
        return XmlNamespace(self.xmlschema.Namespace)

    @property
    def Parent(self):
        return self.xmlschema.Parent

    @property
    def XML(self):
        return self.xmlschema.XML


class XmlSchemas:

    def __init__(self, xmlschemas=None):
        self.xmlschemas = xmlschemas

    @property
    def Application(self):
        return self.xmlschemas.Application

    @property
    def Count(self):
        return self.xmlschemas.Count

    @property
    def Creator(self):
        return self.xmlschemas.Creator

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        if callable(self.xmlschemas.Item):
            return self.xmlschemas.Item(*params)
        else:
            return self.xmlschemas.GetItem(*params)

    @property
    def Parent(self):
        return self.xmlschemas.Parent


class XPath:

    def __init__(self, xpath=None):
        self.xpath = xpath

    @property
    def Application(self):
        return self.xpath.Application

    @property
    def Creator(self):
        return self.xpath.Creator

    @property
    def Map(self):
        return XmlMap(self.xpath.Map)

    @property
    def Parent(self):
        return self.xpath.Parent

    @property
    def Repeating(self):
        return self.xpath.Repeating

    @property
    def Value(self):
        return self.xpath.Value

    def Clear(self):
        self.xpath.Clear()

    def SetValue(self, Map=None, XPath=None, SelectionNamespace=None, Repeating=None):
        params = [
            Map if Map is not None else pythoncom.Missing,
            XPath if XPath is not None else pythoncom.Missing,
            SelectionNamespace if SelectionNamespace is not None else pythoncom.Missing,
            Repeating if Repeating is not None else pythoncom.Missing,
        ]
        self.xpath.SetValue(*params)

