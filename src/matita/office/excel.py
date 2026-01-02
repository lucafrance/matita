import win32com.client










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

    def ModifyAppliesToRange(self, *args, Range=None):
        arguments = {"Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.aboveaverage.ModifyAppliesToRange(*args, **arguments)

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.actions.Item):
            return Actions(self.actions.Item(*args, **arguments))
        else:
            return Actions(self.actions.GetItem(*args, **arguments))

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.addins.Item):
            return self.addins.Item(*args, **arguments)
        else:
            return self.addins.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.addins.Parent

    def Add(self, *args, FileName=None, CopyFile=None):
        arguments = {"FileName": FileName, "CopyFile": CopyFile}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return AddIn(self.addins.Add(*args, **arguments))








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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.addins2.Item):
            return self.addins2.Item(*args, **arguments)
        else:
            return self.addins2.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.addins2.Parent

    def Add(self, *args, FileName=None, CopyFile=None):
        arguments = {"FileName": FileName, "CopyFile": CopyFile}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return AddIns(self.addins2.Add(*args, **arguments))








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

    def ChangePassword(self, *args, Password=None):
        arguments = {"Password": Password}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.alloweditrange.ChangePassword(*args, **arguments)

    def Delete(self):
        self.alloweditrange.Delete()

    def Unprotect(self, *args, Password=None):
        arguments = {"Password": Password}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.alloweditrange.Unprotect(*args, **arguments)









class AllowEditRanges:

    def __init__(self, alloweditranges=None):
        self.alloweditranges = alloweditranges

    def __call__(self, item):
        return AllowEditRange(self.alloweditranges(item))

    @property
    def Count(self):
        return self.alloweditranges.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.alloweditranges.Item):
            return self.alloweditranges.Item(*args, **arguments)
        else:
            return self.alloweditranges.GetItem(*args, **arguments)

    def Add(self, *args, Title=None, Range=None, Password=None):
        arguments = {"Title": Title, "Range": Range, "Password": Password}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return AllowEditRange(self.alloweditranges.Add(*args, **arguments))


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

    def Caller(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.application.Caller):
            return self.application.Caller(*args, **arguments)
        else:
            return self.application.GetCaller(*args, **arguments)

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

    def Cells(self, *args, RowIndex=None, ColumnIndex=None):
        arguments = {"RowIndex": RowIndex, "ColumnIndex": ColumnIndex}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.application.Cells):
            return Range(self.application.Cells(*args, **arguments))
        else:
            return Range(self.application.GetCells(*args, **arguments))

    @property
    def Charts(self):
        return Sheets(self.application.Charts)

    def ClipboardFormats(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.application.ClipboardFormats):
            return self.application.ClipboardFormats(*args, **arguments)
        else:
            return self.application.GetClipboardFormats(*args, **arguments)

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

    def FileConverters(self, *args, Index1=None, Index2=None):
        arguments = {"Index1": Index1, "Index2": Index2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.application.FileConverters):
            return self.application.FileConverters(*args, **arguments)
        else:
            return self.application.GetFileConverters(*args, **arguments)

    def FileDialog(self, *args, fileDialogType=None):
        arguments = {"fileDialogType": fileDialogType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.application.FileDialog):
            return self.application.FileDialog(*args, **arguments)
        else:
            return self.application.GetFileDialog(*args, **arguments)

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

    def PreviousSelections(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.application.PreviousSelections):
            return Range(self.application.PreviousSelections(*args, **arguments))
        else:
            return Range(self.application.GetPreviousSelections(*args, **arguments))

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

    def Range(self, *args, Cell1=None, Cell2=None):
        arguments = {"Cell1": Cell1, "Cell2": Cell2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.application.Range):
            return Range(self.application.Range(*args, **arguments))
        else:
            return Range(self.application.GetRange(*args, **arguments))

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

    def RegisteredFunctions(self, *args, Index1=None, Index2=None):
        arguments = {"Index1": Index1, "Index2": Index2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.application.RegisteredFunctions):
            return self.application.RegisteredFunctions(*args, **arguments)
        else:
            return self.application.GetRegisteredFunctions(*args, **arguments)

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

    def ActivateMicrosoftApp(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.ActivateMicrosoftApp(*args, **arguments)

    def AddCustomList(self, *args, ListArray=None, ByRow=None):
        arguments = {"ListArray": ListArray, "ByRow": ByRow}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.AddCustomList(*args, **arguments)

    def Calculate(self):
        self.application.Calculate()

    def CalculateFull(self):
        self.application.CalculateFull()

    def CalculateFullRebuild(self):
        self.application.CalculateFullRebuild()

    def CalculateUntilAsyncQueriesDone(self):
        self.application.CalculateUntilAsyncQueriesDone()

    def CentimetersToPoints(self, *args, Centimeters=None):
        arguments = {"Centimeters": Centimeters}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.CentimetersToPoints(*args, **arguments)

    def CheckAbort(self, *args, KeepAbort=None):
        arguments = {"KeepAbort": KeepAbort}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.CheckAbort(*args, **arguments)

    def CheckSpelling(self, *args, Word=None, CustomDictionary=None, IgnoreUppercase=None):
        arguments = {"Word": Word, "CustomDictionary": CustomDictionary, "IgnoreUppercase": IgnoreUppercase}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.CheckSpelling(*args, **arguments)

    def ConvertFormula(self, *args, Formula=None, FromReferenceStyle=None, ToReferenceStyle=None, ToAbsolute=None, RelativeTo=None):
        arguments = {"Formula": Formula, "FromReferenceStyle": FromReferenceStyle, "ToReferenceStyle": ToReferenceStyle, "ToAbsolute": ToAbsolute, "RelativeTo": RelativeTo}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.ConvertFormula(*args, **arguments)

    def DDEExecute(self, *args, Channel=None, String=None):
        arguments = {"Channel": Channel, "String": String}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.DDEExecute(*args, **arguments)

    def DDEInitiate(self, *args, App=None, Topic=None):
        arguments = {"App": App, "Topic": Topic}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.DDEInitiate(*args, **arguments)

    def DDEPoke(self, *args, Channel=None, Item=None, Data=None):
        arguments = {"Channel": Channel, "Item": Item, "Data": Data}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.DDEPoke(*args, **arguments)

    def DDERequest(self, *args, Channel=None, Item=None):
        arguments = {"Channel": Channel, "Item": Item}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.DDERequest(*args, **arguments)

    def DDETerminate(self, *args, Channel=None):
        arguments = {"Channel": Channel}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.DDETerminate(*args, **arguments)

    def DeleteCustomList(self, *args, ListNum=None):
        arguments = {"ListNum": ListNum}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.DeleteCustomList(*args, **arguments)

    def DisplayXMLSourcePane(self, *args, XmlMap=None):
        arguments = {"XmlMap": XmlMap}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.DisplayXMLSourcePane(*args, **arguments)

    def DoubleClick(self):
        self.application.DoubleClick()

    def Evaluate(self, *args, Name=None):
        arguments = {"Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.Evaluate(*args, **arguments)

    def ExecuteExcel4Macro(self, *args, String=None):
        arguments = {"String": String}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.ExecuteExcel4Macro(*args, **arguments)

    def FindFile(self):
        return self.application.FindFile()

    def GetCustomListContents(self, *args, ListNum=None):
        arguments = {"ListNum": ListNum}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.GetCustomListContents(*args, **arguments)

    def GetCustomListNum(self, *args, ListArray=None):
        arguments = {"ListArray": ListArray}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.GetCustomListNum(*args, **arguments)

    def GetOpenFilename(self, *args, FileFilter=None, FilterIndex=None, Title=None, ButtonText=None, MultiSelect=None):
        arguments = {"FileFilter": FileFilter, "FilterIndex": FilterIndex, "Title": Title, "ButtonText": ButtonText, "MultiSelect": MultiSelect}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.GetOpenFilename(*args, **arguments)

    def GetPhonetic(self, *args, Text=None):
        arguments = {"Text": Text}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.GetPhonetic(*args, **arguments)

    def GetSaveAsFilename(self, *args, InitialFilename=None, FileFilter=None, FilterIndex=None, Title=None, ButtonText=None):
        arguments = {"InitialFilename": InitialFilename, "FileFilter": FileFilter, "FilterIndex": FilterIndex, "Title": Title, "ButtonText": ButtonText}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.GetSaveAsFilename(*args, **arguments)

    def Goto(self, *args, Reference=None, Scroll=None):
        arguments = {"Reference": Reference, "Scroll": Scroll}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.Goto(*args, **arguments)

    def Help(self, *args, HelpFile=None, HelpContextID=None):
        arguments = {"HelpFile": HelpFile, "HelpContextID": HelpContextID}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.Help(*args, **arguments)

    def InchesToPoints(self, *args, Inches=None):
        arguments = {"Inches": Inches}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.InchesToPoints(*args, **arguments)

    def InputBox(self, *args, Prompt=None, Title=None, Default=None, Left=None, Top=None, HelpFile=None, HelpContextID=None, Type=None):
        arguments = {"Prompt": Prompt, "Title": Title, "Default": Default, "Left": Left, "Top": Top, "HelpFile": HelpFile, "HelpContextID": HelpContextID, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.InputBox(*args, **arguments)

    def Intersect(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.Intersect(*args, **arguments)

    def MacroOptions(self, *args, Macro=None, Description=None, HasMenu=None, MenuText=None, HasShortcutKey=None, ShortcutKey=None, Category=None, StatusBar=None, HelpContextID=None, HelpFile=None, ArgumentDescriptions=None):
        arguments = {"Macro": Macro, "Description": Description, "HasMenu": HasMenu, "MenuText": MenuText, "HasShortcutKey": HasShortcutKey, "ShortcutKey": ShortcutKey, "Category": Category, "StatusBar": StatusBar, "HelpContextID": HelpContextID, "HelpFile": HelpFile, "ArgumentDescriptions": ArgumentDescriptions}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.MacroOptions(*args, **arguments)

    def MailLogoff(self):
        self.application.MailLogoff()

    def MailLogon(self, *args, Name=None, Password=None, DownloadNewMail=None):
        arguments = {"Name": Name, "Password": Password, "DownloadNewMail": DownloadNewMail}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.MailLogon(*args, **arguments)

    def NextLetter(self):
        return self.application.NextLetter()

    def OnKey(self, *args, Key=None, Procedure=None):
        arguments = {"Key": Key, "Procedure": Procedure}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.OnKey(*args, **arguments)

    def OnRepeat(self, *args, Text=None, Procedure=None):
        arguments = {"Text": Text, "Procedure": Procedure}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.OnRepeat(*args, **arguments)

    def OnTime(self, *args, EarliestTime=None, Procedure=None, LatestTime=None, Schedule=None):
        arguments = {"EarliestTime": EarliestTime, "Procedure": Procedure, "LatestTime": LatestTime, "Schedule": Schedule}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.OnTime(*args, **arguments)

    def OnUndo(self, *args, Text=None, Procedure=None):
        arguments = {"Text": Text, "Procedure": Procedure}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.OnUndo(*args, **arguments)

    def Quit(self):
        self.application.Quit()

    def RecordMacro(self, *args, BasicCode=None, XlmCode=None):
        arguments = {"BasicCode": BasicCode, "XlmCode": XlmCode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.RecordMacro(*args, **arguments)

    def RegisterXLL(self, *args, FileName=None):
        arguments = {"FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.RegisterXLL(*args, **arguments)

    def Repeat(self):
        self.application.Repeat()

    def Run(self, *args, Macro=None, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Macro": Macro, "Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.Run(*args, **arguments)

    def SaveWorkspace(self, *args, FileName=None):
        arguments = {"FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.SaveWorkspace(*args, **arguments)

    def SendKeys(self, *args, Keys=None, Wait=None):
        arguments = {"Keys": Keys, "Wait": Wait}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.SendKeys(*args, **arguments)

    def SharePointVersion(self, *args, bstrUrl=None):
        arguments = {"bstrUrl": bstrUrl}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.SharePointVersion(*args, **arguments)

    def Undo(self):
        self.application.Undo()

    def Union(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.Union(*args, **arguments)

    def Volatile(self, *args, Volatile=None):
        arguments = {"Volatile": Volatile}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.Volatile(*args, **arguments)

    def Wait(self, *args, Time=None):
        arguments = {"Time": Time}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.Wait(*args, **arguments)





















































































































































































































































































































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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.areas.Item):
            return self.areas.Item(*args, **arguments)
        else:
            return self.areas.GetItem(*args, **arguments)

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

    def ReplacementList(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.autocorrect.ReplacementList):
            return self.autocorrect.ReplacementList(*args, **arguments)
        else:
            return self.autocorrect.GetReplacementList(*args, **arguments)

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

    def AddReplacement(self, *args, What=None, Replacement=None):
        arguments = {"What": What, "Replacement": Replacement}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.autocorrect.AddReplacement(*args, **arguments)

    def DeleteReplacement(self, *args, What=None):
        arguments = {"What": What}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.autocorrect.DeleteReplacement(*args, **arguments)





















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

    def Item(self, *args, Type=None, AxisGroup=None):
        arguments = {"Type": Type, "AxisGroup": AxisGroup}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.axes.Item(*args, **arguments)







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

    def Characters(self, *args, Start=None, Length=None):
        arguments = {"Start": Start, "Length": Length}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.axistitle.Characters):
            return Characters(self.axistitle.Characters(*args, **arguments))
        else:
            return Characters(self.axistitle.GetCharacters(*args, **arguments))

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.borders.Item):
            return Border(self.borders.Item(*args, **arguments))
        else:
            return Border(self.borders.GetItem(*args, **arguments))

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

    def Add(self, *args, Name=None, Formula=None, UseStandardFormula=None):
        arguments = {"Name": Name, "Formula": Formula, "UseStandardFormula": UseStandardFormula}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return CalculatedField(self.calculatedfields.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return PivotField(self.calculatedfields.Item(*args, **arguments))








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

    def Add(self, *args, Name=None, Formula=None, UseStandardFormula=None):
        arguments = {"Name": Name, "Formula": Formula, "UseStandardFormula": UseStandardFormula}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return CalculatedItem(self.calculateditems.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return PivotItem(self.calculateditems.Item(*args, **arguments))












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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.calculatedmembers.Item):
            return self.calculatedmembers.Item(*args, **arguments)
        else:
            return self.calculatedmembers.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.calculatedmembers.Parent

    def Add(self, *args, Name=None, Formula=None, SolveOrder=None, Type=None, Dynamic=None, DisplayFolder=None, HierarchizeDistinct=None):
        arguments = {"Name": Name, "Formula": Formula, "SolveOrder": SolveOrder, "Type": Type, "Dynamic": Dynamic, "DisplayFolder": DisplayFolder, "HierarchizeDistinct": HierarchizeDistinct}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return CalculatedMember(self.calculatedmembers.Add(*args, **arguments))

















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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return ChartCategory(self.categorycollection.Item(*args, **arguments))
















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

    def Insert(self, *args, String=None):
        arguments = {"String": String}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.characters.Insert(*args, **arguments)





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

    def ChartGroups(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.chart.ChartGroups(*args, **arguments)

    def ChartObjects(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.chart.ChartObjects(*args, **arguments)

    def ChartWizard(self, *args, Source=None, Gallery=None, Format=None, PlotBy=None, CategoryLabels=None, SeriesLabels=None, HasLegend=None, Title=None, CategoryTitle=None, ValueTitle=None, ExtraTitle=None):
        arguments = {"Source": Source, "Gallery": Gallery, "Format": Format, "PlotBy": PlotBy, "CategoryLabels": CategoryLabels, "SeriesLabels": SeriesLabels, "HasLegend": HasLegend, "Title": Title, "CategoryTitle": CategoryTitle, "ValueTitle": ValueTitle, "ExtraTitle": ExtraTitle}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.ChartWizard(*args, **arguments)

    def CheckSpelling(self, *args, CustomDictionary=None, IgnoreUppercase=None, AlwaysSuggest=None, SpellLang=None):
        arguments = {"CustomDictionary": CustomDictionary, "IgnoreUppercase": IgnoreUppercase, "AlwaysSuggest": AlwaysSuggest, "SpellLang": SpellLang}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.CheckSpelling(*args, **arguments)

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

    def Evaluate(self, *args, Name=None):
        arguments = {"Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.chart.Evaluate(*args, **arguments)

    def Export(self, *args, FileName=None, FilterName=None, Interactive=None):
        arguments = {"FileName": FileName, "FilterName": FilterName, "Interactive": Interactive}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.chart.Export(*args, **arguments)

    def ExportAsFixedFormat(self, *args, Type=None, FileName=None, Quality=None, IncludeDocProperties=None, IgnorePrintAreas=None, From=None, To=None, OpenAfterPublish=None, FixedFormatExtClassPtr=None):
        arguments = {"Type": Type, "FileName": FileName, "Quality": Quality, "IncludeDocProperties": IncludeDocProperties, "IgnorePrintAreas": IgnorePrintAreas, "From": From, "To": To, "OpenAfterPublish": OpenAfterPublish, "FixedFormatExtClassPtr": FixedFormatExtClassPtr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.ExportAsFixedFormat(*args, **arguments)

    def GetChartElement(self, *args, x=None, y=None, ElementID=None, Arg1=None, Arg2=None):
        arguments = {"x": x, "y": y, "ElementID": ElementID, "Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.GetChartElement(*args, **arguments)

    def Location(self, *args, Where=None, Name=None):
        arguments = {"Where": Where, "Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.chart.Location(*args, **arguments)

    def Move(self, *args, Before=None, After=None):
        arguments = {"Before": Before, "After": After}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.Move(*args, **arguments)

    def OLEObjects(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.chart.OLEObjects(*args, **arguments)

    def Paste(self, *args, Type=None):
        arguments = {"Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.Paste(*args, **arguments)

    def PrintOut(self, *args, From=None, To=None, Copies=None, Preview=None, ActivePrinter=None, PrintToFile=None, Collate=None, PrToFileName=None, IgnorePrintAreas=None):
        arguments = {"From": From, "To": To, "Copies": Copies, "Preview": Preview, "ActivePrinter": ActivePrinter, "PrintToFile": PrintToFile, "Collate": Collate, "PrToFileName": PrToFileName, "IgnorePrintAreas": IgnorePrintAreas}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.chart.PrintOut(*args, **arguments)

    def PrintPreview(self, *args, EnableChanges=None):
        arguments = {"EnableChanges": EnableChanges}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.PrintPreview(*args, **arguments)

    def Protect(self, *args, Password=None, DrawingObjects=None, Contents=None, Scenarios=None, UserInterfaceOnly=None):
        arguments = {"Password": Password, "DrawingObjects": DrawingObjects, "Contents": Contents, "Scenarios": Scenarios, "UserInterfaceOnly": UserInterfaceOnly}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.Protect(*args, **arguments)

    def Refresh(self):
        self.chart.Refresh()

    def SaveAs(self, *args, FileName=None, FileFormat=None, Password=None, WriteResPassword=None, ReadOnlyRecommended=None, CreateBackup=None, AddToMru=None, TextCodepage=None, TextVisualLayout=None, Local=None):
        arguments = {"FileName": FileName, "FileFormat": FileFormat, "Password": Password, "WriteResPassword": WriteResPassword, "ReadOnlyRecommended": ReadOnlyRecommended, "CreateBackup": CreateBackup, "AddToMru": AddToMru, "TextCodepage": TextCodepage, "TextVisualLayout": TextVisualLayout, "Local": Local}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.SaveAs(*args, **arguments)

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
        return self.chart.SeriesCollection(*args, **arguments)

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
        return self.chart.SetElement(*args, **arguments)

    def SetSourceData(self, *args, Source=None, PlotBy=None):
        arguments = {"Source": Source, "PlotBy": PlotBy}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.SetSourceData(*args, **arguments)

    def Unprotect(self, *args, Password=None):
        arguments = {"Password": Password}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.Unprotect(*args, **arguments)




































































































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

    def SeriesCollection(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.chartgroup.SeriesCollection(*args, **arguments)































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

    def CopyPicture(self, *args, Appearance=None, Format=None):
        arguments = {"Appearance": Appearance, "Format": Format}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.chartobject.CopyPicture(*args, **arguments)

    def Cut(self):
        return self.chartobject.Cut()

    def Delete(self):
        return self.chartobject.Delete()

    def Duplicate(self):
        return self.chartobject.Duplicate()

    def Select(self, *args, Replace=None):
        arguments = {"Replace": Replace}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.chartobject.Select(*args, **arguments)

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

    def Add(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return ChartObject(self.chartobjects.Add(*args, **arguments))

    def Copy(self):
        return self.chartobjects.Copy()

    def CopyPicture(self, *args, Appearance=None, Format=None):
        arguments = {"Appearance": Appearance, "Format": Format}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.chartobjects.CopyPicture(*args, **arguments)

    def Cut(self):
        return self.chartobjects.Cut()

    def Delete(self):
        return self.chartobjects.Delete()

    def Duplicate(self):
        return self.chartobjects.Duplicate()

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.chartobjects.Item(*args, **arguments)

    def Select(self, *args, Replace=None):
        arguments = {"Replace": Replace}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.chartobjects.Select(*args, **arguments)


















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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.charts.Item):
            return self.charts.Item(*args, **arguments)
        else:
            return self.charts.GetItem(*args, **arguments)

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

    def Copy(self, *args, Before=None, After=None):
        arguments = {"Before": Before, "After": After}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.charts.Copy(*args, **arguments)

    def Delete(self):
        self.charts.Delete()

    def Move(self, *args, Before=None, After=None):
        arguments = {"Before": Before, "After": After}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.charts.Move(*args, **arguments)

    def PrintOut(self, *args, From=None, To=None, Copies=None, Preview=None, ActivePrinter=None, PrintToFile=None, Collate=None, PrToFileName=None, IgnorePrintAreas=None):
        arguments = {"From": From, "To": To, "Copies": Copies, "Preview": Preview, "ActivePrinter": ActivePrinter, "PrintToFile": PrintToFile, "Collate": Collate, "PrToFileName": PrToFileName, "IgnorePrintAreas": IgnorePrintAreas}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.charts.PrintOut(*args, **arguments)

    def PrintPreview(self, *args, EnableChanges=None):
        arguments = {"EnableChanges": EnableChanges}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.charts.PrintPreview(*args, **arguments)

    def Select(self, *args, Replace=None):
        arguments = {"Replace": Replace}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.charts.Select(*args, **arguments)









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
        if callable(self.charttitle.Characters):
            return Characters(self.charttitle.Characters(*args, **arguments))
        else:
            return Characters(self.charttitle.GetCharacters(*args, **arguments))

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

    def ModifyAppliesToRange(self, *args, Range=None):
        arguments = {"Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.colorscale.ModifyAppliesToRange(*args, **arguments)

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.colorscalecriteria.Item):
            return ColorScaleCriterion(self.colorscalecriteria.Item(*args, **arguments))
        else:
            return ColorScaleCriterion(self.colorscalecriteria.GetItem(*args, **arguments))




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

    def Add(self, *args, Position=None):
        arguments = {"Position": Position}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return ColorStop(self.colorstops.Add(*args, **arguments))

    def Clear(self):
        return self.colorstops.Clear()

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.colorstops.Item(*args, **arguments)












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

    def Text(self, *args, Text=None, Start=None, Overwrite=None):
        arguments = {"Text": Text, "Start": Start, "Overwrite": Overwrite}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.comment.Text(*args, **arguments)












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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Comment(self.comments.Item(*args, **arguments))







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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return CommentThreaded(self.commentsthreaded.Item(*args, **arguments))









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

    def AddReply(self, *args, Text=None):
        arguments = {"Text": Text}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.commentthreaded.AddReply(*args, **arguments)

    def Delete(self):
        self.commentthreaded.Delete()

    def Next(self):
        return self.commentthreaded.Next()

    def Previous(self):
        return self.commentthreaded.Previous()

    def Text(self, *args, Text=None, Start=None, Overwrite=None):
        arguments = {"Text": Text, "Start": Start, "Overwrite": Overwrite}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.commentthreaded.Text(*args, **arguments)









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

    def Modify(self, *args, NewType=None, NewValue=None):
        arguments = {"NewType": NewType, "NewValue": NewValue}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.conditionvalue.Modify(*args, **arguments)












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

    def Add(self, *args, Name=None, Description=None, ConnectionString=None, CommandText=None, lCmdtype=None, CreateModelConnection=None, ImportRelationships=None):
        arguments = {"Name": Name, "Description": Description, "ConnectionString": ConnectionString, "CommandText": CommandText, "lCmdtype": lCmdtype, "CreateModelConnection": CreateModelConnection, "ImportRelationships": ImportRelationships}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Connection(self.connections.Add(*args, **arguments))

    def AddFromFile(self, *args, FileName=None, CreateModelConnection=None, ImportRelationships=None):
        arguments = {"FileName": FileName, "CreateModelConnection": CreateModelConnection, "ImportRelationships": ImportRelationships}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.connections.AddFromFile(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.connections.Item(*args, **arguments)















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

    def BeginConnect(self, *args, ConnectedShape=None, ConnectionSite=None):
        arguments = {"ConnectedShape": ConnectedShape, "ConnectionSite": ConnectionSite}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.connectorformat.BeginConnect(*args, **arguments)

    def BeginDisconnect(self):
        self.connectorformat.BeginDisconnect()

    def EndConnect(self, *args, ConnectedShape=None, ConnectionSite=None):
        arguments = {"ConnectedShape": ConnectedShape, "ConnectionSite": ConnectionSite}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.connectorformat.EndConnect(*args, **arguments)

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

    def AddItem(self, *args, Text=None, Index=None):
        arguments = {"Text": Text, "Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.controlformat.AddItem(*args, **arguments)

    def List(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.controlformat.List(*args, **arguments)

    def RemoveAllItems(self):
        self.controlformat.RemoveAllItems()

    def RemoveItem(self, *args, Index=None, Count=None):
        arguments = {"Index": Index, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.controlformat.RemoveItem(*args, **arguments)








































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

    def AddMemberPropertyField(self, *args, Property=None, PropertyOrder=None, PropertyDisplayedIn=None):
        arguments = {"Property": Property, "PropertyOrder": PropertyOrder, "PropertyDisplayedIn": PropertyDisplayedIn}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.cubefield.AddMemberPropertyField(*args, **arguments)

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.cubefields.Item):
            return self.cubefields.Item(*args, **arguments)
        else:
            return self.cubefields.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.cubefields.Parent

    def AddSet(self, *args, Name=None, Caption=None):
        arguments = {"Name": Name, "Caption": Caption}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.cubefields.AddSet(*args, **arguments)








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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.customproperties.Item):
            return self.customproperties.Item(*args, **arguments)
        else:
            return self.customproperties.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.customproperties.Parent

    def Add(self, *args, Name=None, Value=None):
        arguments = {"Name": Name, "Value": Value}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return CustomProperty(self.customproperties.Add(*args, **arguments))






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

    def Add(self, *args, ViewName=None, PrintSettings=None, RowColSettings=None):
        arguments = {"ViewName": ViewName, "PrintSettings": PrintSettings, "RowColSettings": RowColSettings}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return CustomView(self.customviews.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return CustomView(self.customviews.Item(*args, **arguments))
















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

    def ModifyAppliesToRange(self, *args, Range=None):
        arguments = {"Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.databar.ModifyAppliesToRange(*args, **arguments)

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

    def Characters(self, *args, Start=None, Length=None):
        arguments = {"Start": Start, "Length": Length}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.datalabel.Characters):
            return Characters(self.datalabel.Characters(*args, **arguments))
        else:
            return Characters(self.datalabel.GetCharacters(*args, **arguments))

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return DataLabel(self.datalabels.Item(*args, **arguments))

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

    def Show(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.dialog.Show(*args, **arguments)








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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.dialogs.Item):
            return self.dialogs.Item(*args, **arguments)
        else:
            return self.dialogs.GetItem(*args, **arguments)

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

    def Characters(self, *args, Start=None, Length=None):
        arguments = {"Start": Start, "Length": Length}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.displayformat.Characters):
            return Characters(self.displayformat.Characters(*args, **arguments))
        else:
            return Characters(self.displayformat.GetCharacters(*args, **arguments))

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

    def Characters(self, *args, Start=None, Length=None):
        arguments = {"Start": Start, "Length": Length}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.displayunitlabel.Characters):
            return Characters(self.displayunitlabel.Characters(*args, **arguments))
        else:
            return Characters(self.displayunitlabel.GetCharacters(*args, **arguments))

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.errors.Item):
            return Error(self.errors.Item(*args, **arguments))
        else:
            return Error(self.errors.GetItem(*args, **arguments))

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.fileexportconverters.Item):
            return FileExportConverter(self.fileexportconverters.Item(*args, **arguments))
        else:
            return FileExportConverter(self.fileexportconverters.GetItem(*args, **arguments))

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.filters.Item):
            return self.filters.Item(*args, **arguments)
        else:
            return self.filters.GetItem(*args, **arguments)

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

    def Modify(self, *args, Type=None, Operator=None, Formula1=None, Formula2=None):
        arguments = {"Type": Type, "Operator": Operator, "Formula1": Formula1, "Formula2": Formula2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.formatcondition.Modify(*args, **arguments)

    def ModifyAppliesToRange(self, *args, Range=None):
        arguments = {"Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.formatcondition.ModifyAppliesToRange(*args, **arguments)

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

    def Add(self, *args, Type=None, Operator=None, Formula1=None, Formula2=None):
        arguments = {"Type": Type, "Operator": Operator, "Formula1": Formula1, "Formula2": Formula2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return FormatCondition(self.formatconditions.Add(*args, **arguments))

    def AddAboveAverage(self):
        return self.formatconditions.AddAboveAverage()

    def AddColorScale(self, *args, ColorScaleType=None):
        arguments = {"ColorScaleType": ColorScaleType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.formatconditions.AddColorScale(*args, **arguments)

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.formatconditions.Item(*args, **arguments)








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

    def AddNodes(self, *args, SegmentType=None, EditingType=None, X1=None, Y1=None, X2=None, Y2=None, X3=None, Y3=None):
        arguments = {"SegmentType": SegmentType, "EditingType": EditingType, "X1": X1, "Y1": Y1, "X2": X2, "Y2": Y2, "X3": X3, "Y3": Y3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.freeformbuilder.AddNodes(*args, **arguments)

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

    def Range(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.groupshapes.Range):
            return ShapeRange(self.groupshapes.Range(*args, **arguments))
        else:
            return ShapeRange(self.groupshapes.GetRange(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Shape(self.groupshapes.Item(*args, **arguments))


























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

    def DragOff(self, *args, Direction=None, RegionIndex=None):
        arguments = {"Direction": Direction, "RegionIndex": RegionIndex}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.hpagebreak.DragOff(*args, **arguments)









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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.hpagebreaks.Item):
            return self.hpagebreaks.Item(*args, **arguments)
        else:
            return self.hpagebreaks.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.hpagebreaks.Parent

    def Add(self, *args, Before=None):
        arguments = {"Before": Before}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return HPageBreak(self.hpagebreaks.Add(*args, **arguments))











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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.hyperlinks.Item):
            return self.hyperlinks.Item(*args, **arguments)
        else:
            return self.hyperlinks.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.hyperlinks.Parent

    def Add(self, *args, Anchor=None, Address=None, SubAddress=None, ScreenTip=None, TextToDisplay=None):
        arguments = {"Anchor": Anchor, "Address": Address, "SubAddress": SubAddress, "ScreenTip": ScreenTip, "TextToDisplay": TextToDisplay}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Hyperlink(self.hyperlinks.Add(*args, **arguments))

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.iconcriteria.Item):
            return IconCriterion(self.iconcriteria.Item(*args, **arguments))
        else:
            return IconCriterion(self.iconcriteria.GetItem(*args, **arguments))




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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.iconset.Item):
            return Icon(self.iconset.Item(*args, **arguments))
        else:
            return Icon(self.iconset.GetItem(*args, **arguments))

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

    def ModifyAppliesToRange(self, *args, Range=None):
        arguments = {"Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.iconsetcondition.ModifyAppliesToRange(*args, **arguments)

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.iconsets.Item):
            return IconSet(self.iconsets.Item(*args, **arguments))
        else:
            return IconSet(self.iconsets.GetItem(*args, **arguments))

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

    def ConnectData(self, *args, TopicID=None, Strings=None, GetNewValues=None):
        arguments = {"TopicID": TopicID, "Strings": Strings, "GetNewValues": GetNewValues}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.irtdserver.ConnectData(*args, **arguments)

    def DisconnectData(self, *args, TopicID=None):
        arguments = {"TopicID": TopicID}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.irtdserver.DisconnectData(*args, **arguments)

    def Heartbeat(self):
        return self.irtdserver.Heartbeat()

    def RefreshData(self, *args, TopicCount=None):
        arguments = {"TopicCount": TopicCount}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.irtdserver.RefreshData(*args, **arguments)

    def ServerStart(self, *args, CallbackObject=None):
        arguments = {"CallbackObject": CallbackObject}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.irtdserver.ServerStart(*args, **arguments)

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

    def LegendEntries(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.legend.LegendEntries(*args, **arguments)

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.listcolumns.Item):
            return self.listcolumns.Item(*args, **arguments)
        else:
            return self.listcolumns.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.listcolumns.Parent

    def Add(self, *args, Position=None):
        arguments = {"Position": Position}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return ListColumn(self.listcolumns.Add(*args, **arguments))













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

    def Publish(self, *args, Target=None, LinkSource=None):
        arguments = {"Target": Target, "LinkSource": LinkSource}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.listobject.Publish(*args, **arguments)

    def Refresh(self):
        self.listobject.Refresh()

    def Resize(self, *args, Range=None):
        arguments = {"Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.listobject.Resize(*args, **arguments)

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.listobjects.Item):
            return self.listobjects.Item(*args, **arguments)
        else:
            return self.listobjects.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.listobjects.Parent

    def Add(self, *args, SourceType=None, Source=None, LinkSource=None, XlListObjectHasHeaders=None, Destination=None, TableStyleName=None):
        arguments = {"SourceType": SourceType, "Source": Source, "LinkSource": LinkSource, "XlListObjectHasHeaders": XlListObjectHasHeaders, "Destination": Destination, "TableStyleName": TableStyleName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return ListObject(self.listobjects.Add(*args, **arguments))







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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.listrows.Item):
            return self.listrows.Item(*args, **arguments)
        else:
            return self.listrows.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.listrows.Parent

    def Add(self, *args, Position=None, AlwaysInsert=None):
        arguments = {"Position": Position, "AlwaysInsert": AlwaysInsert}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return ListRow(self.listrows.Add(*args, **arguments))








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

    def Add(self, *args, Name=None, RefersTo=None, Visible=None, MacroType=None, ShortcutKey=None, Category=None, NameLocal=None, RefersToLocal=None, CategoryLocal=None, RefersToR1C1=None, RefersToR1C1Local=None):
        arguments = {"Name": Name, "RefersTo": RefersTo, "Visible": Visible, "MacroType": MacroType, "ShortcutKey": ShortcutKey, "Category": Category, "NameLocal": NameLocal, "RefersToLocal": RefersToLocal, "CategoryLocal": CategoryLocal, "RefersToR1C1": RefersToR1C1, "RefersToR1C1Local": RefersToR1C1Local}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Name(self.names.Add(*args, **arguments))

    def Item(self, *args, Index=None, IndexLocal=None, RefersTo=None):
        arguments = {"Index": Index, "IndexLocal": IndexLocal, "RefersTo": RefersTo}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.names.Item(*args, **arguments)









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

    def SaveAsODC(self, *args, ODCFileName=None, Description=None, Keywords=None):
        arguments = {"ODCFileName": ODCFileName, "Description": Description, "Keywords": Keywords}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.odbcconnection.SaveAsODC(*args, **arguments)



















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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return ODBCError(self.odbcerrors.Item(*args, **arguments))





















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

    def SaveAsODC(self, *args, ODCFileName=None, Description=None, Keywords=None):
        arguments = {"ODCFileName": ODCFileName, "Description": Description, "Keywords": Keywords}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.oledbconnection.SaveAsODC(*args, **arguments)


























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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return OLEDBError(self.oledberrors.Item(*args, **arguments))






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

    def Verb(self, *args, Verb=None):
        arguments = {"Verb": Verb}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.oleformat.Verb(*args, **arguments)





























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

    def CopyPicture(self, *args, Appearance=None, Format=None):
        arguments = {"Appearance": Appearance, "Format": Format}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.oleobject.CopyPicture(*args, **arguments)

    def Cut(self):
        return self.oleobject.Cut()

    def Delete(self):
        return self.oleobject.Delete()

    def Duplicate(self):
        return self.oleobject.Duplicate()

    def Select(self, *args, Replace=None):
        arguments = {"Replace": Replace}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.oleobject.Select(*args, **arguments)

    def SendToBack(self):
        return self.oleobject.SendToBack()

    def Update(self):
        return self.oleobject.Update()

    def Verb(self, *args, Verb=None):
        arguments = {"Verb": Verb}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.oleobject.Verb(*args, **arguments)







































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

    def Add(self, *args, ClassType=None, FileName=None, Link=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Left=None, Top=None, Width=None, Height=None):
        arguments = {"ClassType": ClassType, "FileName": FileName, "Link": Link, "DisplayAsIcon": DisplayAsIcon, "IconFileName": IconFileName, "IconIndex": IconIndex, "IconLabel": IconLabel, "Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return OLEObject(self.oleobjects.Add(*args, **arguments))

    def BringToFront(self):
        return self.oleobjects.BringToFront()

    def Copy(self):
        return self.oleobjects.Copy()

    def CopyPicture(self, *args, Appearance=None, Format=None):
        arguments = {"Appearance": Appearance, "Format": Format}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.oleobjects.CopyPicture(*args, **arguments)

    def Cut(self):
        return self.oleobjects.Cut()

    def Delete(self):
        return self.oleobjects.Delete()

    def Duplicate(self):
        return self.oleobjects.Duplicate()

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.oleobjects.Item(*args, **arguments)

    def Select(self, *args, Replace=None):
        arguments = {"Replace": Replace}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.oleobjects.Select(*args, **arguments)

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

    def ShowLevels(self, *args, RowLevels=None, ColumnLevels=None):
        arguments = {"RowLevels": RowLevels, "ColumnLevels": ColumnLevels}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.outline.ShowLevels(*args, **arguments)












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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.pages.Item):
            return Page(self.pages.Item(*args, **arguments))
        else:
            return Page(self.pages.GetItem(*args, **arguments))



























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

    def LargeScroll(self, *args, Down=None, Up=None, ToRight=None, ToLeft=None):
        arguments = {"Down": Down, "Up": Up, "ToRight": ToRight, "ToLeft": ToLeft}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.pane.LargeScroll(*args, **arguments)

    def PointsToScreenPixelsX(self, *args, Points=None):
        arguments = {"Points": Points}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.pane.PointsToScreenPixelsX(*args, **arguments)

    def PointsToScreenPixelsY(self, *args, Points=None):
        arguments = {"Points": Points}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.pane.PointsToScreenPixelsY(*args, **arguments)

    def ScrollIntoView(self, *args, Left=None, Top=None, Width=None, Height=None, Start=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height, "Start": Start}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pane.ScrollIntoView(*args, **arguments)

    def SmallScroll(self, *args, Down=None, Up=None, ToRight=None, ToLeft=None):
        arguments = {"Down": Down, "Up": Up, "ToRight": ToRight, "ToLeft": ToLeft}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.pane.SmallScroll(*args, **arguments)














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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.panes.Item):
            return self.panes.Item(*args, **arguments)
        else:
            return self.panes.GetItem(*args, **arguments)

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

    def SetParam(self, *args, Type=None, Value=None):
        arguments = {"Type": Type, "Value": Value}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.parameter.SetParam(*args, **arguments)
















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

    def Add(self, *args, Name=None, iDataType=None):
        arguments = {"Name": Name, "iDataType": iDataType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Parameter(self.parameters.Add(*args, **arguments))

    def Delete(self):
        self.parameters.Delete()

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Parameter(self.parameters.Item(*args, **arguments))
















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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.phonetics.Item):
            return self.phonetics.Item(*args, **arguments)
        else:
            return self.phonetics.GetItem(*args, **arguments)

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

    def Add(self, *args, Start=None, Length=None, Text=None):
        arguments = {"Start": Start, "Length": Length, "Text": Text}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.phonetics.Add(*args, **arguments)

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

    def IncrementBrightness(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pictureformat.IncrementBrightness(*args, **arguments)

    def IncrementContrast(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pictureformat.IncrementContrast(*args, **arguments)











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

    def CreatePivotTable(self, *args, TableDestination=None, TableName=None, ReadData=None, DefaultVersion=None):
        arguments = {"TableDestination": TableDestination, "TableName": TableName, "ReadData": ReadData, "DefaultVersion": DefaultVersion}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.pivotcache.CreatePivotTable(*args, **arguments)

    def MakeConnection(self):
        self.pivotcache.MakeConnection()

    def Refresh(self):
        self.pivotcache.Refresh()

    def ResetTimer(self):
        self.pivotcache.ResetTimer()

    def SaveAsODC(self, *args, ODCFileName=None, Description=None, Keywords=None):
        arguments = {"ODCFileName": ODCFileName, "Description": Description, "Keywords": Keywords}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pivotcache.SaveAsODC(*args, **arguments)
































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

    def Create(self, *args, SourceType=None, SourceData=None, Version=None):
        arguments = {"SourceType": SourceType, "SourceData": SourceData, "Version": Version}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.pivotcaches.Create(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return PivotCache(self.pivotcaches.Item(*args, **arguments))












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

    def ChildItems(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.pivotfield.ChildItems):
            return PivotItem(self.pivotfield.ChildItems(*args, **arguments))
        else:
            return PivotItem(self.pivotfield.GetChildItems(*args, **arguments))

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

    def HiddenItems(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.pivotfield.HiddenItems):
            return PivotItem(self.pivotfield.HiddenItems(*args, **arguments))
        else:
            return PivotItem(self.pivotfield.GetHiddenItems(*args, **arguments))

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

    def ParentItems(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.pivotfield.ParentItems):
            return PivotItem(self.pivotfield.ParentItems(*args, **arguments))
        else:
            return PivotItem(self.pivotfield.GetParentItems(*args, **arguments))

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

    def VisibleItems(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.pivotfield.VisibleItems):
            return PivotItem(self.pivotfield.VisibleItems(*args, **arguments))
        else:
            return PivotItem(self.pivotfield.GetVisibleItems(*args, **arguments))

    @property
    def VisibleItemsList(self):
        return self.pivotfield.VisibleItemsList

    @VisibleItemsList.setter
    def VisibleItemsList(self, value):
        self.pivotfield.VisibleItemsList = value

    def AddPageItem(self, *args, Item=None, ClearList=None):
        arguments = {"Item": Item, "ClearList": ClearList}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pivotfield.AddPageItem(*args, **arguments)

    def AutoShow(self, *args, Type=None, Range=None, Count=None, Field=None):
        arguments = {"Type": Type, "Range": Range, "Count": Count, "Field": Field}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pivotfield.AutoShow(*args, **arguments)

    def AutoSort(self, *args, Order=None, Field=None, PivotLine=None, CustomSubtotal=None):
        arguments = {"Order": Order, "Field": Field, "PivotLine": PivotLine, "CustomSubtotal": CustomSubtotal}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pivotfield.AutoSort(*args, **arguments)

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

    def DrillTo(self, *args, PivotFieldName=None):
        arguments = {"PivotFieldName": PivotFieldName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pivotfield.DrillTo(*args, **arguments)

    def PivotItems(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.pivotfield.PivotItems(*args, **arguments)


































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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.pivotfields.Item(*args, **arguments)












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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.pivotfilters.Item):
            return PivotFilters(self.pivotfilters.Item(*args, **arguments))
        else:
            return PivotFilters(self.pivotfilters.GetItem(*args, **arguments))

    @property
    def Parent(self):
        return PivotFilters(self.pivotfilters.Parent)

    def Add(self, *args, Type=None, DataField=None, Value1=None, Value2=None, Order=None, Name=None, Description=None, MemberPropertyField=None, WholeDayFilter=None):
        arguments = {"Type": Type, "DataField": DataField, "Value1": Value1, "Value2": Value2, "Order": Order, "Name": Name, "Description": Description, "MemberPropertyField": MemberPropertyField, "WholeDayFilter": WholeDayFilter}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.pivotfilters.Add(*args, **arguments)








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

    def Add(self, *args, Formula=None, UseStandardFormula=None):
        arguments = {"Formula": Formula, "UseStandardFormula": UseStandardFormula}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return PivotFormula(self.pivotformulas.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return PivotFormula(self.pivotformulas.Item(*args, **arguments))














class PivotItem:

    def __init__(self, pivotitem=None):
        self.pivotitem = pivotitem

    @property
    def Application(self):
        return self.pivotitem.Application

    @property
    def Caption(self):
        return self.pivotitem.Caption

    def ChildItems(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.pivotitem.ChildItems):
            return PivotItem(self.pivotitem.ChildItems(*args, **arguments))
        else:
            return PivotItem(self.pivotitem.GetChildItems(*args, **arguments))

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

    def DrillTo(self, *args, PivotItemName=None):
        arguments = {"PivotItemName": PivotItemName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pivotitem.DrillTo(*args, **arguments)


















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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return PivotItem(self.pivotitemlist.Item(*args, **arguments))








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

    def Add(self, *args, Name=None):
        arguments = {"Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pivotitems.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.pivotitems.Item(*args, **arguments)





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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.pivotlinecells.Item):
            return PivotLineCells(self.pivotlinecells.Item(*args, **arguments))
        else:
            return PivotLineCells(self.pivotlinecells.GetItem(*args, **arguments))

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.pivotlines.Item):
            return PivotLines(self.pivotlines.Item(*args, **arguments))
        else:
            return PivotLines(self.pivotlines.GetItem(*args, **arguments))

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

    def ColumnFields(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.pivottable.ColumnFields):
            return PivotField(self.pivottable.ColumnFields(*args, **arguments))
        else:
            return PivotField(self.pivottable.GetColumnFields(*args, **arguments))

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

    def DataFields(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.pivottable.DataFields):
            return PivotField(self.pivottable.DataFields(*args, **arguments))
        else:
            return PivotField(self.pivottable.GetDataFields(*args, **arguments))

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

    def HiddenFields(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.pivottable.HiddenFields):
            return PivotField(self.pivottable.HiddenFields(*args, **arguments))
        else:
            return PivotField(self.pivottable.GetHiddenFields(*args, **arguments))

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

    def PageFields(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.pivottable.PageFields):
            return PivotField(self.pivottable.PageFields(*args, **arguments))
        else:
            return PivotField(self.pivottable.GetPageFields(*args, **arguments))

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

    def RowFields(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.pivottable.RowFields):
            return PivotField(self.pivottable.RowFields(*args, **arguments))
        else:
            return PivotField(self.pivottable.GetRowFields(*args, **arguments))

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

    def VisibleFields(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.pivottable.VisibleFields):
            return PivotField(self.pivottable.VisibleFields(*args, **arguments))
        else:
            return PivotField(self.pivottable.GetVisibleFields(*args, **arguments))

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

    def AddDataField(self, *args, Field=None, Caption=None, Function=None):
        arguments = {"Field": Field, "Caption": Caption, "Function": Function}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.pivottable.AddDataField(*args, **arguments)

    def AddFields(self, *args, RowFields=None, ColumnFields=None, PageFields=None, AddToTable=None):
        arguments = {"RowFields": RowFields, "ColumnFields": ColumnFields, "PageFields": PageFields, "AddToTable": AddToTable}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.pivottable.AddFields(*args, **arguments)

    def AllocateChanges(self):
        return self.pivottable.AllocateChanges()

    def CalculatedFields(self):
        return self.pivottable.CalculatedFields()

    def ChangeConnection(self, *args, conn=None):
        arguments = {"conn": conn}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pivottable.ChangeConnection(*args, **arguments)

    def ChangePivotCache(self, *args, bstr=None):
        arguments = {"bstr": bstr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pivottable.ChangePivotCache(*args, **arguments)

    def ClearAllFilters(self):
        self.pivottable.ClearAllFilters()

    def ClearTable(self):
        self.pivottable.ClearTable()

    def CommitChanges(self):
        return self.pivottable.CommitChanges()

    def ConvertToFormulas(self, *args, ConvertFilters=None):
        arguments = {"ConvertFilters": ConvertFilters}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pivottable.ConvertToFormulas(*args, **arguments)

    def CreateCubeFile(self, *args, File=None, Measures=None, Levels=None, Members=None, Properties=None):
        arguments = {"File": File, "Measures": Measures, "Levels": Levels, "Members": Members, "Properties": Properties}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.pivottable.CreateCubeFile(*args, **arguments)

    def DiscardChanges(self):
        return self.pivottable.DiscardChanges()

    def GetData(self, *args, Name=None):
        arguments = {"Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.pivottable.GetData(*args, **arguments)

    def GetPivotData(self, *args, DataField=None, Field1=None, Item1=None, Field2=None, Item2=None, Field3=None, Item3=None, Field4=None, Item4=None, Field5=None, Item5=None, Field6=None, Item6=None, Field7=None, Item7=None, Field8=None, Item8=None, Field9=None, Item9=None, Field10=None, Item10=None, Field11=None, Item11=None, Field12=None, Item12=None, Field13=None, Item13=None, Field14=None, Item14=None):
        arguments = {"DataField": DataField, "Field1": Field1, "Item1": Item1, "Field2": Field2, "Item2": Item2, "Field3": Field3, "Item3": Item3, "Field4": Field4, "Item4": Item4, "Field5": Field5, "Item5": Item5, "Field6": Field6, "Item6": Item6, "Field7": Field7, "Item7": Item7, "Field8": Field8, "Item8": Item8, "Field9": Field9, "Item9": Item9, "Field10": Field10, "Item10": Item10, "Field11": Field11, "Item11": Item11, "Field12": Field12, "Item12": Item12, "Field13": Field13, "Item13": Item13, "Field14": Field14, "Item14": Item14}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.pivottable.GetPivotData(*args, **arguments)

    def ListFormulas(self):
        self.pivottable.ListFormulas()

    def PivotCache(self):
        return self.pivottable.PivotCache()

    def PivotFields(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.pivottable.PivotFields(*args, **arguments)

    def PivotSelect(self, *args, Name=None, Mode=None, UseStandardName=None):
        arguments = {"Name": Name, "Mode": Mode, "UseStandardName": UseStandardName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pivottable.PivotSelect(*args, **arguments)

    def PivotTableWizard(self, *args, SourceType=None, SourceData=None, TableDestination=None, TableName=None, RowGrand=None, ColumnGrand=None, SaveData=None, HasAutoFormat=None, AutoPage=None, Reserved=None, BackgroundQuery=None, OptimizeCache=None, PageFieldOrder=None, PageFieldWrapCount=None, ReadData=None, Connection=None):
        arguments = {"SourceType": SourceType, "SourceData": SourceData, "TableDestination": TableDestination, "TableName": TableName, "RowGrand": RowGrand, "ColumnGrand": ColumnGrand, "SaveData": SaveData, "HasAutoFormat": HasAutoFormat, "AutoPage": AutoPage, "Reserved": Reserved, "BackgroundQuery": BackgroundQuery, "OptimizeCache": OptimizeCache, "PageFieldOrder": PageFieldOrder, "PageFieldWrapCount": PageFieldWrapCount, "ReadData": ReadData, "Connection": Connection}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pivottable.PivotTableWizard(*args, **arguments)

    def RefreshDataSourceValues(self):
        return self.pivottable.RefreshDataSourceValues()

    def RefreshTable(self):
        return self.pivottable.RefreshTable()

    def RepeatAllLabels(self, *args, Repeat=None):
        arguments = {"Repeat": Repeat}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.pivottable.RepeatAllLabels(*args, **arguments)

    def RowAxisLayout(self, *args, RowLayout=None):
        arguments = {"RowLayout": RowLayout}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pivottable.RowAxisLayout(*args, **arguments)

    def ShowPages(self, *args, PageField=None):
        arguments = {"PageField": PageField}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.pivottable.ShowPages(*args, **arguments)

    def SubtotalLocation(self, *args, Location=None):
        arguments = {"Location": Location}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pivottable.SubtotalLocation(*args, **arguments)

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.pivottablechangelist.Item):
            return ValueChange(self.pivottablechangelist.Item(*args, **arguments))
        else:
            return ValueChange(self.pivottablechangelist.GetItem(*args, **arguments))

    @property
    def Parent(self):
        return PivotTable(self.pivottablechangelist.Parent)

    def Add(self, *args, Tuple=None, Value=None, AllocationValue=None, AllocationMethod=None, AllocationWeightExpression=None):
        arguments = {"Tuple": Tuple, "Value": Value, "AllocationValue": AllocationValue, "AllocationMethod": AllocationMethod, "AllocationWeightExpression": AllocationWeightExpression}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.pivottablechangelist.Add(*args, **arguments)








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

    def Add(self, *args, PivotCache=None, TableDestination=None, TableName=None, ReadData=None, DefaultVersion=None):
        arguments = {"PivotCache": PivotCache, "TableDestination": TableDestination, "TableName": TableName, "ReadData": ReadData, "DefaultVersion": DefaultVersion}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return PivotTable(self.pivottables.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return PivotTable(self.pivottables.Item(*args, **arguments))



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

    def ApplyDataLabels(self, *args, Type=None, LegendKey=None, AutoText=None, HasLeaderLines=None, ShowSeriesName=None, ShowCategoryName=None, ShowValue=None, ShowPercentage=None, ShowBubbleSize=None, Separator=None):
        arguments = {"Type": Type, "LegendKey": LegendKey, "AutoText": AutoText, "HasLeaderLines": HasLeaderLines, "ShowSeriesName": ShowSeriesName, "ShowCategoryName": ShowCategoryName, "ShowValue": ShowValue, "ShowPercentage": ShowPercentage, "ShowBubbleSize": ShowBubbleSize, "Separator": Separator}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.point.ApplyDataLabels(*args, **arguments)

    def ClearFormats(self):
        return self.point.ClearFormats()

    def Copy(self):
        return self.point.Copy()

    def Delete(self):
        return self.point.Delete()

    def Paste(self):
        return self.point.Paste()

    def PieSliceLocation(self, *args, loc=None, Index=None):
        arguments = {"loc": loc, "Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.point.PieSliceLocation(*args, **arguments)

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Point(self.points.Item(*args, **arguments))



















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

    def Edit(self, *args, WriteResPassword=None, UpdateLinks=None):
        arguments = {"WriteResPassword": WriteResPassword, "UpdateLinks": UpdateLinks}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Workbook(self.protectedviewwindow.Edit(*args, **arguments))













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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.protectedviewwindows.Item):
            return self.protectedviewwindows.Item(*args, **arguments)
        else:
            return self.protectedviewwindows.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.protectedviewwindows.Parent

    def Open(self, *args, FileName=None, Password=None, AddToMru=None, RepairMode=None):
        arguments = {"FileName": FileName, "Password": Password, "AddToMru": AddToMru, "RepairMode": RepairMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return ProtectedViewWindow(self.protectedviewwindows.Open(*args, **arguments))
















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

    def Publish(self, *args, Create=None):
        arguments = {"Create": Create}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.publishobject.Publish(*args, **arguments)














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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.publishobjects.Item):
            return self.publishobjects.Item(*args, **arguments)
        else:
            return self.publishobjects.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.publishobjects.Parent

    def Add(self, *args, SourceType=None, FileName=None, Sheet=None, Source=None, HtmlType=None, DivID=None, Title=None):
        arguments = {"SourceType": SourceType, "FileName": FileName, "Sheet": Sheet, "Source": Source, "HtmlType": HtmlType, "DivID": DivID, "Title": Title}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return PublishObject(self.publishobjects.Add(*args, **arguments))

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

    def Refresh(self, *args, BackgroundQuery=None):
        arguments = {"BackgroundQuery": BackgroundQuery}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.querytable.Refresh(*args, **arguments)

    def ResetTimer(self):
        self.querytable.ResetTimer()

    def SaveAsODC(self, *args, ODCFileName=None, Description=None, Keywords=None):
        arguments = {"ODCFileName": ODCFileName, "Description": Description, "Keywords": Keywords}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.querytable.SaveAsODC(*args, **arguments)
























































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

    def Add(self, *args, Connection=None, Destination=None, Sql=None):
        arguments = {"Connection": Connection, "Destination": Destination, "Sql": Sql}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return QueryTable(self.querytables.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return QueryTable(self.querytables.Item(*args, **arguments))






class Range:

    def __init__(self, range=None):
        self.range = range

    @property
    def AddIndent(self):
        return self.range.AddIndent

    @AddIndent.setter
    def AddIndent(self, value):
        self.range.AddIndent = value

    def Address(self, *args, RowAbsolute=None, ColumnAbsolute=None, ReferenceStyle=None, External=None, RelativeTo=None):
        arguments = {"RowAbsolute": RowAbsolute, "ColumnAbsolute": ColumnAbsolute, "ReferenceStyle": ReferenceStyle, "External": External, "RelativeTo": RelativeTo}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.range.Address):
            return self.range.Address(*args, **arguments)
        else:
            return self.range.GetAddress(*args, **arguments)

    def AddressLocal(self, *args, RowAbsolute=None, ColumnAbsolute=None, ReferenceStyle=None, External=None, RelativeTo=None):
        arguments = {"RowAbsolute": RowAbsolute, "ColumnAbsolute": ColumnAbsolute, "ReferenceStyle": ReferenceStyle, "External": External, "RelativeTo": RelativeTo}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.range.AddressLocal):
            return self.range.AddressLocal(*args, **arguments)
        else:
            return self.range.GetAddressLocal(*args, **arguments)

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

    def Cells(self, *args, RowIndex=None, ColumnIndex=None):
        arguments = {"RowIndex": RowIndex, "ColumnIndex": ColumnIndex}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.range.Cells):
            return Range(self.range.Cells(*args, **arguments))
        else:
            return Range(self.range.GetCells(*args, **arguments))

    def Characters(self, *args, Start=None, Length=None):
        arguments = {"Start": Start, "Length": Length}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.range.Characters):
            return Characters(self.range.Characters(*args, **arguments))
        else:
            return Characters(self.range.GetCharacters(*args, **arguments))

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

    def End(self, *args, Direction=None):
        arguments = {"Direction": Direction}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.range.End):
            return Range(self.range.End(*args, **arguments))
        else:
            return Range(self.range.GetEnd(*args, **arguments))

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

    def Item(self, *args, RowIndex=None, ColumnIndex=None):
        arguments = {"RowIndex": RowIndex, "ColumnIndex": ColumnIndex}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.range.Item):
            return Range(self.range.Item(*args, **arguments))
        else:
            return Range(self.range.GetItem(*args, **arguments))

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

    def Offset(self, *args, RowOffset=None, ColumnOffset=None):
        arguments = {"RowOffset": RowOffset, "ColumnOffset": ColumnOffset}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.range.Offset):
            return Range(self.range.Offset(*args, **arguments))
        else:
            return Range(self.range.GetOffset(*args, **arguments))

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

    def Range(self, *args, Cell1=None, Cell2=None):
        arguments = {"Cell1": Cell1, "Cell2": Cell2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.range.Range):
            return Range(self.range.Range(*args, **arguments))
        else:
            return Range(self.range.GetRange(*args, **arguments))

    @property
    def ReadingOrder(self):
        return self.range.ReadingOrder

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.range.ReadingOrder = value

    def Resize(self, *args, RowSize=None, ColumnSize=None):
        arguments = {"RowSize": RowSize, "ColumnSize": ColumnSize}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.range.Resize):
            return Range(self.range.Resize(*args, **arguments))
        else:
            return Range(self.range.GetResize(*args, **arguments))

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

    def AddComment(self, *args, Text=None):
        arguments = {"Text": Text}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.AddComment(*args, **arguments)

    def AddCommentThreaded(self, *args, Text=None):
        arguments = {"Text": Text}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.AddCommentThreaded(*args, **arguments)

    def AdvancedFilter(self, *args, Action=None, CriteriaRange=None, CopyToRange=None, Unique=None):
        arguments = {"Action": Action, "CriteriaRange": CriteriaRange, "CopyToRange": CopyToRange, "Unique": Unique}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.AdvancedFilter(*args, **arguments)

    def AllocateChanges(self):
        self.range.AllocateChanges()

    def ApplyNames(self, *args, Names=None, IgnoreRelativeAbsolute=None, UseRowColumnNames=None, OmitColumn=None, OmitRow=None, Order=None, AppendLast=None):
        arguments = {"Names": Names, "IgnoreRelativeAbsolute": IgnoreRelativeAbsolute, "UseRowColumnNames": UseRowColumnNames, "OmitColumn": OmitColumn, "OmitRow": OmitRow, "Order": Order, "AppendLast": AppendLast}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.ApplyNames(*args, **arguments)

    def ApplyOutlineStyles(self):
        return self.range.ApplyOutlineStyles()

    def AutoComplete(self, *args, String=None):
        arguments = {"String": String}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.AutoComplete(*args, **arguments)

    def AutoFill(self, *args, Destination=None, Type=None):
        arguments = {"Destination": Destination, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.AutoFill(*args, **arguments)

    def AutoFilter(self, *args, Field=None, Criteria1=None, Operator=None, Criteria2=None, SubField=None, VisibleDropDown=None):
        arguments = {"Field": Field, "Criteria1": Criteria1, "Operator": Operator, "Criteria2": Criteria2, "SubField": SubField, "VisibleDropDown": VisibleDropDown}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.AutoFilter(*args, **arguments)

    def AutoFit(self):
        return self.range.AutoFit()

    def AutoOutline(self):
        return self.range.AutoOutline()

    def BorderAround(self, *args, LineStyle=None, Weight=None, ColorIndex=None, Color=None, ThemeColor=None):
        arguments = {"LineStyle": LineStyle, "Weight": Weight, "ColorIndex": ColorIndex, "Color": Color, "ThemeColor": ThemeColor}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.BorderAround(*args, **arguments)

    def Calculate(self):
        return self.range.Calculate()

    def CalculateRowMajorOrder(self):
        return self.range.CalculateRowMajorOrder()

    def CheckSpelling(self, *args, CustomDictionary=None, IgnoreUppercase=None, AlwaysSuggest=None, SpellLang=None):
        arguments = {"CustomDictionary": CustomDictionary, "IgnoreUppercase": IgnoreUppercase, "AlwaysSuggest": AlwaysSuggest, "SpellLang": SpellLang}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.CheckSpelling(*args, **arguments)

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

    def ColumnDifferences(self, *args, Comparison=None):
        arguments = {"Comparison": Comparison}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.ColumnDifferences(*args, **arguments)

    def Consolidate(self, *args, Sources=None, Function=None, TopRow=None, LeftColumn=None, CreateLinks=None):
        arguments = {"Sources": Sources, "Function": Function, "TopRow": TopRow, "LeftColumn": LeftColumn, "CreateLinks": CreateLinks}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.Consolidate(*args, **arguments)

    def ConvertToLinkedDataType(self, *args, ServiceID=None, LanguageCulture=None):
        arguments = {"ServiceID": ServiceID, "LanguageCulture": LanguageCulture}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.ConvertToLinkedDataType(*args, **arguments)

    def Copy(self, *args, Destination=None):
        arguments = {"Destination": Destination}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.Copy(*args, **arguments)

    def CopyFromRecordset(self, *args, Data=None, MaxRows=None, MaxColumns=None):
        arguments = {"Data": Data, "MaxRows": MaxRows, "MaxColumns": MaxColumns}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.CopyFromRecordset(*args, **arguments)

    def CopyPicture(self, *args, Appearance=None, Format=None):
        arguments = {"Appearance": Appearance, "Format": Format}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.CopyPicture(*args, **arguments)

    def CreateNames(self, *args, Top=None, Left=None, Bottom=None, Right=None):
        arguments = {"Top": Top, "Left": Left, "Bottom": Bottom, "Right": Right}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.CreateNames(*args, **arguments)

    def Cut(self, *args, Destination=None):
        arguments = {"Destination": Destination}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.Cut(*args, **arguments)

    def DataSeries(self, *args, Rowcol=None, Type=None, Date=None, Step=None, Stop=None, Trend=None):
        arguments = {"Rowcol": Rowcol, "Type": Type, "Date": Date, "Step": Step, "Stop": Stop, "Trend": Trend}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.DataSeries(*args, **arguments)

    def DataTypeToText(self):
        self.range.DataTypeToText()

    def Delete(self, *args, Shift=None):
        arguments = {"Shift": Shift}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.Delete(*args, **arguments)

    def DialogBox(self):
        return self.range.DialogBox()

    def Dirty(self):
        self.range.Dirty()

    def DiscardChanges(self):
        self.range.DiscardChanges()

    def EditionOptions(self, *args, Type=None, Option=None, Name=None, Reference=None, Appearance=None, ChartSize=None, Format=None):
        arguments = {"Type": Type, "Option": Option, "Name": Name, "Reference": Reference, "Appearance": Appearance, "ChartSize": ChartSize, "Format": Format}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.EditionOptions(*args, **arguments)

    def ExportAsFixedFormat(self, *args, Type=None, FileName=None, Quality=None, IncludeDocProperties=None, IgnorePrintAreas=None, From=None, To=None, OpenAfterPublish=None, FixedFormatExtClassPtr=None):
        arguments = {"Type": Type, "FileName": FileName, "Quality": Quality, "IncludeDocProperties": IncludeDocProperties, "IgnorePrintAreas": IgnorePrintAreas, "From": From, "To": To, "OpenAfterPublish": OpenAfterPublish, "FixedFormatExtClassPtr": FixedFormatExtClassPtr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.ExportAsFixedFormat(*args, **arguments)

    def FillDown(self):
        return self.range.FillDown()

    def FillLeft(self):
        return self.range.FillLeft()

    def FillRight(self):
        return self.range.FillRight()

    def FillUp(self):
        return self.range.FillUp()

    def Find(self, *args, What=None, After=None, LookIn=None, LookAt=None, SearchOrder=None, SearchDirection=None, MatchCase=None, MatchByte=None, SearchFormat=None):
        arguments = {"What": What, "After": After, "LookIn": LookIn, "LookAt": LookAt, "SearchOrder": SearchOrder, "SearchDirection": SearchDirection, "MatchCase": MatchCase, "MatchByte": MatchByte, "SearchFormat": SearchFormat}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.Find(*args, **arguments)

    def FindNext(self, *args, After=None):
        arguments = {"After": After}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.FindNext(*args, **arguments)

    def FindPrevious(self, *args, Before=None):
        arguments = {"Before": Before}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.FindPrevious(*args, **arguments)

    def FunctionWizard(self):
        return self.range.FunctionWizard()

    def Group(self, *args, Start=None, End=None, By=None, Periods=None):
        arguments = {"Start": Start, "End": End, "By": By, "Periods": Periods}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.Group(*args, **arguments)

    def Insert(self, *args, Shift=None, CopyOrigin=None):
        arguments = {"Shift": Shift, "CopyOrigin": CopyOrigin}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.Insert(*args, **arguments)

    def InsertIndent(self, *args, InsertAmount=None):
        arguments = {"InsertAmount": InsertAmount}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.InsertIndent(*args, **arguments)

    def Justify(self):
        return self.range.Justify()

    def ListNames(self):
        return self.range.ListNames()

    def Merge(self, *args, Across=None):
        arguments = {"Across": Across}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.Merge(*args, **arguments)

    def NavigateArrow(self, *args, TowardPrecedent=None, ArrowNumber=None, LinkNumber=None):
        arguments = {"TowardPrecedent": TowardPrecedent, "ArrowNumber": ArrowNumber, "LinkNumber": LinkNumber}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.NavigateArrow(*args, **arguments)

    def NoteText(self, *args, Text=None, Start=None, Length=None):
        arguments = {"Text": Text, "Start": Start, "Length": Length}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.NoteText(*args, **arguments)

    def Parse(self, *args, ParseLine=None, Destination=None):
        arguments = {"ParseLine": ParseLine, "Destination": Destination}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.Parse(*args, **arguments)

    def PasteSpecial(self, *args, Paste=None, Operation=None, SkipBlanks=None, Transpose=None):
        arguments = {"Paste": Paste, "Operation": Operation, "SkipBlanks": SkipBlanks, "Transpose": Transpose}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.PasteSpecial(*args, **arguments)

    def PrintOut(self, *args, From=None, To=None, Copies=None, Preview=None, ActivePrinter=None, PrintToFile=None, Collate=None, PrToFileName=None):
        arguments = {"From": From, "To": To, "Copies": Copies, "Preview": Preview, "ActivePrinter": ActivePrinter, "PrintToFile": PrintToFile, "Collate": Collate, "PrToFileName": PrToFileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.PrintOut(*args, **arguments)

    def PrintPreview(self, *args, EnableChanges=None):
        arguments = {"EnableChanges": EnableChanges}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.PrintPreview(*args, **arguments)

    def RemoveDuplicates(self, *args, Columns=None, Header=None):
        arguments = {"Columns": Columns, "Header": Header}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.RemoveDuplicates(*args, **arguments)

    def RemoveSubtotal(self):
        return self.range.RemoveSubtotal()

    def Replace(self, *args, What=None, Replacement=None, LookAt=None, SearchOrder=None, MatchCase=None, MatchByte=None, SearchFormat=None, ReplaceFormat=None, FormulaVersion=None):
        arguments = {"What": What, "Replacement": Replacement, "LookAt": LookAt, "SearchOrder": SearchOrder, "MatchCase": MatchCase, "MatchByte": MatchByte, "SearchFormat": SearchFormat, "ReplaceFormat": ReplaceFormat, "FormulaVersion": FormulaVersion}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.Replace(*args, **arguments)

    def ResetContents(self):
        self.range.ResetContents()

    def RowDifferences(self, *args, Comparison=None):
        arguments = {"Comparison": Comparison}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.RowDifferences(*args, **arguments)

    def Run(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.Run(*args, **arguments)

    def Select(self):
        return self.range.Select()

    def SetCellDataTypeFromCell(self, *args, Range=None, LanguageCulture=None):
        arguments = {"Range": Range, "LanguageCulture": LanguageCulture}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.SetCellDataTypeFromCell(*args, **arguments)

    def SetPhonetic(self):
        self.range.SetPhonetic()

    def Show(self):
        return self.range.Show()

    def ShowCard(self):
        self.range.ShowCard()

    def ShowDependents(self, *args, Remove=None):
        arguments = {"Remove": Remove}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.ShowDependents(*args, **arguments)

    def ShowErrors(self):
        return self.range.ShowErrors()

    def ShowPrecedents(self, *args, Remove=None):
        arguments = {"Remove": Remove}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.ShowPrecedents(*args, **arguments)

    def Sort(self, *args, Key1=None, Order1=None, Key2=None, Type=None, Order2=None, Key3=None, Order3=None, Header=None, OrderCustom=None, MatchCase=None, Orientation=None, SortMethod=None, DataOption1=None, DataOption2=None, DataOption3=None):
        arguments = {"Key1": Key1, "Order1": Order1, "Key2": Key2, "Type": Type, "Order2": Order2, "Key3": Key3, "Order3": Order3, "Header": Header, "OrderCustom": OrderCustom, "MatchCase": MatchCase, "Orientation": Orientation, "SortMethod": SortMethod, "DataOption1": DataOption1, "DataOption2": DataOption2, "DataOption3": DataOption3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.Sort(*args, **arguments)

    def SortSpecial(self, *args, SortMethod=None, Key1=None, Order1=None, Type=None, Key2=None, Order2=None, Key3=None, Order3=None, Header=None, OrderCustom=None, MatchCase=None, Orientation=None, DataOption1=None, DataOption2=None, DataOption3=None):
        arguments = {"SortMethod": SortMethod, "Key1": Key1, "Order1": Order1, "Type": Type, "Key2": Key2, "Order2": Order2, "Key3": Key3, "Order3": Order3, "Header": Header, "OrderCustom": OrderCustom, "MatchCase": MatchCase, "Orientation": Orientation, "DataOption1": DataOption1, "DataOption2": DataOption2, "DataOption3": DataOption3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.SortSpecial(*args, **arguments)

    def Speak(self, *args, SpeakDirection=None, SpeakFormulas=None):
        arguments = {"SpeakDirection": SpeakDirection, "SpeakFormulas": SpeakFormulas}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.Speak(*args, **arguments)

    def SpecialCells(self, *args, Type=None, Value=None):
        arguments = {"Type": Type, "Value": Value}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.SpecialCells(*args, **arguments)

    def SubscribeTo(self, *args, Edition=None, Format=None):
        arguments = {"Edition": Edition, "Format": Format}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.SubscribeTo(*args, **arguments)

    def Subtotal(self, *args, GroupBy=None, Function=None, TotalList=None, Replace=None, PageBreaks=None, SummaryBelowData=None):
        arguments = {"GroupBy": GroupBy, "Function": Function, "TotalList": TotalList, "Replace": Replace, "PageBreaks": PageBreaks, "SummaryBelowData": SummaryBelowData}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.Subtotal(*args, **arguments)

    def Table(self, *args, RowInput=None, ColumnInput=None):
        arguments = {"RowInput": RowInput, "ColumnInput": ColumnInput}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.Table(*args, **arguments)

    def TextToColumns(self, *args, Destination=None, DataType=None, TextQualifier=None, ConsecutiveDelimiter=None, Tab=None, Semicolon=None, Comma=None, Space=None, Other=None, OtherChar=None, FieldInfo=None, DecimalSeparator=None, ThousandsSeparator=None, TrailingMinusNumbers=None):
        arguments = {"Destination": Destination, "DataType": DataType, "TextQualifier": TextQualifier, "ConsecutiveDelimiter": ConsecutiveDelimiter, "Tab": Tab, "Semicolon": Semicolon, "Comma": Comma, "Space": Space, "Other": Other, "OtherChar": OtherChar, "FieldInfo": FieldInfo, "DecimalSeparator": DecimalSeparator, "ThousandsSeparator": ThousandsSeparator, "TrailingMinusNumbers": TrailingMinusNumbers}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.TextToColumns(*args, **arguments)

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.ranges.Item):
            return Range(self.ranges.Item(*args, **arguments))
        else:
            return Range(self.ranges.GetItem(*args, **arguments))

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.recentfiles.Item):
            return self.recentfiles.Item(*args, **arguments)
        else:
            return self.recentfiles.GetItem(*args, **arguments)

    @property
    def Maximum(self):
        return self.recentfiles.Maximum

    @Maximum.setter
    def Maximum(self, value):
        self.recentfiles.Maximum = value

    @property
    def Parent(self):
        return self.recentfiles.Parent

    def Add(self, *args, Name=None):
        arguments = {"Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return RecentFile(self.recentfiles.Add(*args, **arguments))






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

    def Values(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.scenario.Values):
            return self.scenario.Values(*args, **arguments)
        else:
            return self.scenario.GetValues(*args, **arguments)

    def ChangeScenario(self, *args, ChangingCells=None, Values=None):
        arguments = {"ChangingCells": ChangingCells, "Values": Values}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.scenario.ChangeScenario(*args, **arguments)

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

    def Add(self, *args, Name=None, ChangingCells=None, Values=None, Comment=None, Locked=None, Hidden=None):
        arguments = {"Name": Name, "ChangingCells": ChangingCells, "Values": Values, "Comment": Comment, "Locked": Locked, "Hidden": Hidden}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Scenario(self.scenarios.Add(*args, **arguments))

    def CreateSummary(self, *args, ReportType=None, ResultCells=None):
        arguments = {"ReportType": ReportType, "ResultCells": ResultCells}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.scenarios.CreateSummary(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Scenario(self.scenarios.Item(*args, **arguments))

    def Merge(self, *args, Source=None):
        arguments = {"Source": Source}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.scenarios.Merge(*args, **arguments)








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

    def ApplyDataLabels(self, *args, Type=None, LegendKey=None, AutoText=None, HasLeaderLines=None, ShowSeriesName=None, ShowCategoryName=None, ShowValue=None, ShowPercentage=None, ShowBubbleSize=None, Separator=None):
        arguments = {"Type": Type, "LegendKey": LegendKey, "AutoText": AutoText, "HasLeaderLines": HasLeaderLines, "ShowSeriesName": ShowSeriesName, "ShowCategoryName": ShowCategoryName, "ShowValue": ShowValue, "ShowPercentage": ShowPercentage, "ShowBubbleSize": ShowBubbleSize, "Separator": Separator}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.series.ApplyDataLabels(*args, **arguments)

    def ClearFormats(self):
        return self.series.ClearFormats()

    def Copy(self):
        return self.series.Copy()

    def DataLabels(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.series.DataLabels(*args, **arguments)

    def Delete(self):
        return self.series.Delete()

    def ErrorBar(self, *args, Direction=None, Include=None, Type=None, Amount=None, MinusValues=None):
        arguments = {"Direction": Direction, "Include": Include, "Type": Type, "Amount": Amount, "MinusValues": MinusValues}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.series.ErrorBar(*args, **arguments)

    def GeoMappingLevel(self):
        self.series.GeoMappingLevel()

    def GeoProjectionType(self):
        self.series.GeoProjectionType()

    def Paste(self):
        return self.series.Paste()

    def Points(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.series.Points(*args, **arguments)

    def RegionLabelOptions(self):
        self.series.RegionLabelOptions()

    def Select(self):
        return self.series.Select()

    def Trendlines(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.series.Trendlines(*args, **arguments)
























































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

    def Add(self, *args, Source=None, Rowcol=None, SeriesLabels=None, CategoryLabels=None, Replace=None):
        arguments = {"Source": Source, "Rowcol": Rowcol, "SeriesLabels": SeriesLabels, "CategoryLabels": CategoryLabels, "Replace": Replace}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Series(self.seriescollection.Add(*args, **arguments))

    def Extend(self, *args, Source=None, RowCol=None, CategoryLabels=None):
        arguments = {"Source": Source, "RowCol": RowCol, "CategoryLabels": CategoryLabels}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.seriescollection.Extend(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Series(self.seriescollection.Item(*args, **arguments))

    def NewSeries(self):
        return self.seriescollection.NewSeries()

    def Paste(self, *args, RowCol=None, SeriesLabels=None, CategoryLabels=None, Replace=None, NewSeries=None):
        arguments = {"RowCol": RowCol, "SeriesLabels": SeriesLabels, "CategoryLabels": CategoryLabels, "Replace": Replace, "NewSeries": NewSeries}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.seriescollection.Paste(*args, **arguments)











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

    def Add(self, *args, Obj=None):
        arguments = {"Obj": Obj}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return ServerViewableItem(self.serverviewableitems.Add(*args, **arguments))

    def Delete(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.serverviewableitems.Delete(*args, **arguments)

    def DeleteAll(self):
        self.serverviewableitems.DeleteAll()

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.serverviewableitems.Item(*args, **arguments)












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

    def CopyPicture(self, *args, Appearance=None, Format=None):
        arguments = {"Appearance": Appearance, "Format": Format}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shape.CopyPicture(*args, **arguments)

    def Cut(self):
        return self.shape.Cut()

    def Delete(self):
        self.shape.Delete()

    def Duplicate(self):
        return self.shape.Duplicate()

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

    def RerouteConnections(self):
        self.shape.RerouteConnections()

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
        return ShapeRange(self.shape.Ungroup())

    def ZOrder(self, *args, ZOrderCmd=None):
        arguments = {"ZOrderCmd": ZOrderCmd}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shape.ZOrder(*args, **arguments)









































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

    def Delete(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shapenodes.Delete(*args, **arguments)

    def Insert(self, *args, Index=None, SegmentType=None, EditingType=None, X1=None, Y1=None, X2=None, Y2=None, X3=None, Y3=None):
        arguments = {"Index": Index, "SegmentType": SegmentType, "EditingType": EditingType, "X1": X1, "Y1": Y1, "X2": X2, "Y2": Y2, "X3": X3, "Y3": Y3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shapenodes.Insert(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return ShapeNode(self.shapenodes.Item(*args, **arguments))

    def SetEditingType(self, *args, Index=None, EditingType=None):
        arguments = {"Index": Index, "EditingType": EditingType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shapenodes.SetEditingType(*args, **arguments)

    def SetPosition(self, *args, Index=None, X1=None, Y1=None):
        arguments = {"Index": Index, "X1": X1, "Y1": Y1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shapenodes.SetPosition(*args, **arguments)

    def SetSegmentType(self, *args, Index=None, SegmentType=None):
        arguments = {"Index": Index, "SegmentType": SegmentType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shapenodes.SetSegmentType(*args, **arguments)






























































































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

    def Range(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.shapes.Range):
            return ShapeRange(self.shapes.Range(*args, **arguments))
        else:
            return ShapeRange(self.shapes.GetRange(*args, **arguments))

    def Add3DModel(self, *args, FileName=None, LinkToFile=None, SaveWithDocument=None, Left=None, Top=None, Width=None, Height=None):
        arguments = {"FileName": FileName, "LinkToFile": LinkToFile, "SaveWithDocument": SaveWithDocument, "Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.Add3DModel(*args, **arguments)

    def AddCallout(self, *args, Type=None, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Type": Type, "Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddCallout(*args, **arguments)

    def AddConnector(self, *args, Type=None, BeginX=None, BeginY=None, EndX=None, EndY=None):
        arguments = {"Type": Type, "BeginX": BeginX, "BeginY": BeginY, "EndX": EndX, "EndY": EndY}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddConnector(*args, **arguments)

    def AddCurve(self, *args, SafeArrayOfPoints=None):
        arguments = {"SafeArrayOfPoints": SafeArrayOfPoints}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddCurve(*args, **arguments)

    def AddFormControl(self, *args, Type=None, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Type": Type, "Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddFormControl(*args, **arguments)

    def AddLabel(self, *args, Orientation=None, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Orientation": Orientation, "Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddLabel(*args, **arguments)

    def AddLine(self, *args, BeginX=None, BeginY=None, EndX=None, EndY=None):
        arguments = {"BeginX": BeginX, "BeginY": BeginY, "EndX": EndX, "EndY": EndY}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddLine(*args, **arguments)

    def AddOLEObject(self, *args, ClassType=None, FileName=None, Link=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Left=None, Top=None, Width=None, Height=None):
        arguments = {"ClassType": ClassType, "FileName": FileName, "Link": Link, "DisplayAsIcon": DisplayAsIcon, "IconFileName": IconFileName, "IconIndex": IconIndex, "IconLabel": IconLabel, "Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddOLEObject(*args, **arguments)

    def AddPicture(self, *args, FileName=None, LinkToFile=None, SaveWithDocument=None, Left=None, Top=None, Width=None, Height=None):
        arguments = {"FileName": FileName, "LinkToFile": LinkToFile, "SaveWithDocument": SaveWithDocument, "Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddPicture(*args, **arguments)

    def AddPolyline(self, *args, SafeArrayOfPoints=None):
        arguments = {"SafeArrayOfPoints": SafeArrayOfPoints}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddPolyline(*args, **arguments)

    def AddShape(self, *args, Type=None, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Type": Type, "Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddShape(*args, **arguments)

    def AddSmartArt(self, *args, Layout=None, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Layout": Layout, "Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddSmartArt(*args, **arguments)

    def AddTextbox(self, *args, Orientation=None, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Orientation": Orientation, "Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddTextbox(*args, **arguments)

    def AddTextEffect(self, *args, PresetTextEffect=None, Text=None, FontName=None, FontSize=None, FontBold=None, FontItalic=None, Left=None, Top=None):
        arguments = {"PresetTextEffect": PresetTextEffect, "Text": Text, "FontName": FontName, "FontSize": FontSize, "FontBold": FontBold, "FontItalic": FontItalic, "Left": Left, "Top": Top}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddTextEffect(*args, **arguments)

    def BuildFreeform(self, *args, EditingType=None, X1=None, Y1=None):
        arguments = {"EditingType": EditingType, "X1": X1, "Y1": Y1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.BuildFreeform(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Shape(self.shapes.Item(*args, **arguments))

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.sheets.Item):
            return self.sheets.Item(*args, **arguments)
        else:
            return self.sheets.GetItem(*args, **arguments)

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

    def Add(self, *args, Before=None, After=None, Count=None, Type=None):
        arguments = {"Before": Before, "After": After, "Count": Count, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Sheet(self.sheets.Add(*args, **arguments))

    def Copy(self, *args, Before=None, After=None):
        arguments = {"Before": Before, "After": After}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sheets.Copy(*args, **arguments)

    def Delete(self):
        self.sheets.Delete()

    def FillAcrossSheets(self, *args, Range=None, Type=None):
        arguments = {"Range": Range, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sheets.FillAcrossSheets(*args, **arguments)

    def Move(self, *args, Before=None, After=None):
        arguments = {"Before": Before, "After": After}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sheets.Move(*args, **arguments)

    def PrintOut(self, *args, From=None, To=None, Copies=None, Preview=None, ActivePrinter=None, PrintToFile=None, Collate=None, PrToFileName=None, IgnorePrintAreas=None):
        arguments = {"From": From, "To": To, "Copies": Copies, "Preview": Preview, "ActivePrinter": ActivePrinter, "PrintToFile": PrintToFile, "Collate": Collate, "PrToFileName": PrToFileName, "IgnorePrintAreas": IgnorePrintAreas}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sheets.PrintOut(*args, **arguments)

    def PrintPreview(self, *args, EnableChanges=None):
        arguments = {"EnableChanges": EnableChanges}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sheets.PrintPreview(*args, **arguments)

    def Select(self, *args, Replace=None):
        arguments = {"Replace": Replace}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sheets.Select(*args, **arguments)













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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.sheetviews.Item):
            return self.sheetviews.Item(*args, **arguments)
        else:
            return self.sheetviews.GetItem(*args, **arguments)

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

    def Item(self, *args, Level=None):
        arguments = {"Level": Level}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.slicercachelevels.Item):
            return SlicerCacheLevel(self.slicercachelevels.Item(*args, **arguments))
        else:
            return SlicerCacheLevel(self.slicercachelevels.GetItem(*args, **arguments))

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.slicercaches.Item):
            return SlicerCache(self.slicercaches.Item(*args, **arguments))
        else:
            return SlicerCache(self.slicercaches.GetItem(*args, **arguments))

    @property
    def Parent(self):
        return Workbook(self.slicercaches.Parent)

    def Add(self, *args, Source=None, SourceField=None, Name=None, SlicerCacheType=None):
        arguments = {"Source": Source, "SourceField": SourceField, "Name": Name, "SlicerCacheType": SlicerCacheType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.slicercaches.Add(*args, **arguments)







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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.sliceritems.Item):
            return SlicerItem(self.sliceritems.Item(*args, **arguments))
        else:
            return SlicerItem(self.sliceritems.GetItem(*args, **arguments))

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.slicerpivottables.Item):
            return PivotTable(self.slicerpivottables.Item(*args, **arguments))
        else:
            return PivotTable(self.slicerpivottables.GetItem(*args, **arguments))

    @property
    def Parent(self):
        return SlicerCache(self.slicerpivottables.Parent)

    def AddPivotTable(self, *args, PivotTable=None):
        arguments = {"PivotTable": PivotTable}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.slicerpivottables.AddPivotTable(*args, **arguments)

    def RemovePivotTable(self, *args, PivotTable=None):
        arguments = {"PivotTable": PivotTable}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.slicerpivottables.RemovePivotTable(*args, **arguments)









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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.slicers.Item):
            return Slicer(self.slicers.Item(*args, **arguments))
        else:
            return Slicer(self.slicers.GetItem(*args, **arguments))

    @property
    def Parent(self):
        return SlicerCache(self.slicers.Parent)

    def Add(self, *args, SlicerDestination=None, Level=None, Name=None, Caption=None, Top=None, Left=None, Width=None, Height=None):
        arguments = {"SlicerDestination": SlicerDestination, "Level": Level, "Name": Name, "Caption": Caption, "Top": Top, "Left": Left, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Slicer(self.slicers.Add(*args, **arguments))










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

    def SetRange(self, *args, Rng=None):
        arguments = {"Rng": Rng}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sort.SetRange(*args, **arguments)














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

    def ModifyKey(self, *args, Key=None):
        arguments = {"Key": Key}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sortfield.ModifyKey(*args, **arguments)

    def SetIcon(self, *args, Icon=None):
        arguments = {"Icon": Icon}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sortfield.SetIcon(*args, **arguments)
















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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.sortfields.Item):
            return SortField(self.sortfields.Item(*args, **arguments))
        else:
            return SortField(self.sortfields.GetItem(*args, **arguments))

    @property
    def Parent(self):
        return self.sortfields.Parent

    def Add(self, *args, Key=None, SortOn=None, Order=None, CustomOrder=None, DataOption=None):
        arguments = {"Key": Key, "SortOn": SortOn, "Order": Order, "CustomOrder": CustomOrder, "DataOption": DataOption}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sortfields.Add(*args, **arguments)

    def Add2(self, *args, Key=None, SortOn=None, Order=None, CustomOrder=None, DataOption=None, SubField=None):
        arguments = {"Key": Key, "SortOn": SortOn, "Order": Order, "CustomOrder": CustomOrder, "DataOption": DataOption, "SubField": SubField}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sortfields.Add2(*args, **arguments)

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

    def ModifyLocation(self, *args, Range=None):
        arguments = {"Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sparkline.ModifyLocation(*args, **arguments)

    def ModifySourceData(self, *args, Formula=None):
        arguments = {"Formula": Formula}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sparkline.ModifySourceData(*args, **arguments)
















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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.sparklinegroup.Item):
            return Sparkline(self.sparklinegroup.Item(*args, **arguments))
        else:
            return Sparkline(self.sparklinegroup.GetItem(*args, **arguments))

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

    def Modify(self, *args, Location=None, SourceData=None):
        arguments = {"Location": Location, "SourceData": SourceData}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sparklinegroup.Modify(*args, **arguments)

    def ModifyDateRange(self, *args, DateRange=None):
        arguments = {"DateRange": DateRange}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sparklinegroup.ModifyDateRange(*args, **arguments)

    def ModifyLocation(self, *args, Location=None):
        arguments = {"Location": Location}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sparklinegroup.ModifyLocation(*args, **arguments)

    def ModifySourceData(self, *args, SourceData=None):
        arguments = {"SourceData": SourceData}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sparklinegroup.ModifySourceData(*args, **arguments)


















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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.sparklinegroups.Item):
            return SparklineGroup(self.sparklinegroups.Item(*args, **arguments))
        else:
            return SparklineGroup(self.sparklinegroups.GetItem(*args, **arguments))

    @property
    def Parent(self):
        return Range(self.sparklinegroups.Parent)

    def Add(self, *args, Type=None, SourceData=None):
        arguments = {"Type": Type, "SourceData": SourceData}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sparklinegroups.Add(*args, **arguments)

    def Clear(self):
        return self.sparklinegroups.Clear()

    def ClearGroups(self):
        return self.sparklinegroups.ClearGroups()

    def Group(self, *args, Location=None):
        arguments = {"Location": Location}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sparklinegroups.Group(*args, **arguments)

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

    def Speak(self, *args, Text=None, SpeakAsync=None, SpeakXML=None, Purge=None):
        arguments = {"Text": Text, "SpeakAsync": SpeakAsync, "SpeakXML": SpeakXML, "Purge": Purge}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.speech.Speak(*args, **arguments)


















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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.styles.Item):
            return self.styles.Item(*args, **arguments)
        else:
            return self.styles.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.styles.Parent

    def Add(self, *args, Name=None):
        arguments = {"Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Style(self.styles.Add(*args, **arguments))

    def Merge(self, *args, Workbook=None):
        arguments = {"Workbook": Workbook}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.styles.Merge(*args, **arguments)











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

    def Duplicate(self, *args, NewTableStyleName=None):
        arguments = {"NewTableStyleName": NewTableStyleName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.tablestyle.Duplicate(*args, **arguments)
















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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return TableStyleElement(self.tablestyleelements.Item(*args, **arguments))








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

    def Add(self, *args, TableStyleName=None):
        arguments = {"TableStyleName": TableStyleName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.tablestyles.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.tablestyles.Item(*args, **arguments)












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

    def Characters(self, *args, Start=None, Length=None):
        arguments = {"Start": Start, "Length": Length}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.textframe.Characters(*args, **arguments)


















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

    def ModifyAppliesToRange(self, *args, Range=None):
        arguments = {"Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.top10.ModifyAppliesToRange(*args, **arguments)

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

    def Add(self, *args, Type=None, Order=None, Period=None, Forward=None, Backward=None, Intercept=None, DisplayEquation=None, DisplayRSquared=None, Name=None):
        arguments = {"Type": Type, "Order": Order, "Period": Period, "Forward": Forward, "Backward": Backward, "Intercept": Intercept, "DisplayEquation": DisplayEquation, "DisplayRSquared": DisplayRSquared, "Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Trendline(self.trendlines.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Trendline(self.trendlines.Item(*args, **arguments))




















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

    def ModifyAppliesToRange(self, *args, Range=None):
        arguments = {"Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.uniquevalues.ModifyAppliesToRange(*args, **arguments)

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.usedobjects.Item):
            return self.usedobjects.Item(*args, **arguments)
        else:
            return self.usedobjects.GetItem(*args, **arguments)

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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.useraccesslist.Item):
            return self.useraccesslist.Item(*args, **arguments)
        else:
            return self.useraccesslist.GetItem(*args, **arguments)

    def Add(self, *args, Name=None, AllowEdit=None):
        arguments = {"Name": Name, "AllowEdit": AllowEdit}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return UserAccess(self.useraccesslist.Add(*args, **arguments))

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

    def Add(self, *args, Type=None, AlertStyle=None, Operator=None, Formula1=None, Formula2=None):
        arguments = {"Type": Type, "AlertStyle": AlertStyle, "Operator": Operator, "Formula1": Formula1, "Formula2": Formula2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.validation.Add(*args, **arguments)

    def Delete(self):
        self.validation.Delete()

    def Modify(self, *args, Type=None, AlertStyle=None, Operator=None, Formula1=None, Formula2=None):
        arguments = {"Type": Type, "AlertStyle": AlertStyle, "Operator": Operator, "Formula1": Formula1, "Formula2": Formula2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.validation.Modify(*args, **arguments)















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

    def DragOff(self, *args, Direction=None, RegionIndex=None):
        arguments = {"Direction": Direction, "RegionIndex": RegionIndex}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.vpagebreak.DragOff(*args, **arguments)









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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.vpagebreaks.Item):
            return self.vpagebreaks.Item(*args, **arguments)
        else:
            return self.vpagebreaks.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.vpagebreaks.Parent

    def Add(self, *args, Before=None):
        arguments = {"Before": Before}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return VPageBreak(self.vpagebreaks.Add(*args, **arguments))



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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.watches.Item):
            return self.watches.Item(*args, **arguments)
        else:
            return self.watches.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.watches.Parent

    def Add(self, *args, Source=None):
        arguments = {"Source": Source}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Watch(self.watches.Add(*args, **arguments))

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

    def Close(self, *args, SaveChanges=None, FileName=None, RouteWorkbook=None):
        arguments = {"SaveChanges": SaveChanges, "FileName": FileName, "RouteWorkbook": RouteWorkbook}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.window.Close(*args, **arguments)

    def LargeScroll(self, *args, Down=None, Up=None, ToRight=None, ToLeft=None):
        arguments = {"Down": Down, "Up": Up, "ToRight": ToRight, "ToLeft": ToLeft}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.window.LargeScroll(*args, **arguments)

    def NewWindow(self):
        return self.window.NewWindow()

    def PointsToScreenPixelsX(self, *args, Points=None):
        arguments = {"Points": Points}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.window.PointsToScreenPixelsX(*args, **arguments)

    def PointsToScreenPixelsY(self, *args, Points=None):
        arguments = {"Points": Points}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.window.PointsToScreenPixelsY(*args, **arguments)

    def PrintOut(self, *args, From=None, To=None, Copies=None, Preview=None, ActivePrinter=None, PrintToFile=None, Collate=None, PrToFileName=None):
        arguments = {"From": From, "To": To, "Copies": Copies, "Preview": Preview, "ActivePrinter": ActivePrinter, "PrintToFile": PrintToFile, "Collate": Collate, "PrToFileName": PrToFileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.window.PrintOut(*args, **arguments)

    def PrintPreview(self, *args, EnableChanges=None):
        arguments = {"EnableChanges": EnableChanges}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.window.PrintPreview(*args, **arguments)

    def RangeFromPoint(self, *args, x=None, y=None):
        arguments = {"x": x, "y": y}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.window.RangeFromPoint(*args, **arguments)

    def ScrollIntoView(self, *args, Left=None, Top=None, Width=None, Height=None, Start=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height, "Start": Start}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.window.ScrollIntoView(*args, **arguments)

    def ScrollWorkbookTabs(self, *args, Sheets=None, Position=None):
        arguments = {"Sheets": Sheets, "Position": Position}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.window.ScrollWorkbookTabs(*args, **arguments)

    def SmallScroll(self, *args, Down=None, Up=None, ToRight=None, ToLeft=None):
        arguments = {"Down": Down, "Up": Up, "ToRight": ToRight, "ToLeft": ToLeft}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.window.SmallScroll(*args, **arguments)












































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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.windows.Item):
            return self.windows.Item(*args, **arguments)
        else:
            return self.windows.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.windows.Parent

    @property
    def SyncScrollingSideBySide(self):
        return self.windows.SyncScrollingSideBySide

    def Arrange(self, *args, ArrangeStyle=None, ActiveWorkbook=None, SyncHorizontal=None, SyncVertical=None):
        arguments = {"ArrangeStyle": ArrangeStyle, "ActiveWorkbook": ActiveWorkbook, "SyncHorizontal": SyncHorizontal, "SyncVertical": SyncVertical}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.windows.Arrange(*args, **arguments)

    def BreakSideBySide(self):
        return self.windows.BreakSideBySide()

    def CompareSideBySideWith(self, *args, WindowName=None):
        arguments = {"WindowName": WindowName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.windows.CompareSideBySideWith(*args, **arguments)

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

    def AcceptAllChanges(self, *args, When=None, Who=None, Where=None):
        arguments = {"When": When, "Who": Who, "Where": Where}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.AcceptAllChanges(*args, **arguments)

    def Activate(self):
        self.workbook.Activate()

    def AddToFavorites(self):
        self.workbook.AddToFavorites()

    def ApplyTheme(self, *args, FileName=None):
        arguments = {"FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.ApplyTheme(*args, **arguments)

    def BreakLink(self, *args, Name=None, Type=None):
        arguments = {"Name": Name, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.BreakLink(*args, **arguments)

    def CanCheckIn(self):
        return self.workbook.CanCheckIn()

    def ChangeFileAccess(self, *args, Mode=None, WritePassword=None, Notify=None):
        arguments = {"Mode": Mode, "WritePassword": WritePassword, "Notify": Notify}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.ChangeFileAccess(*args, **arguments)

    def ChangeLink(self, *args, Name=None, NewName=None, Type=None):
        arguments = {"Name": Name, "NewName": NewName, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.ChangeLink(*args, **arguments)

    def CheckIn(self, *args, SaveChanges=None, Comments=None, MakePublic=None):
        arguments = {"SaveChanges": SaveChanges, "Comments": Comments, "MakePublic": MakePublic}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.CheckIn(*args, **arguments)

    def CheckInWithVersion(self, *args, SaveChanges=None, Comments=None, MakePublic=None, VersionType=None):
        arguments = {"SaveChanges": SaveChanges, "Comments": Comments, "MakePublic": MakePublic, "VersionType": VersionType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.workbook.CheckInWithVersion(*args, **arguments)

    def Close(self, *args, SaveChanges=None, FileName=None, RouteWorkbook=None):
        arguments = {"SaveChanges": SaveChanges, "FileName": FileName, "RouteWorkbook": RouteWorkbook}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.Close(*args, **arguments)

    def ConvertComments(self):
        self.workbook.ConvertComments()

    def DeleteNumberFormat(self, *args, NumberFormat=None):
        arguments = {"NumberFormat": NumberFormat}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.DeleteNumberFormat(*args, **arguments)

    def EnableConnections(self):
        self.workbook.EnableConnections()

    def EndReview(self):
        self.workbook.EndReview()

    def ExclusiveAccess(self):
        return self.workbook.ExclusiveAccess()

    def ExportAsFixedFormat(self, *args, Type=None, FileName=None, Quality=None, IncludeDocProperties=None, IgnorePrintAreas=None, From=None, To=None, OpenAfterPublish=None, FixedFormatExtClassPtr=None):
        arguments = {"Type": Type, "FileName": FileName, "Quality": Quality, "IncludeDocProperties": IncludeDocProperties, "IgnorePrintAreas": IgnorePrintAreas, "From": From, "To": To, "OpenAfterPublish": OpenAfterPublish, "FixedFormatExtClassPtr": FixedFormatExtClassPtr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.ExportAsFixedFormat(*args, **arguments)

    def FollowHyperlink(self, *args, Address=None, SubAddress=None, NewWindow=None, AddHistory=None, ExtraInfo=None, Method=None, HeaderInfo=None):
        arguments = {"Address": Address, "SubAddress": SubAddress, "NewWindow": NewWindow, "AddHistory": AddHistory, "ExtraInfo": ExtraInfo, "Method": Method, "HeaderInfo": HeaderInfo}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.FollowHyperlink(*args, **arguments)

    def ForwardMailer(self):
        self.workbook.ForwardMailer()

    def GetWorkflowTasks(self):
        return self.workbook.GetWorkflowTasks()

    def GetWorkflowTemplates(self):
        return self.workbook.GetWorkflowTemplates()

    def HighlightChangesOptions(self, *args, When=None, Who=None, Where=None):
        arguments = {"When": When, "Who": Who, "Where": Where}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.HighlightChangesOptions(*args, **arguments)

    def LinkInfo(self, *args, Name=None, LinkInfo=None, Type=None, EditionRef=None):
        arguments = {"Name": Name, "LinkInfo": LinkInfo, "Type": Type, "EditionRef": EditionRef}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.workbook.LinkInfo(*args, **arguments)

    def LinkSources(self, *args, Type=None):
        arguments = {"Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.workbook.LinkSources(*args, **arguments)

    def LockServerFile(self):
        self.workbook.LockServerFile()

    def MergeWorkbook(self, *args, FileName=None):
        arguments = {"FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.MergeWorkbook(*args, **arguments)

    def NewWindow(self):
        return self.workbook.NewWindow()

    def OpenLinks(self, *args, Name=None, ReadOnly=None, Type=None):
        arguments = {"Name": Name, "ReadOnly": ReadOnly, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.OpenLinks(*args, **arguments)

    def PivotCaches(self):
        return self.workbook.PivotCaches()

    def Post(self, *args, DestName=None):
        arguments = {"DestName": DestName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.Post(*args, **arguments)

    def PrintOut(self, *args, From=None, To=None, Copies=None, Preview=None, ActivePrinter=None, PrintToFile=None, Collate=None, PrToFileName=None, IgnorePrintAreas=None):
        arguments = {"From": From, "To": To, "Copies": Copies, "Preview": Preview, "ActivePrinter": ActivePrinter, "PrintToFile": PrintToFile, "Collate": Collate, "PrToFileName": PrToFileName, "IgnorePrintAreas": IgnorePrintAreas}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.workbook.PrintOut(*args, **arguments)

    def PrintPreview(self, *args, EnableChanges=None):
        arguments = {"EnableChanges": EnableChanges}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.PrintPreview(*args, **arguments)

    def Protect(self, *args, Password=None, Structure=None, Windows=None):
        arguments = {"Password": Password, "Structure": Structure, "Windows": Windows}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.Protect(*args, **arguments)

    def ProtectSharing(self, *args, FileName=None, Password=None, WriteResPassword=None, ReadOnlyRecommended=None, CreateBackup=None, SharingPassword=None, FileFormat=None):
        arguments = {"FileName": FileName, "Password": Password, "WriteResPassword": WriteResPassword, "ReadOnlyRecommended": ReadOnlyRecommended, "CreateBackup": CreateBackup, "SharingPassword": SharingPassword, "FileFormat": FileFormat}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.ProtectSharing(*args, **arguments)

    def PurgeChangeHistoryNow(self, *args, Days=None, SharingPassword=None):
        arguments = {"Days": Days, "SharingPassword": SharingPassword}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.PurgeChangeHistoryNow(*args, **arguments)

    def RefreshAll(self):
        self.workbook.RefreshAll()

    def RejectAllChanges(self, *args, When=None, Who=None, Where=None):
        arguments = {"When": When, "Who": Who, "Where": Where}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.RejectAllChanges(*args, **arguments)

    def ReloadAs(self, *args, Encoding=None):
        arguments = {"Encoding": Encoding}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.ReloadAs(*args, **arguments)

    def RemoveDocumentInformation(self, *args, RemoveDocInfoType=None):
        arguments = {"RemoveDocInfoType": RemoveDocInfoType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.RemoveDocumentInformation(*args, **arguments)

    def RemoveUser(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.RemoveUser(*args, **arguments)

    def Reply(self):
        self.workbook.Reply()

    def ReplyAll(self):
        self.workbook.ReplyAll()

    def ReplyWithChanges(self, *args, ShowMessage=None):
        arguments = {"ShowMessage": ShowMessage}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.ReplyWithChanges(*args, **arguments)

    def ResetColors(self):
        self.workbook.ResetColors()

    def RunAutoMacros(self, *args, Which=None):
        arguments = {"Which": Which}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.RunAutoMacros(*args, **arguments)

    def Save(self):
        self.workbook.Save()

    def SaveAs(self, *args, FileName=None, FileFormat=None, Password=None, WriteResPassword=None, ReadOnlyRecommended=None, CreateBackup=None, AccessMode=None, ConflictResolution=None, AddToMru=None, TextCodepage=None, TextVisualLayout=None, Local=None):
        arguments = {"FileName": FileName, "FileFormat": FileFormat, "Password": Password, "WriteResPassword": WriteResPassword, "ReadOnlyRecommended": ReadOnlyRecommended, "CreateBackup": CreateBackup, "AccessMode": AccessMode, "ConflictResolution": ConflictResolution, "AddToMru": AddToMru, "TextCodepage": TextCodepage, "TextVisualLayout": TextVisualLayout, "Local": Local}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.SaveAs(*args, **arguments)

    def SaveAsXMLData(self, *args, FileName=None, Map=None):
        arguments = {"FileName": FileName, "Map": Map}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.SaveAsXMLData(*args, **arguments)

    def SaveCopyAs(self, *args, FileName=None):
        arguments = {"FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.SaveCopyAs(*args, **arguments)

    def SendFaxOverInternet(self, *args, Recipients=None, Subject=None, ShowMessage=None):
        arguments = {"Recipients": Recipients, "Subject": Subject, "ShowMessage": ShowMessage}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.SendFaxOverInternet(*args, **arguments)

    def SendForReview(self, *args, Recipients=None, Subject=None, ShowMessage=None, IncludeAttachment=None):
        arguments = {"Recipients": Recipients, "Subject": Subject, "ShowMessage": ShowMessage, "IncludeAttachment": IncludeAttachment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.SendForReview(*args, **arguments)

    def SendMail(self, *args, Recipients=None, Subject=None, ReturnReceipt=None):
        arguments = {"Recipients": Recipients, "Subject": Subject, "ReturnReceipt": ReturnReceipt}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.SendMail(*args, **arguments)

    def SendMailer(self, *args, FileFormat=None, Priority=None):
        arguments = {"FileFormat": FileFormat, "Priority": Priority}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.SendMailer(*args, **arguments)

    def SetLinkOnData(self, *args, Name=None, Procedure=None):
        arguments = {"Name": Name, "Procedure": Procedure}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.SetLinkOnData(*args, **arguments)

    def SetPasswordEncryptionOptions(self, *args, PasswordEncryptionProvider=None, PasswordEncryptionAlgorithm=None, PasswordEncryptionKeyLength=None, PasswordEncryptionFileProperties=None):
        arguments = {"PasswordEncryptionProvider": PasswordEncryptionProvider, "PasswordEncryptionAlgorithm": PasswordEncryptionAlgorithm, "PasswordEncryptionKeyLength": PasswordEncryptionKeyLength, "PasswordEncryptionFileProperties": PasswordEncryptionFileProperties}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.SetPasswordEncryptionOptions(*args, **arguments)

    def ToggleFormsDesign(self):
        self.workbook.ToggleFormsDesign()

    def Unprotect(self, *args, Password=None):
        arguments = {"Password": Password}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.Unprotect(*args, **arguments)

    def UnprotectSharing(self, *args, SharingPassword=None):
        arguments = {"SharingPassword": SharingPassword}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.UnprotectSharing(*args, **arguments)

    def UpdateFromFile(self):
        self.workbook.UpdateFromFile()

    def UpdateLink(self, *args, Name=None, Type=None):
        arguments = {"Name": Name, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbook.UpdateLink(*args, **arguments)

    def WebPagePreview(self):
        self.workbook.WebPagePreview()

    def XmlImport(self, *args, Url=None, ImportMap=None, Overwrite=None, Destination=None):
        arguments = {"Url": Url, "ImportMap": ImportMap, "Overwrite": Overwrite, "Destination": Destination}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return XlXmlImportResult(self.workbook.XmlImport(*args, **arguments))

    def XmlImportXml(self, *args, Data=None, ImportMap=None, Overwrite=None, Destination=None):
        arguments = {"Data": Data, "ImportMap": ImportMap, "Overwrite": Overwrite, "Destination": Destination}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return XlXmlImportResult(self.workbook.XmlImportXml(*args, **arguments))





































































































































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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.workbooks.Item):
            return self.workbooks.Item(*args, **arguments)
        else:
            return self.workbooks.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.workbooks.Parent

    def Add(self, *args, Template=None):
        arguments = {"Template": Template}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Workbook(self.workbooks.Add(*args, **arguments))

    def CanCheckOut(self, *args, FileName=None):
        arguments = {"FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.workbooks.CanCheckOut(*args, **arguments)

    def CheckOut(self, *args, FileName=None):
        arguments = {"FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbooks.CheckOut(*args, **arguments)

    def Close(self):
        self.workbooks.Close()

    def Open(self, *args, FileName=None, UpdateLinks=None, ReadOnly=None, Format=None, Password=None, WriteResPassword=None, IgnoreReadOnlyRecommended=None, Origin=None, Delimiter=None, Editable=None, Notify=None, Converter=None, AddToMru=None, Local=None, CorruptLoad=None):
        arguments = {"FileName": FileName, "UpdateLinks": UpdateLinks, "ReadOnly": ReadOnly, "Format": Format, "Password": Password, "WriteResPassword": WriteResPassword, "IgnoreReadOnlyRecommended": IgnoreReadOnlyRecommended, "Origin": Origin, "Delimiter": Delimiter, "Editable": Editable, "Notify": Notify, "Converter": Converter, "AddToMru": AddToMru, "Local": Local, "CorruptLoad": CorruptLoad}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Workbook(self.workbooks.Open(*args, **arguments))

    def OpenDatabase(self, *args, FileName=None, CommandText=None, CommandType=None, BackgroundQuery=None, ImportDataAs=None):
        arguments = {"FileName": FileName, "CommandText": CommandText, "CommandType": CommandType, "BackgroundQuery": BackgroundQuery, "ImportDataAs": ImportDataAs}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Workbook(self.workbooks.OpenDatabase(*args, **arguments))

    def OpenText(self, *args, FileName=None, Origin=None, StartRow=None, DataType=None, TextQualifier=None, ConsecutiveDelimiter=None, Tab=None, Semicolon=None, Comma=None, Space=None, Other=None, OtherChar=None, FieldInfo=None, TextVisualLayout=None, DecimalSeparator=None, ThousandsSeparator=None, TrailingMinusNumbers=None, Local=None):
        arguments = {"FileName": FileName, "Origin": Origin, "StartRow": StartRow, "DataType": DataType, "TextQualifier": TextQualifier, "ConsecutiveDelimiter": ConsecutiveDelimiter, "Tab": Tab, "Semicolon": Semicolon, "Comma": Comma, "Space": Space, "Other": Other, "OtherChar": OtherChar, "FieldInfo": FieldInfo, "TextVisualLayout": TextVisualLayout, "DecimalSeparator": DecimalSeparator, "ThousandsSeparator": ThousandsSeparator, "TrailingMinusNumbers": TrailingMinusNumbers, "Local": Local}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.workbooks.OpenText(*args, **arguments)

    def OpenXML(self, *args, FileName=None, Stylesheets=None, LoadOption=None):
        arguments = {"FileName": FileName, "Stylesheets": Stylesheets, "LoadOption": LoadOption}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Workbook(self.workbooks.OpenXML(*args, **arguments))



















































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

    def Cells(self, *args, RowIndex=None, ColumnIndex=None):
        arguments = {"RowIndex": RowIndex, "ColumnIndex": ColumnIndex}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.worksheet.Cells):
            return Range(self.worksheet.Cells(*args, **arguments))
        else:
            return Range(self.worksheet.GetCells(*args, **arguments))

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

    def Range(self, *args, Cell1=None, Cell2=None):
        arguments = {"Cell1": Cell1, "Cell2": Cell2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.worksheet.Range):
            return Range(self.worksheet.Range(*args, **arguments))
        else:
            return Range(self.worksheet.GetRange(*args, **arguments))

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

    def ChartObjects(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheet.ChartObjects(*args, **arguments)

    def CheckSpelling(self, *args, CustomDictionary=None, IgnoreUppercase=None, AlwaysSuggest=None, SpellLang=None):
        arguments = {"CustomDictionary": CustomDictionary, "IgnoreUppercase": IgnoreUppercase, "AlwaysSuggest": AlwaysSuggest, "SpellLang": SpellLang}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.worksheet.CheckSpelling(*args, **arguments)

    def CircleInvalid(self):
        self.worksheet.CircleInvalid()

    def ClearArrows(self):
        self.worksheet.ClearArrows()

    def ClearCircles(self):
        self.worksheet.ClearCircles()

    def Copy(self, *args, Before=None, After=None):
        arguments = {"Before": Before, "After": After}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.worksheet.Copy(*args, **arguments)

    def Delete(self):
        return self.worksheet.Delete()

    def Evaluate(self, *args, Name=None):
        arguments = {"Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheet.Evaluate(*args, **arguments)

    def ExportAsFixedFormat(self, *args, Type=None, FileName=None, Quality=None, IncludeDocProperties=None, IgnorePrintAreas=None, From=None, To=None, OpenAfterPublish=None, FixedFormatExtClassPtr=None):
        arguments = {"Type": Type, "FileName": FileName, "Quality": Quality, "IncludeDocProperties": IncludeDocProperties, "IgnorePrintAreas": IgnorePrintAreas, "From": From, "To": To, "OpenAfterPublish": OpenAfterPublish, "FixedFormatExtClassPtr": FixedFormatExtClassPtr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.worksheet.ExportAsFixedFormat(*args, **arguments)

    def Move(self, *args, Before=None, After=None):
        arguments = {"Before": Before, "After": After}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.worksheet.Move(*args, **arguments)

    def OLEObjects(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheet.OLEObjects(*args, **arguments)

    def Paste(self, *args, Destination=None, Link=None):
        arguments = {"Destination": Destination, "Link": Link}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.worksheet.Paste(*args, **arguments)

    def PasteSpecial(self, *args, Format=None, Link=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, NoHTMLFormatting=None):
        arguments = {"Format": Format, "Link": Link, "DisplayAsIcon": DisplayAsIcon, "IconFileName": IconFileName, "IconIndex": IconIndex, "IconLabel": IconLabel, "NoHTMLFormatting": NoHTMLFormatting}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.worksheet.PasteSpecial(*args, **arguments)

    def PivotTables(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return PivotTable(self.worksheet.PivotTables(*args, **arguments))

    def PivotTableWizard(self, *args, SourceType=None, SourceData=None, TableDestination=None, TableName=None, RowGrand=None, ColumnGrand=None, SaveData=None, HasAutoFormat=None, AutoPage=None, Reserved=None, BackgroundQuery=None, OptimizeCache=None, PageFieldOrder=None, PageFieldWrapCount=None, ReadData=None, Connection=None):
        arguments = {"SourceType": SourceType, "SourceData": SourceData, "TableDestination": TableDestination, "TableName": TableName, "RowGrand": RowGrand, "ColumnGrand": ColumnGrand, "SaveData": SaveData, "HasAutoFormat": HasAutoFormat, "AutoPage": AutoPage, "Reserved": Reserved, "BackgroundQuery": BackgroundQuery, "OptimizeCache": OptimizeCache, "PageFieldOrder": PageFieldOrder, "PageFieldWrapCount": PageFieldWrapCount, "ReadData": ReadData, "Connection": Connection}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return PivotTable(self.worksheet.PivotTableWizard(*args, **arguments))

    def PrintOut(self, *args, From=None, To=None, Copies=None, Preview=None, ActivePrinter=None, PrintToFile=None, Collate=None, PrToFileName=None, IgnorePrintAreas=None):
        arguments = {"From": From, "To": To, "Copies": Copies, "Preview": Preview, "ActivePrinter": ActivePrinter, "PrintToFile": PrintToFile, "Collate": Collate, "PrToFileName": PrToFileName, "IgnorePrintAreas": IgnorePrintAreas}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheet.PrintOut(*args, **arguments)

    def PrintPreview(self, *args, EnableChanges=None):
        arguments = {"EnableChanges": EnableChanges}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.worksheet.PrintPreview(*args, **arguments)

    def Protect(self, *args, Password=None, DrawingObjects=None, Contents=None, Scenarios=None, UserInterfaceOnly=None, AllowFormattingCells=None, AllowFormattingColumns=None, AllowFormattingRows=None, AllowInsertingColumns=None, AllowInsertingRows=None, AllowInsertingHyperlinks=None, AllowDeletingColumns=None, AllowDeletingRows=None, AllowSorting=None, AllowFiltering=None, AllowUsingPivotTables=None):
        arguments = {"Password": Password, "DrawingObjects": DrawingObjects, "Contents": Contents, "Scenarios": Scenarios, "UserInterfaceOnly": UserInterfaceOnly, "AllowFormattingCells": AllowFormattingCells, "AllowFormattingColumns": AllowFormattingColumns, "AllowFormattingRows": AllowFormattingRows, "AllowInsertingColumns": AllowInsertingColumns, "AllowInsertingRows": AllowInsertingRows, "AllowInsertingHyperlinks": AllowInsertingHyperlinks, "AllowDeletingColumns": AllowDeletingColumns, "AllowDeletingRows": AllowDeletingRows, "AllowSorting": AllowSorting, "AllowFiltering": AllowFiltering, "AllowUsingPivotTables": AllowUsingPivotTables}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.worksheet.Protect(*args, **arguments)

    def ResetAllPageBreaks(self):
        self.worksheet.ResetAllPageBreaks()

    def SaveAs(self, *args, FileName=None, FileFormat=None, Password=None, WriteResPassword=None, ReadOnlyRecommended=None, CreateBackup=None, AddToMru=None, TextCodepage=None, TextVisualLayout=None, Local=None):
        arguments = {"FileName": FileName, "FileFormat": FileFormat, "Password": Password, "WriteResPassword": WriteResPassword, "ReadOnlyRecommended": ReadOnlyRecommended, "CreateBackup": CreateBackup, "AddToMru": AddToMru, "TextCodepage": TextCodepage, "TextVisualLayout": TextVisualLayout, "Local": Local}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.worksheet.SaveAs(*args, **arguments)

    def Scenarios(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheet.Scenarios(*args, **arguments)

    def Select(self, *args, Replace=None):
        arguments = {"Replace": Replace}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.worksheet.Select(*args, **arguments)

    def SetBackgroundPicture(self, *args, FileName=None):
        arguments = {"FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.worksheet.SetBackgroundPicture(*args, **arguments)

    def ShowAllData(self):
        self.worksheet.ShowAllData()

    def ShowDataForm(self):
        self.worksheet.ShowDataForm()

    def Unprotect(self, *args, Password=None):
        arguments = {"Password": Password}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.worksheet.Unprotect(*args, **arguments)

    def XmlDataQuery(self, *args, XPath=None, SelectionNamespaces=None, Map=None):
        arguments = {"XPath": XPath, "SelectionNamespaces": SelectionNamespaces, "Map": Map}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheet.XmlDataQuery(*args, **arguments)

    def XmlMapQuery(self, *args, XPath=None, SelectionNamespaces=None, Map=None):
        arguments = {"XPath": XPath, "SelectionNamespaces": SelectionNamespaces, "Map": Map}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheet.XmlMapQuery(*args, **arguments)

























































































































































































































































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

    def AccrInt(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.AccrInt(*args, **arguments)

    def AccrIntM(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.AccrIntM(*args, **arguments)

    def Acos(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Acos(*args, **arguments)

    def Acosh(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Acosh(*args, **arguments)

    def Aggregate(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Aggregate(*args, **arguments)

    def AmorDegrc(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.AmorDegrc(*args, **arguments)

    def AmorLinc(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.AmorLinc(*args, **arguments)

    def And(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.And(*args, **arguments)

    def Asc(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Asc(*args, **arguments)

    def Asin(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Asin(*args, **arguments)

    def Asinh(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Asinh(*args, **arguments)

    def Atan2(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Atan2(*args, **arguments)

    def Atanh(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Atanh(*args, **arguments)

    def AveDev(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.AveDev(*args, **arguments)

    def Average(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Average(*args, **arguments)

    def AverageIf(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.AverageIf(*args, **arguments)

    def AverageIfs(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.AverageIfs(*args, **arguments)

    def BahtText(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.BahtText(*args, **arguments)

    def BesselI(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.BesselI(*args, **arguments)

    def BesselJ(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.BesselJ(*args, **arguments)

    def BesselK(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.BesselK(*args, **arguments)

    def BesselY(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.BesselY(*args, **arguments)

    def BetaDist(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.BetaDist(*args, **arguments)

    def BetaInv(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.BetaInv(*args, **arguments)

    def Beta_Dist(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Beta_Dist(*args, **arguments)

    def Beta_Inv(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Beta_Inv(*args, **arguments)

    def Bin2Dec(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Bin2Dec(*args, **arguments)

    def Bin2Hex(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Bin2Hex(*args, **arguments)

    def Bin2Oct(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Bin2Oct(*args, **arguments)

    def BinomDist(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.BinomDist(*args, **arguments)

    def Binom_Dist(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Binom_Dist(*args, **arguments)

    def Binom_Inv(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Binom_Inv(*args, **arguments)

    def Ceiling(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Ceiling(*args, **arguments)

    def Ceiling_Precise(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Ceiling_Precise(*args, **arguments)

    def ChiDist(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ChiDist(*args, **arguments)

    def ChiInv(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ChiInv(*args, **arguments)

    def ChiSq_Dist(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ChiSq_Dist(*args, **arguments)

    def ChiSq_Dist_RT(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ChiSq_Dist_RT(*args, **arguments)

    def ChiSq_Inv(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ChiSq_Inv(*args, **arguments)

    def ChiSq_Inv_RT(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ChiSq_Inv_RT(*args, **arguments)

    def ChiSq_Test(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ChiSq_Test(*args, **arguments)

    def ChiTest(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ChiTest(*args, **arguments)

    def Choose(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Choose(*args, **arguments)

    def Clean(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Clean(*args, **arguments)

    def Combin(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Combin(*args, **arguments)

    def Complex(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Complex(*args, **arguments)

    def Confidence(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Confidence(*args, **arguments)

    def Confidence_Norm(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Confidence_Norm(*args, **arguments)

    def Confidence_T(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Confidence_T(*args, **arguments)

    def Convert(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Convert(*args, **arguments)

    def Correl(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Correl(*args, **arguments)

    def Cosh(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Cosh(*args, **arguments)

    def Count(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Count(*args, **arguments)

    def CountA(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.CountA(*args, **arguments)

    def CountBlank(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.CountBlank(*args, **arguments)

    def CountIf(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.CountIf(*args, **arguments)

    def CountIfs(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.CountIfs(*args, **arguments)

    def CoupDayBs(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.CoupDayBs(*args, **arguments)

    def CoupDays(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.CoupDays(*args, **arguments)

    def CoupDaysNc(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.CoupDaysNc(*args, **arguments)

    def CoupNcd(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.CoupNcd(*args, **arguments)

    def CoupNum(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.CoupNum(*args, **arguments)

    def Covar(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Covar(*args, **arguments)

    def Covariance_P(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Covariance_P(*args, **arguments)

    def Covariance_S(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Covariance_S(*args, **arguments)

    def CritBinom(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.CritBinom(*args, **arguments)

    def CumIPmt(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.CumIPmt(*args, **arguments)

    def CumPrinc(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.CumPrinc(*args, **arguments)

    def DAverage(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.DAverage(*args, **arguments)

    def Days360(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Days360(*args, **arguments)

    def Db(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Db(*args, **arguments)

    def Dbcs(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Dbcs(*args, **arguments)

    def DCount(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.DCount(*args, **arguments)

    def DCountA(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.DCountA(*args, **arguments)

    def Ddb(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Ddb(*args, **arguments)

    def Dec2Bin(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Dec2Bin(*args, **arguments)

    def Dec2Hex(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Dec2Hex(*args, **arguments)

    def Dec2Oct(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Dec2Oct(*args, **arguments)

    def Degrees(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Degrees(*args, **arguments)

    def Delta(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Delta(*args, **arguments)

    def DevSq(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.DevSq(*args, **arguments)

    def DGet(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.DGet(*args, **arguments)

    def Disc(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Disc(*args, **arguments)

    def DMax(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.DMax(*args, **arguments)

    def DMin(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.DMin(*args, **arguments)

    def Dollar(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Dollar(*args, **arguments)

    def DollarDe(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.DollarDe(*args, **arguments)

    def DollarFr(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.DollarFr(*args, **arguments)

    def DProduct(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.DProduct(*args, **arguments)

    def DStDev(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.DStDev(*args, **arguments)

    def DStDevP(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.DStDevP(*args, **arguments)

    def DSum(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.DSum(*args, **arguments)

    def Duration(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Duration(*args, **arguments)

    def DVar(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.DVar(*args, **arguments)

    def DVarP(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.DVarP(*args, **arguments)

    def EDate(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.EDate(*args, **arguments)

    def Effect(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Effect(*args, **arguments)

    def EoMonth(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.EoMonth(*args, **arguments)

    def Erf(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Erf(*args, **arguments)

    def ErfC(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ErfC(*args, **arguments)

    def ErfC_Precise(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ErfC_Precise(*args, **arguments)

    def Erf_Precise(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Erf_Precise(*args, **arguments)

    def Even(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Even(*args, **arguments)

    def ExponDist(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ExponDist(*args, **arguments)

    def Expon_Dist(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Expon_Dist(*args, **arguments)

    def Fact(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Fact(*args, **arguments)

    def FactDouble(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.FactDouble(*args, **arguments)

    def FDist(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.FDist(*args, **arguments)

    def Find(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Find(*args, **arguments)

    def FindB(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.FindB(*args, **arguments)

    def FInv(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.FInv(*args, **arguments)

    def Fisher(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Fisher(*args, **arguments)

    def FisherInv(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.FisherInv(*args, **arguments)

    def Fixed(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Fixed(*args, **arguments)

    def Floor(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Floor(*args, **arguments)

    def Floor_Precise(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Floor_Precise(*args, **arguments)

    def Forecast(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Forecast(*args, **arguments)

    def Frequency(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Frequency(*args, **arguments)

    def FTest(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.FTest(*args, **arguments)

    def Fv(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Fv(*args, **arguments)

    def FVSchedule(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.FVSchedule(*args, **arguments)

    def F_Dist(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.F_Dist(*args, **arguments)

    def F_Dist_RT(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.F_Dist_RT(*args, **arguments)

    def F_Inv(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.F_Inv(*args, **arguments)

    def F_Inv_RT(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.F_Inv_RT(*args, **arguments)

    def F_Test(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.F_Test(*args, **arguments)

    def GammaDist(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.GammaDist(*args, **arguments)

    def GammaInv(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.GammaInv(*args, **arguments)

    def GammaLn(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.GammaLn(*args, **arguments)

    def GammaLn_Precise(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.GammaLn_Precise(*args, **arguments)

    def Gamma_Dist(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Gamma_Dist(*args, **arguments)

    def Gamma_Inv(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Gamma_Inv(*args, **arguments)

    def Gcd(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Gcd(*args, **arguments)

    def GeoMean(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.GeoMean(*args, **arguments)

    def GeStep(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.GeStep(*args, **arguments)

    def Growth(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Growth(*args, **arguments)

    def HarMean(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.HarMean(*args, **arguments)

    def Hex2Bin(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Hex2Bin(*args, **arguments)

    def Hex2Dec(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Hex2Dec(*args, **arguments)

    def Hex2Oct(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Hex2Oct(*args, **arguments)

    def HLookup(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.HLookup(*args, **arguments)

    def HypGeomDist(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.HypGeomDist(*args, **arguments)

    def HypGeom_Dist(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.HypGeom_Dist(*args, **arguments)

    def IfError(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.IfError(*args, **arguments)

    def ImAbs(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ImAbs(*args, **arguments)

    def Imaginary(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Imaginary(*args, **arguments)

    def ImArgument(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ImArgument(*args, **arguments)

    def ImConjugate(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ImConjugate(*args, **arguments)

    def ImCos(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ImCos(*args, **arguments)

    def ImDiv(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ImDiv(*args, **arguments)

    def ImExp(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ImExp(*args, **arguments)

    def ImLn(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ImLn(*args, **arguments)

    def ImLog10(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ImLog10(*args, **arguments)

    def ImLog2(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ImLog2(*args, **arguments)

    def ImPower(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ImPower(*args, **arguments)

    def ImProduct(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ImProduct(*args, **arguments)

    def ImReal(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ImReal(*args, **arguments)

    def ImSin(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ImSin(*args, **arguments)

    def ImSqrt(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ImSqrt(*args, **arguments)

    def ImSub(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ImSub(*args, **arguments)

    def ImSum(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ImSum(*args, **arguments)

    def Intercept(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Intercept(*args, **arguments)

    def IntRate(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.IntRate(*args, **arguments)

    def Ipmt(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Ipmt(*args, **arguments)

    def Irr(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Irr(*args, **arguments)

    def IsErr(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.IsErr(*args, **arguments)

    def IsError(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.IsError(*args, **arguments)

    def IsEven(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.IsEven(*args, **arguments)

    def IsLogical(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.IsLogical(*args, **arguments)

    def IsNA(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.IsNA(*args, **arguments)

    def IsNonText(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.IsNonText(*args, **arguments)

    def IsNumber(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.IsNumber(*args, **arguments)

    def IsOdd(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.IsOdd(*args, **arguments)

    def ISO_Ceiling(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ISO_Ceiling(*args, **arguments)

    def Ispmt(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Ispmt(*args, **arguments)

    def IsText(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.IsText(*args, **arguments)

    def Kurt(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Kurt(*args, **arguments)

    def Large(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Large(*args, **arguments)

    def Lcm(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Lcm(*args, **arguments)

    def LinEst(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.LinEst(*args, **arguments)

    def Ln(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Ln(*args, **arguments)

    def Log(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Log(*args, **arguments)

    def Log10(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Log10(*args, **arguments)

    def LogEst(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.LogEst(*args, **arguments)

    def LogInv(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.LogInv(*args, **arguments)

    def LogNormDist(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.LogNormDist(*args, **arguments)

    def LogNorm_Dist(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.LogNorm_Dist(*args, **arguments)

    def LogNorm_Inv(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.LogNorm_Inv(*args, **arguments)

    def Lookup(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Lookup(*args, **arguments)

    def Match(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Match(*args, **arguments)

    def Max(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Max(*args, **arguments)

    def MDeterm(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.MDeterm(*args, **arguments)

    def MDuration(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.MDuration(*args, **arguments)

    def Median(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Median(*args, **arguments)

    def Min(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Min(*args, **arguments)

    def MInverse(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.MInverse(*args, **arguments)

    def MIrr(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.MIrr(*args, **arguments)

    def MMult(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.MMult(*args, **arguments)

    def Mode(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Mode(*args, **arguments)

    def Mode_Mult(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Mode_Mult(*args, **arguments)

    def Mode_Sngl(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Mode_Sngl(*args, **arguments)

    def MRound(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.MRound(*args, **arguments)

    def MultiNomial(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.MultiNomial(*args, **arguments)

    def NegBinomDist(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.NegBinomDist(*args, **arguments)

    def NegBinom_Dist(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.NegBinom_Dist(*args, **arguments)

    def NetworkDays(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.NetworkDays(*args, **arguments)

    def NetworkDays_Intl(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.NetworkDays_Intl(*args, **arguments)

    def Nominal(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Nominal(*args, **arguments)

    def NormDist(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.NormDist(*args, **arguments)

    def NormInv(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.NormInv(*args, **arguments)

    def NormSDist(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.NormSDist(*args, **arguments)

    def NormSInv(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.NormSInv(*args, **arguments)

    def Norm_Dist(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Norm_Dist(*args, **arguments)

    def Norm_Inv(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Norm_Inv(*args, **arguments)

    def Norm_S_Dist(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Norm_S_Dist(*args, **arguments)

    def Norm_S_Inv(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Norm_S_Inv(*args, **arguments)

    def NPer(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.NPer(*args, **arguments)

    def Npv(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Npv(*args, **arguments)

    def Oct2Bin(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Oct2Bin(*args, **arguments)

    def Oct2Dec(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Oct2Dec(*args, **arguments)

    def Oct2Hex(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Oct2Hex(*args, **arguments)

    def Odd(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Odd(*args, **arguments)

    def OddFPrice(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.OddFPrice(*args, **arguments)

    def OddFYield(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.OddFYield(*args, **arguments)

    def OddLPrice(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.OddLPrice(*args, **arguments)

    def OddLYield(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.OddLYield(*args, **arguments)

    def Or(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Or(*args, **arguments)

    def Pearson(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Pearson(*args, **arguments)

    def Percentile(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Percentile(*args, **arguments)

    def Percentile_Exc(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Percentile_Exc(*args, **arguments)

    def Percentile_Inc(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Percentile_Inc(*args, **arguments)

    def PercentRank(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.PercentRank(*args, **arguments)

    def PercentRank_Exc(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.PercentRank_Exc(*args, **arguments)

    def PercentRank_Inc(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.PercentRank_Inc(*args, **arguments)

    def Permut(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Permut(*args, **arguments)

    def Phonetic(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Phonetic(*args, **arguments)

    def Pi(self):
        return self.worksheetfunction.Pi()

    def Pmt(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Pmt(*args, **arguments)

    def Poisson(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Poisson(*args, **arguments)

    def Poisson_Dist(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Poisson_Dist(*args, **arguments)

    def Power(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Power(*args, **arguments)

    def Ppmt(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Ppmt(*args, **arguments)

    def Price(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Price(*args, **arguments)

    def PriceDisc(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.PriceDisc(*args, **arguments)

    def PriceMat(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.PriceMat(*args, **arguments)

    def Prob(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Prob(*args, **arguments)

    def Product(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Product(*args, **arguments)

    def Proper(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Proper(*args, **arguments)

    def Pv(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Pv(*args, **arguments)

    def Quartile(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Quartile(*args, **arguments)

    def Quartile_Exc(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Quartile_Exc(*args, **arguments)

    def Quartile_Inc(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Quartile_Inc(*args, **arguments)

    def Quotient(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Quotient(*args, **arguments)

    def Radians(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Radians(*args, **arguments)

    def RandBetween(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.RandBetween(*args, **arguments)

    def Rank(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Rank(*args, **arguments)

    def Rank_Avg(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Rank_Avg(*args, **arguments)

    def Rank_Eq(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Rank_Eq(*args, **arguments)

    def Rate(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Rate(*args, **arguments)

    def Received(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Received(*args, **arguments)

    def Replace(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Replace(*args, **arguments)

    def ReplaceB(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ReplaceB(*args, **arguments)

    def Rept(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Rept(*args, **arguments)

    def Roman(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Roman(*args, **arguments)

    def Round(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Round(*args, **arguments)

    def RoundDown(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.RoundDown(*args, **arguments)

    def RoundUp(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.RoundUp(*args, **arguments)

    def RSq(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.RSq(*args, **arguments)

    def RTD(self, *args, progID=None, server=None, topic1=None, topic2=None, topic3=None, topic4=None, topic5=None, topic6=None, topic7=None, topic8=None, topic9=None, topic10=None, topic11=None, topic12=None, topic13=None, topic14=None, topic15=None, topic16=None, topic17=None, topic18=None, topic19=None, topic20=None, topic21=None, topic22=None, topic23=None, topic24=None, topic25=None, topic26=None, topic27=None, topic28=None):
        arguments = {"progID": progID, "server": server, "topic1": topic1, "topic2": topic2, "topic3": topic3, "topic4": topic4, "topic5": topic5, "topic6": topic6, "topic7": topic7, "topic8": topic8, "topic9": topic9, "topic10": topic10, "topic11": topic11, "topic12": topic12, "topic13": topic13, "topic14": topic14, "topic15": topic15, "topic16": topic16, "topic17": topic17, "topic18": topic18, "topic19": topic19, "topic20": topic20, "topic21": topic21, "topic22": topic22, "topic23": topic23, "topic24": topic24, "topic25": topic25, "topic26": topic26, "topic27": topic27, "topic28": topic28}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.RTD(*args, **arguments)

    def Search(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Search(*args, **arguments)

    def SearchB(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.SearchB(*args, **arguments)

    def SeriesSum(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.SeriesSum(*args, **arguments)

    def Sinh(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Sinh(*args, **arguments)

    def Skew(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Skew(*args, **arguments)

    def Sln(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Sln(*args, **arguments)

    def Slope(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Slope(*args, **arguments)

    def Small(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Small(*args, **arguments)

    def SqrtPi(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.SqrtPi(*args, **arguments)

    def Standardize(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Standardize(*args, **arguments)

    def StDev(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.StDev(*args, **arguments)

    def StDevP(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.StDevP(*args, **arguments)

    def StDev_P(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.StDev_P(*args, **arguments)

    def StDev_S(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.StDev_S(*args, **arguments)

    def StEyx(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.StEyx(*args, **arguments)

    def Substitute(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Substitute(*args, **arguments)

    def Subtotal(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Subtotal(*args, **arguments)

    def Sum(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Sum(*args, **arguments)

    def SumIf(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.SumIf(*args, **arguments)

    def SumIfs(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.SumIfs(*args, **arguments)

    def SumProduct(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.SumProduct(*args, **arguments)

    def SumSq(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.SumSq(*args, **arguments)

    def SumX2MY2(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.SumX2MY2(*args, **arguments)

    def SumX2PY2(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.SumX2PY2(*args, **arguments)

    def SumXMY2(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.SumXMY2(*args, **arguments)

    def Syd(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Syd(*args, **arguments)

    def Tanh(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Tanh(*args, **arguments)

    def TBillEq(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.TBillEq(*args, **arguments)

    def TBillPrice(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.TBillPrice(*args, **arguments)

    def TBillYield(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.TBillYield(*args, **arguments)

    def TDist(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.TDist(*args, **arguments)

    def Text(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Text(*args, **arguments)

    def TInv(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.TInv(*args, **arguments)

    def Transpose(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Transpose(*args, **arguments)

    def Trend(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Trend(*args, **arguments)

    def Trim(self, *args, Arg1=None):
        arguments = {"Arg1": Arg1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Trim(*args, **arguments)

    def TrimMean(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.TrimMean(*args, **arguments)

    def TTest(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.TTest(*args, **arguments)

    def T_Dist(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.T_Dist(*args, **arguments)

    def T_Dist_2T(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.T_Dist_2T(*args, **arguments)

    def T_Dist_RT(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.T_Dist_RT(*args, **arguments)

    def T_Inv(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.T_Inv(*args, **arguments)

    def T_Inv_2T(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.T_Inv_2T(*args, **arguments)

    def T_Test(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.T_Test(*args, **arguments)

    def USDollar(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.USDollar(*args, **arguments)

    def Var(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Var(*args, **arguments)

    def VarP(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.VarP(*args, **arguments)

    def Var_P(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Var_P(*args, **arguments)

    def Var_S(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Var_S(*args, **arguments)

    def Vdb(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Vdb(*args, **arguments)

    def VLookup(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.VLookup(*args, **arguments)

    def Weekday(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Weekday(*args, **arguments)

    def WeekNum(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.WeekNum(*args, **arguments)

    def Weibull(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Weibull(*args, **arguments)

    def Weibull_Dist(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Weibull_Dist(*args, **arguments)

    def WorkDay(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.WorkDay(*args, **arguments)

    def WorkDay_Intl(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.WorkDay_Intl(*args, **arguments)

    def Xirr(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Xirr(*args, **arguments)

    def Xnpv(self, *args, Arg1=None, Arg2=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Xnpv(*args, **arguments)

    def YearFrac(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.YearFrac(*args, **arguments)

    def YieldDisc(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.YieldDisc(*args, **arguments)

    def YieldMat(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.YieldMat(*args, **arguments)

    def ZTest(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.ZTest(*args, **arguments)

    def Z_Test(self, *args, Arg1=None, Arg2=None, Arg3=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheetfunction.Z_Test(*args, **arguments)


























































































































































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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.worksheets.Item):
            return self.worksheets.Item(*args, **arguments)
        else:
            return self.worksheets.GetItem(*args, **arguments)

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

    def Add(self, *args, Before=None, After=None, Count=None, Type=None):
        arguments = {"Before": Before, "After": After, "Count": Count, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Worksheet(self.worksheets.Add(*args, **arguments))

    def Copy(self, *args, Before=None, After=None):
        arguments = {"Before": Before, "After": After}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.worksheets.Copy(*args, **arguments)

    def Delete(self):
        self.worksheets.Delete()

    def FillAcrossSheets(self, *args, Range=None, Type=None):
        arguments = {"Range": Range, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.worksheets.FillAcrossSheets(*args, **arguments)

    def Move(self, *args, Before=None, After=None):
        arguments = {"Before": Before, "After": After}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.worksheets.Move(*args, **arguments)

    def PrintOut(self, *args, From=None, To=None, Copies=None, Preview=None, ActivePrinter=None, PrintToFile=None, Collate=None, PrToFileName=None, IgnorePrintAreas=None):
        arguments = {"From": From, "To": To, "Copies": Copies, "Preview": Preview, "ActivePrinter": ActivePrinter, "PrintToFile": PrintToFile, "Collate": Collate, "PrToFileName": PrToFileName, "IgnorePrintAreas": IgnorePrintAreas}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.worksheets.PrintOut(*args, **arguments)

    def PrintPreview(self, *args, EnableChanges=None):
        arguments = {"EnableChanges": EnableChanges}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.worksheets.PrintPreview(*args, **arguments)

    def Select(self, *args, Replace=None):
        arguments = {"Replace": Replace}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.worksheets.Select(*args, **arguments)
















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

    def LoadSettings(self, *args, Url=None):
        arguments = {"Url": Url}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.xmldatabinding.LoadSettings(*args, **arguments)

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

    def Export(self, *args, Url=None, Overwrite=None):
        arguments = {"Url": Url, "Overwrite": Overwrite}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return XlXmlExportResult(self.xmlmap.Export(*args, **arguments))

    def ExportXml(self, *args, Data=None):
        arguments = {"Data": Data}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return XlXmlExportResult(self.xmlmap.ExportXml(*args, **arguments))

    def Import(self, *args, Url=None, Overwrite=None):
        arguments = {"Url": Url, "Overwrite": Overwrite}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return XlXmlImportResult(self.xmlmap.Import(*args, **arguments))

    def ImportXml(self, *args, XmlData=None, Overwrite=None):
        arguments = {"XmlData": XmlData, "Overwrite": Overwrite}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return XlXmlImportResult(self.xmlmap.ImportXml(*args, **arguments))

















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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.xmlmaps.Item):
            return self.xmlmaps.Item(*args, **arguments)
        else:
            return self.xmlmaps.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.xmlmaps.Parent

    def Add(self, *args, Schema=None, RootElementName=None):
        arguments = {"Schema": Schema, "RootElementName": RootElementName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return XmlMap(self.xmlmaps.Add(*args, **arguments))





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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.xmlnamespaces.Item):
            return self.xmlnamespaces.Item(*args, **arguments)
        else:
            return self.xmlnamespaces.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.xmlnamespaces.Parent

    @property
    def Value(self):
        return self.xmlnamespaces.Value

    def InstallManifest(self, *args, Path=None, InstallForAllUsers=None):
        arguments = {"Path": Path, "InstallForAllUsers": InstallForAllUsers}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.xmlnamespaces.InstallManifest(*args, **arguments)






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

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.xmlschemas.Item):
            return self.xmlschemas.Item(*args, **arguments)
        else:
            return self.xmlschemas.GetItem(*args, **arguments)

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

    def SetValue(self, *args, Map=None, XPath=None, SelectionNamespace=None, Repeating=None):
        arguments = {"Map": Map, "XPath": XPath, "SelectionNamespace": SelectionNamespace, "Repeating": Repeating}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.xpath.SetValue(*args, **arguments)






