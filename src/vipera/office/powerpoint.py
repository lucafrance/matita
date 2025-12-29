import win32com.client

class ActionSetting:

    def __init__(self, actionsetting=None):
        self.actionsetting = actionsetting

    @property
    def Action(self):
        return self.actionsetting.Action

    @Action.setter
    def Action(self, value):
        self.actionsetting.Action = value

    @property
    def ActionVerb(self):
        return self.actionsetting.ActionVerb

    @ActionVerb.setter
    def ActionVerb(self, value):
        self.actionsetting.ActionVerb = value

    @property
    def AnimateAction(self):
        return self.actionsetting.AnimateAction

    @property
    def Application(self):
        return Application(self.actionsetting.Application)

    @property
    def Hyperlink(self):
        return Hyperlink(self.actionsetting.Hyperlink)

    @property
    def Parent(self):
        return self.actionsetting.Parent

    @property
    def Run(self):
        return self.actionsetting.Run

    @Run.setter
    def Run(self, value):
        self.actionsetting.Run = value

    @property
    def ShowAndReturn(self):
        return self.actionsetting.ShowAndReturn

    @property
    def SlideShowName(self):
        return self.actionsetting.SlideShowName

    @SlideShowName.setter
    def SlideShowName(self, value):
        self.actionsetting.SlideShowName = value

    @property
    def SoundEffect(self):
        return SoundEffect(self.actionsetting.SoundEffect)

class ActionSettings:

    def __init__(self, actionsettings=None):
        self.actionsettings = actionsettings

    def __call__(self, item):
        return ActionSetting(self.actionsettings(item))

    @property
    def Application(self):
        return Application(self.actionsettings.Application)

    @property
    def Count(self):
        return self.actionsettings.Count

    @property
    def Parent(self):
        return self.actionsettings.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.actionsettings.Item(*args, **arguments)

class AddIn:

    def __init__(self, addin=None):
        self.addin = addin

    @property
    def Application(self):
        return Application(self.addin.Application)

    @property
    def AutoLoad(self):
        return self.addin.AutoLoad

    @property
    def FullName(self):
        return self.addin.FullName

    @property
    def Loaded(self):
        return self.addin.Loaded

    @property
    def Name(self):
        return self.addin.Name

    @property
    def Parent(self):
        return self.addin.Parent

    @property
    def Path(self):
        return AddIn(self.addin.Path)

    @property
    def Registered(self):
        return self.addin.Registered

class AddIns:

    def __init__(self, addins=None):
        self.addins = addins

    def __call__(self, item):
        return AddIn(self.addins(item))

    @property
    def Application(self):
        return Application(self.addins.Application)

    @property
    def Count(self):
        return self.addins.Count

    @property
    def Parent(self):
        return self.addins.Parent

    def Add(self, *args, FileName=None):
        arguments = {"FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return AddIn(self.addins.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.addins.Item(*args, **arguments)

    def Remove(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.addins.Remove(*args, **arguments)

class Adjustments:

    def __init__(self, adjustments=None):
        self.adjustments = adjustments

    @property
    def Application(self):
        return Application(self.adjustments.Application)

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

class AnimationBehavior:

    def __init__(self, animationbehavior=None):
        self.animationbehavior = animationbehavior

    @property
    def Accumulate(self):
        return self.animationbehavior.Accumulate

    @property
    def Additive(self):
        return self.animationbehavior.Additive

    @property
    def Application(self):
        return Application(self.animationbehavior.Application)

    @property
    def ColorEffect(self):
        return ColorEffect(self.animationbehavior.ColorEffect)

    @property
    def CommandEffect(self):
        return CommandEffect(self.animationbehavior.CommandEffect)

    @property
    def FilterEffect(self):
        return FilterEffect(self.animationbehavior.FilterEffect)

    @property
    def MotionEffect(self):
        return MotionEffect(self.animationbehavior.MotionEffect)

    @property
    def Parent(self):
        return self.animationbehavior.Parent

    @property
    def PropertyEffect(self):
        return PropertyEffect(self.animationbehavior.PropertyEffect)

    @property
    def RotationEffect(self):
        return RotationEffect(self.animationbehavior.RotationEffect)

    @property
    def ScaleEffect(self):
        return ScaleEffect(self.animationbehavior.ScaleEffect)

    @property
    def SetEffect(self):
        return SetEffect(self.animationbehavior.SetEffect)

    @property
    def Timing(self):
        return Timing(self.animationbehavior.Timing)

    @property
    def Type(self):
        return self.animationbehavior.Type

    def Delete(self):
        self.animationbehavior.Delete()

class AnimationBehaviors:

    def __init__(self, animationbehaviors=None):
        self.animationbehaviors = animationbehaviors

    @property
    def Application(self):
        return Application(self.animationbehaviors.Application)

    @property
    def Count(self):
        return self.animationbehaviors.Count

    @property
    def Parent(self):
        return self.animationbehaviors.Parent

    def Add(self, *args, Type=None, Index=None):
        arguments = {"Type": Type, "Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.animationbehaviors.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.animationbehaviors.Item(*args, **arguments)

class AnimationPoint:

    def __init__(self, animationpoint=None):
        self.animationpoint = animationpoint

    @property
    def Application(self):
        return Application(self.animationpoint.Application)

    @property
    def Formula(self):
        return self.animationpoint.Formula

    @Formula.setter
    def Formula(self, value):
        self.animationpoint.Formula = value

    @property
    def Parent(self):
        return self.animationpoint.Parent

    @property
    def Time(self):
        return self.animationpoint.Time

    @property
    def Value(self):
        return self.animationpoint.Value

    def Delete(self):
        self.animationpoint.Delete()

class AnimationPoints:

    def __init__(self, animationpoints=None):
        self.animationpoints = animationpoints

    @property
    def Application(self):
        return Application(self.animationpoints.Application)

    @property
    def Count(self):
        return self.animationpoints.Count

    @property
    def Parent(self):
        return self.animationpoints.Parent

    @property
    def Smooth(self):
        return self.animationpoints.Smooth

    def Add(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.animationpoints.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.animationpoints.Item(*args, **arguments)

class AnimationSettings:

    def __init__(self, animationsettings=None):
        self.animationsettings = animationsettings

    @property
    def AdvanceMode(self):
        return self.animationsettings.AdvanceMode

    @AdvanceMode.setter
    def AdvanceMode(self, value):
        self.animationsettings.AdvanceMode = value

    @property
    def AdvanceTime(self):
        return self.animationsettings.AdvanceTime

    @AdvanceTime.setter
    def AdvanceTime(self, value):
        self.animationsettings.AdvanceTime = value

    @property
    def AfterEffect(self):
        return PpAfterEffect(self.animationsettings.AfterEffect)

    @AfterEffect.setter
    def AfterEffect(self, value):
        self.animationsettings.AfterEffect = value

    @property
    def Animate(self):
        return self.animationsettings.Animate

    @property
    def AnimateBackground(self):
        return self.animationsettings.AnimateBackground

    @property
    def AnimateTextInReverse(self):
        return self.animationsettings.AnimateTextInReverse

    @property
    def AnimationOrder(self):
        return self.animationsettings.AnimationOrder

    @AnimationOrder.setter
    def AnimationOrder(self, value):
        self.animationsettings.AnimationOrder = value

    @property
    def Application(self):
        return Application(self.animationsettings.Application)

    @property
    def ChartUnitEffect(self):
        return self.animationsettings.ChartUnitEffect

    @ChartUnitEffect.setter
    def ChartUnitEffect(self, value):
        self.animationsettings.ChartUnitEffect = value

    @property
    def DimColor(self):
        return ColorFormat(self.animationsettings.DimColor)

    @DimColor.setter
    def DimColor(self, value):
        self.animationsettings.DimColor = value

    @property
    def EntryEffect(self):
        return self.animationsettings.EntryEffect

    @property
    def Parent(self):
        return self.animationsettings.Parent

    @property
    def PlaySettings(self):
        return PlaySettings(self.animationsettings.PlaySettings)

    @property
    def SoundEffect(self):
        return SoundEffect(self.animationsettings.SoundEffect)

    @property
    def TextLevelEffect(self):
        return self.animationsettings.TextLevelEffect

    @property
    def TextUnitEffect(self):
        return self.animationsettings.TextUnitEffect

class Application:

    def __init__(self, application=None):
        self.application = application

    def new(self):
        self.application = win32com.client.Dispatch("PowerPoint.Application")
        return self

    @property
    def Active(self):
        return self.application.Active

    @property
    def ActiveEncryptionSession(self):
        return self.application.ActiveEncryptionSession

    @property
    def ActivePresentation(self):
        return Presentation(self.application.ActivePresentation)

    @property
    def ActivePrinter(self):
        return self.application.ActivePrinter

    @property
    def ActiveProtectedViewWindow(self):
        return ProtectedViewWindow(self.application.ActiveProtectedViewWindow)

    @property
    def ActiveWindow(self):
        return DocumentWindow(self.application.ActiveWindow)

    @property
    def AddIns(self):
        return AddIns(self.application.AddIns)

    @property
    def Assistance(self):
        return self.application.Assistance

    @property
    def AutoCorrect(self):
        return AutoCorrect(self.application.AutoCorrect)

    @property
    def AutomationSecurity(self):
        return self.application.AutomationSecurity

    @property
    def Build(self):
        return self.application.Build

    @property
    def Caption(self):
        return self.application.Caption

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
    def DisplayAlerts(self):
        return self.application.DisplayAlerts

    @property
    def DisplayDocumentInformationPanel(self):
        return self.application.DisplayDocumentInformationPanel

    @DisplayDocumentInformationPanel.setter
    def DisplayDocumentInformationPanel(self, value):
        self.application.DisplayDocumentInformationPanel = value

    @property
    def DisplayGridLines(self):
        return self.application.DisplayGridLines

    @property
    def FeatureInstall(self):
        return self.application.FeatureInstall

    @FeatureInstall.setter
    def FeatureInstall(self, value):
        self.application.FeatureInstall = value

    def FileConverters(self, *args,  `_Index1_`=None, `_Index2_` =None):
        arguments = {" `_Index1_`":  `_Index1_`, "`_Index2_` ": `_Index2_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.FileConverters(*args, **arguments)

    def FileDialog(self, *args,  `_Type_` =None):
        arguments = {" `_Type_` ":  `_Type_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.FileDialog(*args, **arguments)

    @property
    def FileValidation(self):
        return self.application.FileValidation

    @FileValidation.setter
    def FileValidation(self, value):
        self.application.FileValidation = value

    @property
    def Height(self):
        return self.application.Height

    @Height.setter
    def Height(self, value):
        self.application.Height = value

    @property
    def IsSandboxed(self):
        return self.application.IsSandboxed

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
    def Name(self):
        return self.application.Name

    @property
    def NewPresentation(self):
        return self.application.NewPresentation

    @property
    def OperatingSystem(self):
        return self.application.OperatingSystem

    @property
    def Options(self):
        return Options(self.application.Options)

    @property
    def Path(self):
        return Application(self.application.Path)

    @property
    def Presentations(self):
        return Presentations(self.application.Presentations)

    @property
    def ProductCode(self):
        return self.application.ProductCode

    @property
    def ProtectedViewWindows(self):
        return ProtectedViewWindows(self.application.ProtectedViewWindows)

    @property
    def SensitivityLabelPolicy(self):
        return self.application.SensitivityLabelPolicy

    @property
    def ShowStartupDialog(self):
        return self.application.ShowStartupDialog

    @property
    def ShowWindowsInTaskbar(self):
        return self.application.ShowWindowsInTaskbar

    @property
    def SlideShowWindows(self):
        return SlideShowWindows(self.application.SlideShowWindows)

    @property
    def SmartArtColors(self):
        return Application(self.application.SmartArtColors)

    @property
    def SmartArtLayouts(self):
        return Application(self.application.SmartArtLayouts)

    @property
    def SmartArtQuickStyles(self):
        return Application(self.application.SmartArtQuickStyles)

    @property
    def Top(self):
        return self.application.Top

    @Top.setter
    def Top(self, value):
        self.application.Top = value

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
        return DocumentWindows(self.application.Windows)

    @property
    def WindowState(self):
        return self.application.WindowState

    @WindowState.setter
    def WindowState(self, value):
        self.application.WindowState = value

    def Activate(self):
        self.application.Activate()

    def Help(self, *args,  `_HelpFile_`=None, `_ContextID_` =None):
        arguments = {" `_HelpFile_`":  `_HelpFile_`, "`_ContextID_` ": `_ContextID_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.Help(*args, **arguments)

    def Quit(self):
        self.application.Quit()

    def Run(self, *args, MacroName=None, safeArrayOfParams=None):
        arguments = {"MacroName": MacroName, "safeArrayOfParams": safeArrayOfParams}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.Run(*args, **arguments)

    def StartNewUndoEntry(self):
        self.application.StartNewUndoEntry()

class AutoCorrect:

    def __init__(self, autocorrect=None):
        self.autocorrect = autocorrect

    @property
    def DisplayAutoCorrectOptions(self):
        return self.autocorrect.DisplayAutoCorrectOptions

    @property
    def DisplayAutoLayoutOptions(self):
        return self.autocorrect.DisplayAutoLayoutOptions

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
        return XlAxisGroup(self.axis.AxisGroup)

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
        return XlScaleType(self.axis.ScaleType)

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
        return XlAxisType(self.axis.Type)

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

    def Characters(self, *args,  `_Start_`=None, `_Length_` =None):
        arguments = {" `_Start_`":  `_Start_`, "`_Length_` ": `_Length_` }
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

class Borders:

    def __init__(self, borders=None):
        self.borders = borders

    def __call__(self, item):
        return Border(self.borders(item))

    @property
    def Application(self):
        return Application(self.borders.Application)

    @property
    def Count(self):
        return self.borders.Count

    @property
    def Parent(self):
        return self.borders.Parent

    def Item(self, *args, BorderType=None):
        arguments = {"BorderType": BorderType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.borders.Item(*args, **arguments)

class Broadcast:

    def __init__(self, broadcast=None):
        self.broadcast = broadcast

    @property
    def Application(self):
        return Application(self.broadcast.Application)

    @property
    def AttendeeUrl(self):
        return self.broadcast.AttendeeUrl

    @property
    def IsBroadcasting(self):
        return self.broadcast.IsBroadcasting

    @property
    def Parent(self):
        return self.broadcast.Parent

    def End(self):
        return self.broadcast.End()

    def Start(self, *args, serverUrl=None):
        arguments = {"serverUrl": serverUrl}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.broadcast.Start(*args, **arguments)

class BulletFormat:

    def __init__(self, bulletformat=None):
        self.bulletformat = bulletformat

    @property
    def Application(self):
        return Application(self.bulletformat.Application)

    @property
    def Character(self):
        return self.bulletformat.Character

    @Character.setter
    def Character(self, value):
        self.bulletformat.Character = value

    @property
    def Font(self):
        return Font(self.bulletformat.Font)

    @property
    def Number(self):
        return self.bulletformat.Number

    @property
    def Parent(self):
        return self.bulletformat.Parent

    @property
    def RelativeSize(self):
        return self.bulletformat.RelativeSize

    @RelativeSize.setter
    def RelativeSize(self, value):
        self.bulletformat.RelativeSize = value

    @property
    def StartValue(self):
        return self.bulletformat.StartValue

    @StartValue.setter
    def StartValue(self, value):
        self.bulletformat.StartValue = value

    @property
    def Style(self):
        return self.bulletformat.Style

    @Style.setter
    def Style(self, value):
        self.bulletformat.Style = value

    @property
    def Type(self):
        return self.bulletformat.Type

    @property
    def UseTextColor(self):
        return self.bulletformat.UseTextColor

    @property
    def UseTextFont(self):
        return self.bulletformat.UseTextFont

    def Picture(self):
        self.bulletformat.Picture()

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
    def AutoAttach(self):
        return self.calloutformat.AutoAttach

    @property
    def AutoLength(self):
        return self.calloutformat.AutoLength

    @property
    def Border(self):
        return self.calloutformat.Border

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

    def AutomaticLength(self):
        self.calloutformat.AutomaticLength()

    def CustomDrop(self, *args, Drop=None):
        arguments = {"Drop": Drop}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.calloutformat.CustomDrop(*args, **arguments)

    def CustomLength(self, *args, Length=None):
        arguments = {"Length": Length}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.calloutformat.CustomLength(*args, **arguments)

    def PresetDrop(self, *args, DropType=None):
        arguments = {"DropType": DropType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.calloutformat.PresetDrop(*args, **arguments)

class Cell:

    def __init__(self, cell=None):
        self.cell = cell

    @property
    def Application(self):
        return Application(self.cell.Application)

    @property
    def Borders(self):
        return Borders(self.cell.Borders)

    @property
    def Parent(self):
        return self.cell.Parent

    @property
    def Selected(self):
        return self.cell.Selected

    @property
    def Shape(self):
        return Shape(self.cell.Shape)

    def Merge(self, *args, MergeTo=None):
        arguments = {"MergeTo": MergeTo}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.cell.Merge(*args, **arguments)

    def Select(self):
        self.cell.Select()

    def Split(self, *args, NumRows=None, NumColumns=None):
        arguments = {"NumRows": NumRows, "NumColumns": NumColumns}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.cell.Split(*args, **arguments)

class CellRange:

    def __init__(self, cellrange=None):
        self.cellrange = cellrange

    def __call__(self, item):
        return CellRang(self.cellrange(item))

    @property
    def Application(self):
        return Application(self.cellrange.Application)

    @property
    def Borders(self):
        return Borders(self.cellrange.Borders)

    @property
    def Count(self):
        return self.cellrange.Count

    @property
    def Parent(self):
        return self.cellrange.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.cellrange.Item(*args, **arguments)

class Chart:

    def __init__(self, chart=None):
        self.chart = chart

    @property
    def AlternativeText(self):
        return self.chart.AlternativeText

    @AlternativeText.setter
    def AlternativeText(self, value):
        self.chart.AlternativeText = value

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
        return XlBarShape(self.chart.BarShape)

    @BarShape.setter
    def BarShape(self, value):
        self.chart.BarShape = value

    @property
    def ChartArea(self):
        return ChartArea(self.chart.ChartArea)

    @property
    def ChartData(self):
        return ChartData(self.chart.ChartData)

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
    def Format(self):
        return self.chart.Format

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
    def Name(self):
        return self.chart.Name

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
    def Title(self):
        return self.chart.Title

    @property
    def Walls(self):
        return Walls(self.chart.Walls)

    def ApplyChartTemplate(self, *args,  `_FileName_` =None):
        arguments = {" `_FileName_` ":  `_FileName_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.ApplyChartTemplate(*args, **arguments)

    def ApplyDataLabels(self, *args, Type=None, LegendKey=None, AutoText=None, HasLeaderLines=None, ShowSeriesName=None, ShowCategoryName=None, ShowValue=None, ShowPercentage=None, ShowBubbleSize=None, Separator=None):
        arguments = {"Type": Type, "LegendKey": LegendKey, "AutoText": AutoText, "HasLeaderLines": HasLeaderLines, "ShowSeriesName": ShowSeriesName, "ShowCategoryName": ShowCategoryName, "ShowValue": ShowValue, "ShowPercentage": ShowPercentage, "ShowBubbleSize": ShowBubbleSize, "Separator": Separator}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.ApplyDataLabels(*args, **arguments)

    def ApplyLayout(self, *args,  `_Layout_`=None, `_ChartType_` =None):
        arguments = {" `_Layout_`":  `_Layout_`, "`_ChartType_` ": `_ChartType_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.ApplyLayout(*args, **arguments)

    def Axes(self, *args, Type=None, AxisGroup=None):
        arguments = {"Type": Type, "AxisGroup": AxisGroup}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.chart.Axes(*args, **arguments)

    def ChartGroups(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.ChartGroups(*args, **arguments)

    def ChartWizard(self, *args, **_Source_**=None, **_Gallery_**=None, **_Format_**=None, **_PlotBy_**=None, **_CategoryLabels_**=None, **_SeriesLabels_**=None, **_HasLegend_**=None, **_Title_**=None, **_CategoryTitle_**=None, **_ValueTitle_**=None, **_ExtraTitle_**=None):
        arguments = {"**_Source_**": **_Source_**, "**_Gallery_**": **_Gallery_**, "**_Format_**": **_Format_**, "**_PlotBy_**": **_PlotBy_**, "**_CategoryLabels_**": **_CategoryLabels_**, "**_SeriesLabels_**": **_SeriesLabels_**, "**_HasLegend_**": **_HasLegend_**, "**_Title_**": **_Title_**, "**_CategoryTitle_**": **_CategoryTitle_**, "**_ValueTitle_**": **_ValueTitle_**, "**_ExtraTitle_**": **_ExtraTitle_**}
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
        self.chart.Export(*args, **arguments)

    def GetChartElement(self, *args,  `_x_`=None, `_y_`=None, `_ElementID_`=None, `_Arg1_`=None, `_Arg2_` =None):
        arguments = {" `_x_`":  `_x_`, "`_y_`": `_y_`, "`_ElementID_`": `_ElementID_`, "`_Arg1_`": `_Arg1_`, "`_Arg2_` ": `_Arg2_` }
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

    def SetSourceData(self, *args,  `_Source_`=None, `_PlotBy_` =None):
        arguments = {" `_Source_`":  `_Source_`, "`_PlotBy_` ": `_PlotBy_` }
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
        return XlLineStyle(self.chartborder.LineStyle)

    @LineStyle.setter
    def LineStyle(self, value):
        self.chartborder.LineStyle = value

    @property
    def Parent(self):
        return self.chartborder.Parent

    @property
    def Weight(self):
        return XlBorderWeight(self.chartborder.Weight)

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
        return XlBackground(self.chartfont.Background)

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
    def Underline(self):
        return XlUnderlineStyle(self.chartfont.Underline)

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
        return XlAxisGroup(self.chartgroup.AxisGroup)

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

    def Characters(self, *args,  `_Start_`=None, `_Length_` =None):
        arguments = {" `_Start_`":  `_Start_`, "`_Length_` ": `_Length_` }
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

    @Height.setter
    def Height(self, value):
        self.charttitle.Height = value

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
        return XlChartElementPosition(self.charttitle.Position)

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

    @Width.setter
    def Width(self, value):
        self.charttitle.Width = value

    def Delete(self):
        self.charttitle.Delete()

    def Select(self):
        self.charttitle.Select()

class Coauthoring:

    def __init__(self, coauthoring=None):
        self.coauthoring = coauthoring

    @property
    def Application(self):
        return Application(self.coauthoring.Application)

    @property
    def FavorServerEditsDuringMerge(self):
        return self.coauthoring.FavorServerEditsDuringMerge

    @property
    def MergeMode(self):
        return self.coauthoring.MergeMode

    @property
    def Parent(self):
        return self.coauthoring.Parent

    @property
    def PendingUpdates(self):
        return self.coauthoring.PendingUpdates

    def EndReview(self):
        self.coauthoring.EndReview()

class ColorEffect:

    def __init__(self, coloreffect=None):
        self.coloreffect = coloreffect

    @property
    def Application(self):
        return Application(self.coloreffect.Application)

    @property
    def By(self):
        return ColorFormat(self.coloreffect.By)

    @property
    def From(self):
        return self.coloreffect.From

    @property
    def Parent(self):
        return self.coloreffect.Parent

    @property
    def To(self):
        return self.coloreffect.To

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
        return ColorFormat(self.colorformat.ObjectThemeColor)

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

    @property
    def Type(self):
        return self.colorformat.Type

class ColorScheme:

    def __init__(self, colorscheme=None):
        self.colorscheme = colorscheme

    @property
    def Application(self):
        return Application(self.colorscheme.Application)

    @property
    def Count(self):
        return self.colorscheme.Count

    @property
    def Parent(self):
        return self.colorscheme.Parent

    def Colors(self, *args, SchemeColor=None):
        arguments = {"SchemeColor": SchemeColor}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.colorscheme.Colors(*args, **arguments)

    def Delete(self):
        self.colorscheme.Delete()

class ColorSchemes:

    def __init__(self, colorschemes=None):
        self.colorschemes = colorschemes

    def __call__(self, item):
        return ColorScheme(self.colorschemes(item))

    @property
    def Application(self):
        return Application(self.colorschemes.Application)

    @property
    def Count(self):
        return self.colorschemes.Count

    @property
    def Parent(self):
        return self.colorschemes.Parent

    def Add(self, *args, Scheme=None):
        arguments = {"Scheme": Scheme}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return ColorScheme(self.colorschemes.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.colorschemes.Item(*args, **arguments)

class Column:

    def __init__(self, column=None):
        self.column = column

    @property
    def Application(self):
        return Application(self.column.Application)

    def Cells(self, *args, RowIndex=None, ColumnIndex=None):
        arguments = {"RowIndex": RowIndex, "ColumnIndex": ColumnIndex}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return CellRange(self.column.Cells(*args, **arguments))

    @property
    def Parent(self):
        return self.column.Parent

    @property
    def Width(self):
        return self.column.Width

    @Width.setter
    def Width(self, value):
        self.column.Width = value

    def Delete(self):
        self.column.Delete()

    def Select(self):
        self.column.Select()

class Columns:

    def __init__(self, columns=None):
        self.columns = columns

    def __call__(self, item):
        return Column(self.columns(item))

    @property
    def Application(self):
        return Application(self.columns.Application)

    @property
    def Count(self):
        return self.columns.Count

    @property
    def Parent(self):
        return self.columns.Parent

    def Add(self, *args, BeforeColumn=None):
        arguments = {"BeforeColumn": BeforeColumn}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Column(self.columns.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.columns.Item(*args, **arguments)

class CommandEffect:

    def __init__(self, commandeffect=None):
        self.commandeffect = commandeffect

    @property
    def Application(self):
        return Application(self.commandeffect.Application)

    @property
    def Bookmark(self):
        return self.commandeffect.Bookmark

    @property
    def Command(self):
        return self.commandeffect.Command

    @property
    def Parent(self):
        return self.commandeffect.Parent

    @property
    def Type(self):
        return self.commandeffect.Type

class Comment:

    def __init__(self, comment=None):
        self.comment = comment

    @property
    def Application(self):
        return Application(self.comment.Application)

    @property
    def Author(self):
        return Comment(self.comment.Author)

    @property
    def AuthorIndex(self):
        return self.comment.AuthorIndex

    @property
    def AuthorInitials(self):
        return Comment(self.comment.AuthorInitials)

    @property
    def DateTime(self):
        return self.comment.DateTime

    @property
    def Left(self):
        return self.comment.Left

    @property
    def Parent(self):
        return self.comment.Parent

    @property
    def Text(self):
        return self.comment.Text

    @property
    def Top(self):
        return self.comment.Top

    def Delete(self):
        self.comment.Delete()

class Comments:

    def __init__(self, comments=None):
        self.comments = comments

    @property
    def Application(self):
        return Application(self.comments.Application)

    @property
    def Count(self):
        return self.comments.Count

    @property
    def Parent(self):
        return self.comments.Parent

    def Add(self, *args, Left=None, Top=None, Author=None, AuthorInitials=None, Text=None):
        arguments = {"Left": Left, "Top": Top, "Author": Author, "AuthorInitials": AuthorInitials, "Text": Text}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.comments.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.comments.Item(*args, **arguments)

class ConnectorFormat:

    def __init__(self, connectorformat=None):
        self.connectorformat = connectorformat

    @property
    def Application(self):
        return Application(self.connectorformat.Application)

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

    def BeginConnect(self, *args,  `_ConnectedShape_`=None, `_ConnectionSite_` =None):
        arguments = {" `_ConnectedShape_`":  `_ConnectedShape_`, "`_ConnectionSite_` ": `_ConnectionSite_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.connectorformat.BeginConnect(*args, **arguments)

    def BeginDisconnect(self):
        self.connectorformat.BeginDisconnect()

    def EndConnect(self, *args,  `_ConnectedShape_`=None, `_ConnectionSite_` =None):
        arguments = {" `_ConnectedShape_`":  `_ConnectedShape_`, "`_ConnectionSite_` ": `_ConnectionSite_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.connectorformat.EndConnect(*args, **arguments)

    def EndDisconnect(self):
        self.connectorformat.EndDisconnect()

class CustomerData:

    def __init__(self, customerdata=None):
        self.customerdata = customerdata

    @property
    def Application(self):
        return Application(self.customerdata.Application)

    @property
    def Count(self):
        return self.customerdata.Count

    @property
    def Parent(self):
        return CustomerData(self.customerdata.Parent)

    def Add(self):
        return self.customerdata.Add()

    def Delete(self, *args,  `_Id_` =None):
        arguments = {" `_Id_` ":  `_Id_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.customerdata.Delete(*args, **arguments)

    def Item(self, *args, Id=None):
        arguments = {"Id": Id}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.customerdata.Item(*args, **arguments)

class CustomLayout:

    def __init__(self, customlayout=None):
        self.customlayout = customlayout

    @property
    def Application(self):
        return Application(self.customlayout.Application)

    @property
    def Background(self):
        return ShapeRange(self.customlayout.Background)

    @property
    def CustomerData(self):
        return CustomerData(self.customlayout.CustomerData)

    @property
    def Design(self):
        return Design(self.customlayout.Design)

    @property
    def DisplayMasterShapes(self):
        return self.customlayout.DisplayMasterShapes

    @property
    def FollowMasterBackground(self):
        return self.customlayout.FollowMasterBackground

    @property
    def HeadersFooters(self):
        return HeadersFooters(self.customlayout.HeadersFooters)

    @property
    def Height(self):
        return self.customlayout.Height

    @property
    def Hyperlinks(self):
        return Hyperlinks(self.customlayout.Hyperlinks)

    @property
    def Index(self):
        return CustomLayouts(self.customlayout.Index)

    @property
    def MatchingName(self):
        return self.customlayout.MatchingName

    @property
    def Name(self):
        return self.customlayout.Name

    @property
    def Parent(self):
        return CustomLayout(self.customlayout.Parent)

    @property
    def Preserved(self):
        return self.customlayout.Preserved

    @property
    def Shapes(self):
        return Shapes(self.customlayout.Shapes)

    @property
    def SlideShowTransition(self):
        return SlideShowTransition(self.customlayout.SlideShowTransition)

    @property
    def ThemeColorScheme(self):
        return self.customlayout.ThemeColorScheme

    @property
    def TimeLine(self):
        return TimeLine(self.customlayout.TimeLine)

    @property
    def Width(self):
        return self.customlayout.Width

    def Copy(self):
        self.customlayout.Copy()

    def Cut(self):
        self.customlayout.Cut()

    def Delete(self):
        self.customlayout.Delete()

    def Duplicate(self):
        return self.customlayout.Duplicate()

    def MoveTo(self, *args,  `_toPos_` =None):
        arguments = {" `_toPos_` ":  `_toPos_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.customlayout.MoveTo(*args, **arguments)

    def Select(self):
        self.customlayout.Select()

class CustomLayouts:

    def __init__(self, customlayouts=None):
        self.customlayouts = customlayouts

    @property
    def Application(self):
        return Application(self.customlayouts.Application)

    @property
    def Count(self):
        return self.customlayouts.Count

    @property
    def Parent(self):
        return self.customlayouts.Parent

    def Add(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.customlayouts.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.customlayouts.Item(*args, **arguments)

    def Paste(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.customlayouts.Paste(*args, **arguments)

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

    def Characters(self, *args,  `_Start_`=None, `_Length_` =None):
        arguments = {" `_Start_`":  `_Start_`, "`_Length_` ": `_Length_` }
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
        return XlDataLabelPosition(self.datalabel.Position)

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

class Design:

    def __init__(self, design=None):
        self.design = design

    @property
    def Application(self):
        return Application(self.design.Application)

    @property
    def Index(self):
        return self.design.Index

    @property
    def Name(self):
        return self.design.Name

    @Name.setter
    def Name(self, value):
        self.design.Name = value

    @property
    def Parent(self):
        return self.design.Parent

    @property
    def Preserved(self):
        return self.design.Preserved

    @property
    def SlideMaster(self):
        return Master(self.design.SlideMaster)

    def Delete(self):
        self.design.Delete()

    def MoveTo(self, *args,  `_toPos_` =None):
        arguments = {" `_toPos_` ":  `_toPos_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.design.MoveTo(*args, **arguments)

class Designs:

    def __init__(self, designs=None):
        self.designs = designs

    @property
    def Application(self):
        return Application(self.designs.Application)

    @property
    def Count(self):
        return self.designs.Count

    @property
    def Parent(self):
        return self.designs.Parent

    def Add(self, *args, designName=None, Index=None):
        arguments = {"designName": designName, "Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.designs.Add(*args, **arguments)

    def Clone(self, *args,  `_pOriginal_`=None, `_Index_` =None):
        arguments = {" `_pOriginal_`":  `_pOriginal_`, "`_Index_` ": `_Index_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.designs.Clone(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.designs.Item(*args, **arguments)

    def Load(self, *args,  `_TemplateName_`=None, `_Index_` =None):
        arguments = {" `_TemplateName_`":  `_TemplateName_`, "`_Index_` ": `_Index_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.designs.Load(*args, **arguments)

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

    def Characters(self, *args,  `_Start_`=None, `_Length_` =None):
        arguments = {" `_Start_`":  `_Start_`, "`_Length_` ": `_Length_` }
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

class DocumentWindow:

    def __init__(self, documentwindow=None):
        self.documentwindow = documentwindow

    @property
    def Active(self):
        return self.documentwindow.Active

    @property
    def ActivePane(self):
        return Pane(self.documentwindow.ActivePane)

    @property
    def Application(self):
        return Application(self.documentwindow.Application)

    @property
    def BlackAndWhite(self):
        return self.documentwindow.BlackAndWhite

    @property
    def Caption(self):
        return self.documentwindow.Caption

    @property
    def Height(self):
        return self.documentwindow.Height

    @Height.setter
    def Height(self, value):
        self.documentwindow.Height = value

    @property
    def Left(self):
        return self.documentwindow.Left

    @Left.setter
    def Left(self, value):
        self.documentwindow.Left = value

    @property
    def Panes(self):
        return Panes(self.documentwindow.Panes)

    @property
    def Parent(self):
        return self.documentwindow.Parent

    @property
    def Presentation(self):
        return Presentation(self.documentwindow.Presentation)

    @property
    def Selection(self):
        return Selection(self.documentwindow.Selection)

    @property
    def SplitHorizontal(self):
        return self.documentwindow.SplitHorizontal

    @SplitHorizontal.setter
    def SplitHorizontal(self, value):
        self.documentwindow.SplitHorizontal = value

    @property
    def SplitVertical(self):
        return self.documentwindow.SplitVertical

    @SplitVertical.setter
    def SplitVertical(self, value):
        self.documentwindow.SplitVertical = value

    @property
    def Top(self):
        return self.documentwindow.Top

    @Top.setter
    def Top(self, value):
        self.documentwindow.Top = value

    @property
    def View(self):
        return View(self.documentwindow.View)

    @property
    def ViewType(self):
        return self.documentwindow.ViewType

    @ViewType.setter
    def ViewType(self, value):
        self.documentwindow.ViewType = value

    @property
    def Width(self):
        return self.documentwindow.Width

    @Width.setter
    def Width(self, value):
        self.documentwindow.Width = value

    @property
    def WindowState(self):
        return self.documentwindow.WindowState

    @WindowState.setter
    def WindowState(self, value):
        self.documentwindow.WindowState = value

    def Activate(self):
        self.documentwindow.Activate()

    def Close(self):
        self.documentwindow.Close()

    def ExpandSection(self, *args,  `_sectionIndex_`=None, `_Expand_` =None):
        arguments = {" `_sectionIndex_`":  `_sectionIndex_`, "`_Expand_` ": `_Expand_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.documentwindow.ExpandSection(*args, **arguments)

    def FitToPage(self):
        self.documentwindow.FitToPage()

    def IsSectionExpanded(self, *args,  `_sectionIndex_` =None):
        arguments = {" `_sectionIndex_` ":  `_sectionIndex_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.documentwindow.IsSectionExpanded(*args, **arguments)

    def LargeScroll(self, *args, Down=None, Up=None, ToRight=None, ToLeft=None):
        arguments = {"Down": Down, "Up": Up, "ToRight": ToRight, "ToLeft": ToLeft}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.documentwindow.LargeScroll(*args, **arguments)

    def NewWindow(self):
        return self.documentwindow.NewWindow()

    def PointsToScreenPixelsX(self, *args, Points=None):
        arguments = {"Points": Points}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.documentwindow.PointsToScreenPixelsX(*args, **arguments)

    def PointsToScreenPixelsY(self, *args, Points=None):
        arguments = {"Points": Points}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.documentwindow.PointsToScreenPixelsY(*args, **arguments)

    def RangeFromPoint(self, *args, x=None, y=None):
        arguments = {"x": x, "y": y}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.documentwindow.RangeFromPoint(*args, **arguments)

    def ScrollIntoView(self, *args, Left=None, Top=None, Width=None, Height=None, Start=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height, "Start": Start}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.documentwindow.ScrollIntoView(*args, **arguments)

    def SmallScroll(self, *args, Down=None, Up=None, ToRight=None, ToLeft=None):
        arguments = {"Down": Down, "Up": Up, "ToRight": ToRight, "ToLeft": ToLeft}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.documentwindow.SmallScroll(*args, **arguments)

class DocumentWindows:

    def __init__(self, documentwindows=None):
        self.documentwindows = documentwindows

    def __call__(self, item):
        return DocumentWindow(self.documentwindows(item))

    @property
    def Application(self):
        return Application(self.documentwindows.Application)

    @property
    def Count(self):
        return self.documentwindows.Count

    @property
    def Parent(self):
        return self.documentwindows.Parent

    def Arrange(self, *args,  `_arrangeStyle_` =None):
        arguments = {" `_arrangeStyle_` ":  `_arrangeStyle_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.documentwindows.Arrange(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.documentwindows.Item(*args, **arguments)

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

class Effect:

    def __init__(self, effect=None):
        self.effect = effect

    @property
    def Application(self):
        return Application(self.effect.Application)

    @property
    def Behaviors(self):
        return AnimationBehaviors(self.effect.Behaviors)

    @property
    def DisplayName(self):
        return self.effect.DisplayName

    @property
    def EffectInformation(self):
        return EffectInformation(self.effect.EffectInformation)

    @property
    def EffectParameters(self):
        return EffectParameters(self.effect.EffectParameters)

    @property
    def EffectType(self):
        return self.effect.EffectType

    @property
    def Exit(self):
        return self.effect.Exit

    @property
    def Index(self):
        return self.effect.Index

    @property
    def Paragraph(self):
        return self.effect.Paragraph

    @Paragraph.setter
    def Paragraph(self, value):
        self.effect.Paragraph = value

    @property
    def Parent(self):
        return self.effect.Parent

    @property
    def Shape(self):
        return Shape(self.effect.Shape)

    @property
    def TextRangeLength(self):
        return self.effect.TextRangeLength

    @TextRangeLength.setter
    def TextRangeLength(self, value):
        self.effect.TextRangeLength = value

    @property
    def TextRangeStart(self):
        return self.effect.TextRangeStart

    @TextRangeStart.setter
    def TextRangeStart(self, value):
        self.effect.TextRangeStart = value

    @property
    def Timing(self):
        return Timing(self.effect.Timing)

    def Delete(self):
        self.effect.Delete()

    def MoveAfter(self, *args,  `_Effect_` =None):
        arguments = {" `_Effect_` ":  `_Effect_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.effect.MoveAfter(*args, **arguments)

    def MoveBefore(self, *args,  `_Effect_` =None):
        arguments = {" `_Effect_` ":  `_Effect_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.effect.MoveBefore(*args, **arguments)

    def MoveTo(self, *args,  `_toPos_` =None):
        arguments = {" `_toPos_` ":  `_toPos_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.effect.MoveTo(*args, **arguments)

class EffectInformation:

    def __init__(self, effectinformation=None):
        self.effectinformation = effectinformation

    @property
    def AfterEffect(self):
        return PpAfterEffect(self.effectinformation.AfterEffect)

    @property
    def AnimateBackground(self):
        return self.effectinformation.AnimateBackground

    @property
    def AnimateTextInReverse(self):
        return self.effectinformation.AnimateTextInReverse

    @property
    def Application(self):
        return Application(self.effectinformation.Application)

    @property
    def BuildByLevelEffect(self):
        return self.effectinformation.BuildByLevelEffect

    @property
    def Dim(self):
        return ColorFormat(self.effectinformation.Dim)

    @property
    def Parent(self):
        return self.effectinformation.Parent

    @property
    def PlaySettings(self):
        return PlaySettings(self.effectinformation.PlaySettings)

    @property
    def SoundEffect(self):
        return SoundEffect(self.effectinformation.SoundEffect)

    @property
    def TextUnitEffect(self):
        return self.effectinformation.TextUnitEffect

class EffectParameters:

    def __init__(self, effectparameters=None):
        self.effectparameters = effectparameters

    @property
    def Amount(self):
        return self.effectparameters.Amount

    @Amount.setter
    def Amount(self, value):
        self.effectparameters.Amount = value

    @property
    def Application(self):
        return Application(self.effectparameters.Application)

    @property
    def Color2(self):
        return ColorFormat(self.effectparameters.Color2)

    @property
    def Direction(self):
        return self.effectparameters.Direction

    @property
    def FontName(self):
        return self.effectparameters.FontName

    @FontName.setter
    def FontName(self, value):
        self.effectparameters.FontName = value

    @property
    def Parent(self):
        return self.effectparameters.Parent

    @property
    def Relative(self):
        return self.effectparameters.Relative

    @property
    def Size(self):
        return self.effectparameters.Size

    @Size.setter
    def Size(self, value):
        self.effectparameters.Size = value

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

class ExtraColors:

    def __init__(self, extracolors=None):
        self.extracolors = extracolors

    @property
    def Application(self):
        return Application(self.extracolors.Application)

    @property
    def Count(self):
        return self.extracolors.Count

    @property
    def Parent(self):
        return self.extracolors.Parent

    def Add(self, *args, Type=None):
        arguments = {"Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.extracolors.Add(*args, **arguments)

    def Clear(self):
        self.extracolors.Clear()

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return MsoThemeColorSchemeIndex(self.extracolors.Item(*args, **arguments))

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
        return FileConverter(self.fileconverter.Parent)

    @property
    def Path(self):
        return self.fileconverter.Path

    @property
    def SaveFormat(self):
        return self.fileconverter.SaveFormat

class FileConverters:

    def __init__(self, fileconverters=None):
        self.fileconverters = fileconverters

    def __call__(self, item):
        return FileConverter(self.fileconverters(item))

    @property
    def Count(self):
        return self.fileconverters.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.fileconverters.Item(*args, **arguments)

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

    def Background(self):
        self.fillformat.Background()

    def OneColorGradient(self, *args,  `_Style_`=None, `_Variant_`=None, `_Degree_` =None):
        arguments = {" `_Style_`":  `_Style_`, "`_Variant_`": `_Variant_`, "`_Degree_` ": `_Degree_` }
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

class FilterEffect:

    def __init__(self, filtereffect=None):
        self.filtereffect = filtereffect

    @property
    def Application(self):
        return Application(self.filtereffect.Application)

    @property
    def Parent(self):
        return self.filtereffect.Parent

    @property
    def Reveal(self):
        return self.filtereffect.Reveal

    @property
    def Subtype(self):
        return self.filtereffect.Subtype

    @property
    def Type(self):
        return self.filtereffect.Type

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
    def Application(self):
        return Application(self.font.Application)

    @property
    def AutoRotateNumbers(self):
        return self.font.AutoRotateNumbers

    @AutoRotateNumbers.setter
    def AutoRotateNumbers(self, value):
        self.font.AutoRotateNumbers = value

    @property
    def BaselineOffset(self):
        return self.font.BaselineOffset

    @BaselineOffset.setter
    def BaselineOffset(self, value):
        self.font.BaselineOffset = value

    @property
    def Bold(self):
        return self.font.Bold

    @property
    def Color(self):
        return Font(self.font.Color)

    @Color.setter
    def Color(self, value):
        self.font.Color = value

    @property
    def Embeddable(self):
        return self.font.Embeddable

    @property
    def Embedded(self):
        return self.font.Embedded

    @property
    def Emboss(self):
        return self.font.Emboss

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
    def NameComplexScript(self):
        return self.font.NameComplexScript

    @NameComplexScript.setter
    def NameComplexScript(self, value):
        self.font.NameComplexScript = value

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
    def Parent(self):
        return self.font.Parent

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
    def Subscript(self):
        return self.font.Subscript

    @property
    def Superscript(self):
        return self.font.Superscript

    @property
    def Underline(self):
        return self.font.Underline

class Fonts:

    def __init__(self, fonts=None):
        self.fonts = fonts

    def __call__(self, item):
        return Font(self.fonts(item))

    @property
    def Application(self):
        return Application(self.fonts.Application)

    @property
    def Count(self):
        return self.fonts.Count

    @property
    def Parent(self):
        return self.fonts.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.fonts.Item(*args, **arguments)

    def Replace(self, *args, Original=None, Replacement=None):
        arguments = {"Original": Original, "Replacement": Replacement}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.fonts.Replace(*args, **arguments)

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

    def AddNodes(self, *args, **_SegmentType_**=None, **_EditingType_**=None, **_X1_**=None, **_Y1_**=None, **_X2_**=None, **_Y2_**=None, **_X3_**=None, **_Y3_**=None):
        arguments = {"**_SegmentType_**": **_SegmentType_**, "**_EditingType_**": **_EditingType_**, "**_X1_**": **_X1_**, "**_Y1_**": **_Y1_**, "**_X2_**": **_X2_**, "**_Y2_**": **_Y2_**, "**_X3_**": **_X3_**, "**_Y3_**": **_Y3_**}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.freeformbuilder.AddNodes(*args, **arguments)

    def ConvertToShape(self):
        return self.freeformbuilder.ConvertToShape()

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
    def Name(self):
        return self.gridlines.Name

    @property
    def Parent(self):
        return self.gridlines.Parent

    def Delete(self):
        self.gridlines.Delete()

    def Select(self):
        self.gridlines.Select()

class GroupShapes:

    def __init__(self, groupshapes=None):
        self.groupshapes = groupshapes

    @property
    def Application(self):
        return Application(self.groupshapes.Application)

    @property
    def Count(self):
        return self.groupshapes.Count

    @property
    def Creator(self):
        return self.groupshapes.Creator

    @property
    def Parent(self):
        return self.groupshapes.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.groupshapes.Item(*args, **arguments)

    def Range(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.groupshapes.Range(*args, **arguments)

class HeaderFooter:

    def __init__(self, headerfooter=None):
        self.headerfooter = headerfooter

    @property
    def Application(self):
        return Application(self.headerfooter.Application)

    @property
    def Format(self):
        return self.headerfooter.Format

    @Format.setter
    def Format(self, value):
        self.headerfooter.Format = value

    @property
    def Parent(self):
        return self.headerfooter.Parent

    @property
    def Text(self):
        return self.headerfooter.Text

    @Text.setter
    def Text(self, value):
        self.headerfooter.Text = value

    @property
    def UseFormat(self):
        return self.headerfooter.UseFormat

    @property
    def Visible(self):
        return self.headerfooter.Visible

    @Visible.setter
    def Visible(self, value):
        self.headerfooter.Visible = value

class HeadersFooters:

    def __init__(self, headersfooters=None):
        self.headersfooters = headersfooters

    @property
    def Application(self):
        return Application(self.headersfooters.Application)

    @property
    def DateAndTime(self):
        return HeaderFooter(self.headersfooters.DateAndTime)

    @property
    def DisplayOnTitleSlide(self):
        return self.headersfooters.DisplayOnTitleSlide

    @property
    def Footer(self):
        return HeaderFooter(self.headersfooters.Footer)

    @property
    def Header(self):
        return HeaderFooter(self.headersfooters.Header)

    @property
    def Parent(self):
        return self.headersfooters.Parent

    @property
    def SlideNumber(self):
        return HeaderFooter(self.headersfooters.SlideNumber)

    def Clear(self):
        self.headersfooters.Clear()

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
    def EmailSubject(self):
        return self.hyperlink.EmailSubject

    @EmailSubject.setter
    def EmailSubject(self, value):
        self.hyperlink.EmailSubject = value

    @property
    def Parent(self):
        return self.hyperlink.Parent

    @property
    def ScreenTip(self):
        return self.hyperlink.ScreenTip

    @ScreenTip.setter
    def ScreenTip(self, value):
        self.hyperlink.ScreenTip = value

    @property
    def ShowAndReturn(self):
        return self.hyperlink.ShowAndReturn

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

    def CreateNewDocument(self, *args,  `_FileName_`=None, `_EditNow_`=None, `_Overwrite_` =None):
        arguments = {" `_FileName_`":  `_FileName_`, "`_EditNow_`": `_EditNow_`, "`_Overwrite_` ": `_Overwrite_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.hyperlink.CreateNewDocument(*args, **arguments)

    def Delete(self):
        self.hyperlink.Delete()

    def Follow(self):
        self.hyperlink.Follow()

class Hyperlinks:

    def __init__(self, hyperlinks=None):
        self.hyperlinks = hyperlinks

    def __call__(self, item):
        return Hyperlink(self.hyperlinks(item))

    @property
    def Application(self):
        return Application(self.hyperlinks.Application)

    @property
    def Count(self):
        return self.hyperlinks.Count

    @property
    def Parent(self):
        return self.hyperlinks.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.hyperlinks.Item(*args, **arguments)

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
        return XlPattern(self.interior.Pattern)

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
        return XlColorIndex(self.interior.PatternColorIndex)

    @PatternColorIndex.setter
    def PatternColorIndex(self, value):
        self.interior.PatternColorIndex = value

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
        return XlColorIndex(self.legendkey.MarkerBackgroundColorIndex)

    @MarkerBackgroundColorIndex.setter
    def MarkerBackgroundColorIndex(self, value):
        self.legendkey.MarkerBackgroundColorIndex = value

    @property
    def MarkerForegroundColor(self):
        return self.legendkey.MarkerForegroundColor

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
        return Application(self.linkformat.Application)

    @property
    def AutoUpdate(self):
        return self.linkformat.AutoUpdate

    @AutoUpdate.setter
    def AutoUpdate(self, value):
        self.linkformat.AutoUpdate = value

    @property
    def Parent(self):
        return self.linkformat.Parent

    @property
    def SourceFullName(self):
        return self.linkformat.SourceFullName

    @SourceFullName.setter
    def SourceFullName(self, value):
        self.linkformat.SourceFullName = value

    def BreakLink(self):
        return self.linkformat.BreakLink()

    def Update(self):
        self.linkformat.Update()

class Master:

    def __init__(self, master=None):
        self.master = master

    @property
    def Application(self):
        return Application(self.master.Application)

    @property
    def Background(self):
        return ShapeRange(self.master.Background)

    @property
    def BackgroundStyle(self):
        return self.master.BackgroundStyle

    @property
    def ColorScheme(self):
        return ColorScheme(self.master.ColorScheme)

    @ColorScheme.setter
    def ColorScheme(self, value):
        self.master.ColorScheme = value

    @property
    def CustomerData(self):
        return CustomerData(self.master.CustomerData)

    @property
    def CustomLayouts(self):
        return CustomLayouts(self.master.CustomLayouts)

    @property
    def Design(self):
        return Design(self.master.Design)

    @property
    def HeadersFooters(self):
        return HeadersFooters(self.master.HeadersFooters)

    @property
    def Height(self):
        return self.master.Height

    @Height.setter
    def Height(self, value):
        self.master.Height = value

    @property
    def Hyperlinks(self):
        return Hyperlinks(self.master.Hyperlinks)

    @property
    def Name(self):
        return self.master.Name

    @Name.setter
    def Name(self, value):
        self.master.Name = value

    @property
    def Parent(self):
        return self.master.Parent

    @property
    def Shapes(self):
        return Shapes(self.master.Shapes)

    @property
    def SlideShowTransition(self):
        return SlideShowTransition(self.master.SlideShowTransition)

    @property
    def TextStyles(self):
        return TextStyles(self.master.TextStyles)

    @property
    def Theme(self):
        return self.master.Theme

    @property
    def TimeLine(self):
        return TimeLine(self.master.TimeLine)

    @property
    def Width(self):
        return self.master.Width

    def ApplyTheme(self, *args,  `_themeName_` =None):
        arguments = {" `_themeName_` ":  `_themeName_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.master.ApplyTheme(*args, **arguments)

    def Delete(self):
        self.master.Delete()

class MediaBookmark:

    def __init__(self, mediabookmark=None):
        self.mediabookmark = mediabookmark

    @property
    def Index(self):
        return self.mediabookmark.Index

    @property
    def Name(self):
        return self.mediabookmark.Name

    @property
    def Position(self):
        return self.mediabookmark.Position

    def Delete(self):
        return self.mediabookmark.Delete()

class MediaBookmarks:

    def __init__(self, mediabookmarks=None):
        self.mediabookmarks = mediabookmarks

    def __call__(self, item):
        return MediaBookmark(self.mediabookmarks(item))

    @property
    def Count(self):
        return self.mediabookmarks.Count

    def Add(self, *args, Position=None, Name=None):
        arguments = {"Position": Position, "Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return MediaBookmark(self.mediabookmarks.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.mediabookmarks.Item(*args, **arguments)

class MediaFormat:

    def __init__(self, mediaformat=None):
        self.mediaformat = mediaformat

    @property
    def Application(self):
        return Application(self.mediaformat.Application)

    @property
    def AudioCompressionType(self):
        return self.mediaformat.AudioCompressionType

    @property
    def AudioSamplingRate(self):
        return self.mediaformat.AudioSamplingRate

    @property
    def EndPoint(self):
        return self.mediaformat.EndPoint

    @property
    def FadeInDuration(self):
        return self.mediaformat.FadeInDuration

    @property
    def FadeOutDuration(self):
        return self.mediaformat.FadeOutDuration

    @property
    def IsEmbedded(self):
        return self.mediaformat.IsEmbedded

    @property
    def IsLinked(self):
        return self.mediaformat.IsLinked

    @property
    def Length(self):
        return self.mediaformat.Length

    @property
    def MediaBookmarks(self):
        return MediaBookmarks(self.mediaformat.MediaBookmarks)

    @property
    def Muted(self):
        return self.mediaformat.Muted

    @property
    def Parent(self):
        return self.mediaformat.Parent

    @property
    def ResamplingStatus(self):
        return self.mediaformat.ResamplingStatus

    @property
    def SampleHeight(self):
        return self.mediaformat.SampleHeight

    @property
    def SampleWidth(self):
        return self.mediaformat.SampleWidth

    @property
    def StartPoint(self):
        return self.mediaformat.StartPoint

    @property
    def VideoCompressionType(self):
        return self.mediaformat.VideoCompressionType

    @property
    def VideoFrameRate(self):
        return self.mediaformat.VideoFrameRate

    @property
    def Volume(self):
        return self.mediaformat.Volume

    def Resample(self, *args,  `_Trim_`=None, `_SampleHeight_`=None, `_SampleWidth_`=None, `_VideoFrameRate_`=None, `_AudioSamplingRate_`=None, `_VideoBitRate_` =None):
        arguments = {" `_Trim_`":  `_Trim_`, "`_SampleHeight_`": `_SampleHeight_`, "`_SampleWidth_`": `_SampleWidth_`, "`_VideoFrameRate_`": `_VideoFrameRate_`, "`_AudioSamplingRate_`": `_AudioSamplingRate_`, "`_VideoBitRate_` ": `_VideoBitRate_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.mediaformat.Resample(*args, **arguments)

    def ResampleFromProfile(self, *args,  `_profile_` =None):
        arguments = {" `_profile_` ":  `_profile_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.mediaformat.ResampleFromProfile(*args, **arguments)

    def SetDisplayPicture(self, *args,  `_Position_` =None):
        arguments = {" `_Position_` ":  `_Position_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.mediaformat.SetDisplayPicture(*args, **arguments)

    def SetDisplayPictureFromFile(self, *args,  `_FilePath_` =None):
        arguments = {" `_FilePath_` ":  `_FilePath_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.mediaformat.SetDisplayPictureFromFile(*args, **arguments)

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

class MotionEffect:

    def __init__(self, motioneffect=None):
        self.motioneffect = motioneffect

    @property
    def Application(self):
        return Application(self.motioneffect.Application)

    @property
    def ByX(self):
        return self.motioneffect.ByX

    @property
    def ByY(self):
        return self.motioneffect.ByY

    @property
    def FromX(self):
        return self.motioneffect.FromX

    @property
    def FromY(self):
        return MotionEffect(self.motioneffect.FromY)

    @FromY.setter
    def FromY(self, value):
        self.motioneffect.FromY = value

    @property
    def Parent(self):
        return self.motioneffect.Parent

    @property
    def Path(self):
        return self.motioneffect.Path

    @property
    def ToX(self):
        return self.motioneffect.ToX

    @property
    def ToY(self):
        return MotionEffect(self.motioneffect.ToY)

    @ToY.setter
    def ToY(self, value):
        self.motioneffect.ToY = value

class NamedSlideShow:

    def __init__(self, namedslideshow=None):
        self.namedslideshow = namedslideshow

    @property
    def Application(self):
        return Application(self.namedslideshow.Application)

    @property
    def Count(self):
        return self.namedslideshow.Count

    @property
    def Name(self):
        return self.namedslideshow.Name

    @property
    def Parent(self):
        return self.namedslideshow.Parent

    @property
    def SlideIDs(self):
        return self.namedslideshow.SlideIDs

    def Delete(self):
        self.namedslideshow.Delete()

class NamedSlideShows:

    def __init__(self, namedslideshows=None):
        self.namedslideshows = namedslideshows

    def __call__(self, item):
        return NamedSlideShow(self.namedslideshows(item))

    @property
    def Application(self):
        return Application(self.namedslideshows.Application)

    @property
    def Count(self):
        return self.namedslideshows.Count

    @property
    def Parent(self):
        return self.namedslideshows.Parent

    def Add(self, *args, Name=None, SafeArrayOfSlideIDs=None):
        arguments = {"Name": Name, "SafeArrayOfSlideIDs": SafeArrayOfSlideIDs}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return NamedSlideShow(self.namedslideshows.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.namedslideshows.Item(*args, **arguments)

class ObjectVerbs:

    def __init__(self, objectverbs=None):
        self.objectverbs = objectverbs

    @property
    def Application(self):
        return Application(self.objectverbs.Application)

    @property
    def Count(self):
        return self.objectverbs.Count

    @property
    def Parent(self):
        return self.objectverbs.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.objectverbs.Item(*args, **arguments)

class OLEFormat:

    def __init__(self, oleformat=None):
        self.oleformat = oleformat

    @property
    def Application(self):
        return Application(self.oleformat.Application)

    @property
    def FollowColors(self):
        return self.oleformat.FollowColors

    @FollowColors.setter
    def FollowColors(self, value):
        self.oleformat.FollowColors = value

    @property
    def Object(self):
        return self.oleformat.Object

    @property
    def ObjectVerbs(self):
        return ObjectVerbs(self.oleformat.ObjectVerbs)

    @property
    def Parent(self):
        return self.oleformat.Parent

    @property
    def ProgID(self):
        return self.oleformat.ProgID

    def Activate(self):
        self.oleformat.Activate()

    def DoVerb(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.oleformat.DoVerb(*args, **arguments)

class Options:

    def __init__(self, options=None):
        self.options = options

    @property
    def DisplayPasteOptions(self):
        return self.options.DisplayPasteOptions

    @property
    def ShowCoauthoringMergeChanges(self):
        return self.options.ShowCoauthoringMergeChanges

class PageSetup:

    def __init__(self, pagesetup=None):
        self.pagesetup = pagesetup

    @property
    def Application(self):
        return Application(self.pagesetup.Application)

    @property
    def FirstSlideNumber(self):
        return self.pagesetup.FirstSlideNumber

    @FirstSlideNumber.setter
    def FirstSlideNumber(self, value):
        self.pagesetup.FirstSlideNumber = value

    @property
    def NotesOrientation(self):
        return self.pagesetup.NotesOrientation

    @NotesOrientation.setter
    def NotesOrientation(self, value):
        self.pagesetup.NotesOrientation = value

    @property
    def Parent(self):
        return self.pagesetup.Parent

    @property
    def SlideHeight(self):
        return self.pagesetup.SlideHeight

    @SlideHeight.setter
    def SlideHeight(self, value):
        self.pagesetup.SlideHeight = value

    @property
    def SlideOrientation(self):
        return self.pagesetup.SlideOrientation

    @SlideOrientation.setter
    def SlideOrientation(self, value):
        self.pagesetup.SlideOrientation = value

    @property
    def SlideSize(self):
        return self.pagesetup.SlideSize

    @SlideSize.setter
    def SlideSize(self, value):
        self.pagesetup.SlideSize = value

    @property
    def SlideWidth(self):
        return self.pagesetup.SlideWidth

    @SlideWidth.setter
    def SlideWidth(self, value):
        self.pagesetup.SlideWidth = value

class Pane:

    def __init__(self, pane=None):
        self.pane = pane

    @property
    def Active(self):
        return self.pane.Active

    @property
    def Application(self):
        return Application(self.pane.Application)

    @property
    def Parent(self):
        return self.pane.Parent

    @property
    def ViewType(self):
        return self.pane.ViewType

    def Activate(self):
        self.pane.Activate()

class Panes:

    def __init__(self, panes=None):
        self.panes = panes

    def __call__(self, item):
        return Pane(self.panes(item))

    @property
    def Application(self):
        return Application(self.panes.Application)

    @property
    def Count(self):
        return self.panes.Count

    @property
    def Parent(self):
        return self.panes.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.panes.Item(*args, **arguments)

class ParagraphFormat:

    def __init__(self, paragraphformat=None):
        self.paragraphformat = paragraphformat

    @property
    def Alignment(self):
        return self.paragraphformat.Alignment

    @Alignment.setter
    def Alignment(self, value):
        self.paragraphformat.Alignment = value

    @property
    def Application(self):
        return Application(self.paragraphformat.Application)

    @property
    def BaseLineAlignment(self):
        return self.paragraphformat.BaseLineAlignment

    @BaseLineAlignment.setter
    def BaseLineAlignment(self, value):
        self.paragraphformat.BaseLineAlignment = value

    @property
    def Bullet(self):
        return BulletFormat(self.paragraphformat.Bullet)

    @property
    def FarEastLineBreakControl(self):
        return self.paragraphformat.FarEastLineBreakControl

    @FarEastLineBreakControl.setter
    def FarEastLineBreakControl(self, value):
        self.paragraphformat.FarEastLineBreakControl = value

    @property
    def HangingPunctuation(self):
        return self.paragraphformat.HangingPunctuation

    @HangingPunctuation.setter
    def HangingPunctuation(self, value):
        self.paragraphformat.HangingPunctuation = value

    @property
    def LineRuleAfter(self):
        return self.paragraphformat.LineRuleAfter

    @property
    def LineRuleBefore(self):
        return self.paragraphformat.LineRuleBefore

    @property
    def LineRuleWithin(self):
        return self.paragraphformat.LineRuleWithin

    @property
    def Parent(self):
        return self.paragraphformat.Parent

    @property
    def SpaceAfter(self):
        return self.paragraphformat.SpaceAfter

    @SpaceAfter.setter
    def SpaceAfter(self, value):
        self.paragraphformat.SpaceAfter = value

    @property
    def SpaceBefore(self):
        return self.paragraphformat.SpaceBefore

    @SpaceBefore.setter
    def SpaceBefore(self, value):
        self.paragraphformat.SpaceBefore = value

    @property
    def SpaceWithin(self):
        return self.paragraphformat.SpaceWithin

    @SpaceWithin.setter
    def SpaceWithin(self, value):
        self.paragraphformat.SpaceWithin = value

    @property
    def TextDirection(self):
        return self.paragraphformat.TextDirection

    @TextDirection.setter
    def TextDirection(self, value):
        self.paragraphformat.TextDirection = value

    @property
    def WordWrap(self):
        return self.paragraphformat.WordWrap

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

class PlaceholderFormat:

    def __init__(self, placeholderformat=None):
        self.placeholderformat = placeholderformat

    @property
    def Application(self):
        return Application(self.placeholderformat.Application)

    @property
    def ContainedType(self):
        return self.placeholderformat.ContainedType

    @property
    def Name(self):
        return self.placeholderformat.Name

    @Name.setter
    def Name(self, value):
        self.placeholderformat.Name = value

    @property
    def Parent(self):
        return self.placeholderformat.Parent

    @property
    def Type(self):
        return self.placeholderformat.Type

class Placeholders:

    def __init__(self, placeholders=None):
        self.placeholders = placeholders

    def __call__(self, item):
        return Placeholder(self.placeholders(item))

    @property
    def Application(self):
        return Application(self.placeholders.Application)

    @property
    def Count(self):
        return self.placeholders.Count

    @property
    def Parent(self):
        return self.placeholders.Parent

    def FindByName(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.placeholders.FindByName(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.placeholders.Item(*args, **arguments)

class Player:

    def __init__(self, player=None):
        self.player = player

    @property
    def Application(self):
        return Application(self.player.Application)

    @property
    def CurrentPosition(self):
        return self.player.CurrentPosition

    @property
    def Parent(self):
        return self.player.Parent

    @property
    def State(self):
        return self.player.State

    def GoToNextBookmark(self):
        self.player.GoToNextBookmark()

    def GoToPreviousBookmark(self):
        self.player.GoToPreviousBookmark()

    def Pause(self):
        self.player.Pause()

    def Stop(self):
        self.player.Stop()

class PlaySettings:

    def __init__(self, playsettings=None):
        self.playsettings = playsettings

    @property
    def ActionVerb(self):
        return self.playsettings.ActionVerb

    @ActionVerb.setter
    def ActionVerb(self, value):
        self.playsettings.ActionVerb = value

    @property
    def Application(self):
        return Application(self.playsettings.Application)

    @property
    def HideWhileNotPlaying(self):
        return self.playsettings.HideWhileNotPlaying

    @property
    def LoopUntilStopped(self):
        return self.playsettings.LoopUntilStopped

    @property
    def Parent(self):
        return self.playsettings.Parent

    @property
    def PauseAnimation(self):
        return self.playsettings.PauseAnimation

    @property
    def PlayOnEntry(self):
        return self.playsettings.PlayOnEntry

    @property
    def RewindMovie(self):
        return self.playsettings.RewindMovie

    @property
    def StopAfterSlides(self):
        return self.playsettings.StopAfterSlides

    @StopAfterSlides.setter
    def StopAfterSlides(self, value):
        self.playsettings.StopAfterSlides = value

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
        return XlColorIndex(self.point.MarkerBackgroundColorIndex)

    @MarkerBackgroundColorIndex.setter
    def MarkerBackgroundColorIndex(self, value):
        self.point.MarkerBackgroundColorIndex = value

    @property
    def MarkerForegroundColor(self):
        return self.point.MarkerForegroundColor

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

class Presentation:

    def __init__(self, presentation=None):
        self.presentation = presentation

    @property
    def Application(self):
        return Application(self.presentation.Application)

    @property
    def AutoSaveOn(self):
        return self.presentation.AutoSaveOn

    @property
    def Broadcast(self):
        return Broadcast(self.presentation.Broadcast)

    @property
    def BuiltInDocumentProperties(self):
        return self.presentation.BuiltInDocumentProperties

    @property
    def Coauthoring(self):
        return Coauthoring(self.presentation.Coauthoring)

    @property
    def ColorSchemes(self):
        return ColorSchemes(self.presentation.ColorSchemes)

    @property
    def CommandBars(self):
        return self.presentation.CommandBars

    @property
    def Container(self):
        return self.presentation.Container

    @property
    def ContentTypeProperties(self):
        return self.presentation.ContentTypeProperties

    @property
    def CreateVideoStatus(self):
        return Presentation(self.presentation.CreateVideoStatus)

    @property
    def CustomDocumentProperties(self):
        return self.presentation.CustomDocumentProperties

    @property
    def CustomerData(self):
        return CustomerData(self.presentation.CustomerData)

    @property
    def CustomXMLParts(self):
        return self.presentation.CustomXMLParts

    @property
    def DefaultLanguageID(self):
        return self.presentation.DefaultLanguageID

    @DefaultLanguageID.setter
    def DefaultLanguageID(self, value):
        self.presentation.DefaultLanguageID = value

    @property
    def DefaultShape(self):
        return Shape(self.presentation.DefaultShape)

    @property
    def Designs(self):
        return Designs(self.presentation.Designs)

    @property
    def DisplayComments(self):
        return self.presentation.DisplayComments

    @property
    def DocumentInspectors(self):
        return self.presentation.DocumentInspectors

    @property
    def DocumentLibraryVersions(self):
        return self.presentation.DocumentLibraryVersions

    @property
    def EncryptionProvider(self):
        return self.presentation.EncryptionProvider

    @property
    def EnvelopeVisible(self):
        return self.presentation.EnvelopeVisible

    @property
    def ExtraColors(self):
        return ExtraColors(self.presentation.ExtraColors)

    @property
    def FarEastLineBreakLanguage(self):
        return self.presentation.FarEastLineBreakLanguage

    @FarEastLineBreakLanguage.setter
    def FarEastLineBreakLanguage(self, value):
        self.presentation.FarEastLineBreakLanguage = value

    @property
    def FarEastLineBreakLevel(self):
        return self.presentation.FarEastLineBreakLevel

    @FarEastLineBreakLevel.setter
    def FarEastLineBreakLevel(self, value):
        self.presentation.FarEastLineBreakLevel = value

    @property
    def Final(self):
        return self.presentation.Final

    @property
    def Fonts(self):
        return Fonts(self.presentation.Fonts)

    @property
    def FullName(self):
        return self.presentation.FullName

    @property
    def GridDistance(self):
        return self.presentation.GridDistance

    @property
    def HandoutMaster(self):
        return Master(self.presentation.HandoutMaster)

    @property
    def HasHandoutMaster(self):
        return self.presentation.HasHandoutMaster

    @property
    def HasNotesMaster(self):
        return self.presentation.HasNotesMaster

    @property
    def HasTitleMaster(self):
        return self.presentation.HasTitleMaster

    @property
    def HasVBProject(self):
        return self.presentation.HasVBProject

    @property
    def InMergeMode(self):
        return self.presentation.InMergeMode

    @property
    def IsFullyDownloaded(self):
        return self.presentation.IsFullyDownloaded

    @property
    def LayoutDirection(self):
        return self.presentation.LayoutDirection

    @LayoutDirection.setter
    def LayoutDirection(self, value):
        self.presentation.LayoutDirection = value

    @property
    def Name(self):
        return self.presentation.Name

    @property
    def NoLineBreakAfter(self):
        return self.presentation.NoLineBreakAfter

    @NoLineBreakAfter.setter
    def NoLineBreakAfter(self, value):
        self.presentation.NoLineBreakAfter = value

    @property
    def NoLineBreakBefore(self):
        return self.presentation.NoLineBreakBefore

    @NoLineBreakBefore.setter
    def NoLineBreakBefore(self, value):
        self.presentation.NoLineBreakBefore = value

    @property
    def NotesMaster(self):
        return Master(self.presentation.NotesMaster)

    @property
    def PageSetup(self):
        return PageSetup(self.presentation.PageSetup)

    @property
    def Parent(self):
        return self.presentation.Parent

    @property
    def Password(self):
        return self.presentation.Password

    @Password.setter
    def Password(self, value):
        self.presentation.Password = value

    @property
    def PasswordEncryptionAlgorithm(self):
        return self.presentation.PasswordEncryptionAlgorithm

    @property
    def PasswordEncryptionFileProperties(self):
        return self.presentation.PasswordEncryptionFileProperties

    @property
    def PasswordEncryptionKeyLength(self):
        return self.presentation.PasswordEncryptionKeyLength

    @property
    def PasswordEncryptionProvider(self):
        return self.presentation.PasswordEncryptionProvider

    @property
    def Path(self):
        return Presentation(self.presentation.Path)

    @property
    def PrintOptions(self):
        return PrintOptions(self.presentation.PrintOptions)

    @property
    def ReadOnly(self):
        return self.presentation.ReadOnly

    @property
    def ReadOnlyRecommended(self):
        return self.presentation.ReadOnlyRecommended

    @property
    def RemovePersonalInformation(self):
        return self.presentation.RemovePersonalInformation

    @property
    def Research(self):
        return Research(self.presentation.Research)

    @property
    def Saved(self):
        return self.presentation.Saved

    @property
    def SectionProperties(self):
        return SectionProperties(self.presentation.SectionProperties)

    @property
    def SensitivityLabel(self):
        return self.presentation.SensitivityLabel

    @property
    def ServerPolicy(self):
        return self.presentation.ServerPolicy

    @property
    def SharedWorkspace(self):
        return self.presentation.SharedWorkspace

    @property
    def Signatures(self):
        return self.presentation.Signatures

    @property
    def SlideMaster(self):
        return Master(self.presentation.SlideMaster)

    @property
    def Slides(self):
        return Slides(self.presentation.Slides)

    @property
    def SlideShowSettings(self):
        return SlideShowSettings(self.presentation.SlideShowSettings)

    @property
    def SlideShowWindow(self):
        return SlideShowWindow(self.presentation.SlideShowWindow)

    @property
    def SnapToGrid(self):
        return self.presentation.SnapToGrid

    @property
    def Sync(self):
        return self.presentation.Sync

    @property
    def Tags(self):
        return Tags(self.presentation.Tags)

    @property
    def TemplateName(self):
        return self.presentation.TemplateName

    @property
    def TitleMaster(self):
        return Master(self.presentation.TitleMaster)

    @property
    def VBASigned(self):
        return self.presentation.VBASigned

    @property
    def VBProject(self):
        return self.presentation.VBProject

    @property
    def Windows(self):
        return DocumentWindows(self.presentation.Windows)

    @property
    def WritePassword(self):
        return self.presentation.WritePassword

    def AcceptAll(self):
        return self.presentation.AcceptAll()

    def AddTitleMaster(self):
        return self.presentation.AddTitleMaster()

    def AddToFavorites(self):
        self.presentation.AddToFavorites()

    def ApplyTemplate(self, *args,  `_FileName_` =None):
        arguments = {" `_FileName_` ":  `_FileName_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.presentation.ApplyTemplate(*args, **arguments)

    def ApplyTheme(self, *args,  `_themeName_` =None):
        arguments = {" `_themeName_` ":  `_themeName_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.presentation.ApplyTheme(*args, **arguments)

    def CanCheckIn(self):
        return self.presentation.CanCheckIn()

    def CheckIn(self, *args,  `_SaveChanges_`=None, `_Comments_`=None, `_MakePublic_` =None):
        arguments = {" `_SaveChanges_`":  `_SaveChanges_`, "`_Comments_`": `_Comments_`, "`_MakePublic_` ": `_MakePublic_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.presentation.CheckIn(*args, **arguments)

    def CheckInWithVersion(self, *args,  `_SaveChanges_`=None, `_Comments_`=None, `_MakePublic_`=None, `_VersionType_` =None):
        arguments = {" `_SaveChanges_`":  `_SaveChanges_`, "`_Comments_`": `_Comments_`, "`_MakePublic_`": `_MakePublic_`, "`_VersionType_` ": `_VersionType_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.presentation.CheckInWithVersion(*args, **arguments)

    def Close(self):
        self.presentation.Close()

    def Convert2(self, *args,  `_FileName_` =None):
        arguments = {" `_FileName_` ":  `_FileName_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.presentation.Convert2(*args, **arguments)

    def CreateVideo(self, *args,  `_FileName_`=None, `_UseTimingsAndNarrations_`=None, `_DefaultSlideDuration_`=None, `_VertResolution_`=None, `_FramesPerSecond_`=None, `_Quality_` =None):
        arguments = {" `_FileName_`":  `_FileName_`, "`_UseTimingsAndNarrations_`": `_UseTimingsAndNarrations_`, "`_DefaultSlideDuration_`": `_DefaultSlideDuration_`, "`_VertResolution_`": `_VertResolution_`, "`_FramesPerSecond_`": `_FramesPerSecond_`, "`_Quality_` ": `_Quality_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.presentation.CreateVideo(*args, **arguments)

    def EndReview(self):
        return self.presentation.EndReview()

    def EnsureAllMediaUpgraded(self):
        self.presentation.EnsureAllMediaUpgraded()

    def Export(self, *args, Path=None, FilterName=None, ScaleWidth=None, ScaleHeight=None):
        arguments = {"Path": Path, "FilterName": FilterName, "ScaleWidth": ScaleWidth, "ScaleHeight": ScaleHeight}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.presentation.Export(*args, **arguments)

    def ExportAsFixedFormat(self, *args, Path=None, FixedFormatType=None, Intent=None, FrameSlides=None, HandoutOrder=None, OutputType=None, PrintHiddenSlides=None, PrintRange=None, RangeType=None, SlideShowName=None, IncludeDocProperties=None, KeepIRMSettings=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005\_1=None, ExternalExporter=None):
        arguments = {"Path": Path, "FixedFormatType": FixedFormatType, "Intent": Intent, "FrameSlides": FrameSlides, "HandoutOrder": HandoutOrder, "OutputType": OutputType, "PrintHiddenSlides": PrintHiddenSlides, "PrintRange": PrintRange, "RangeType": RangeType, "SlideShowName": SlideShowName, "IncludeDocProperties": IncludeDocProperties, "KeepIRMSettings": KeepIRMSettings, "DocStructureTags": DocStructureTags, "BitmapMissingFonts": BitmapMissingFonts, "UseISO19005\_1": UseISO19005\_1, "ExternalExporter": ExternalExporter}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.presentation.ExportAsFixedFormat(*args, **arguments)

    def FollowHyperlink(self, *args,  `_Address_`=None, `_SubAddress_`=None, `_NewWindow_`=None, `_AddHistory_`=None, `_ExtraInfo_`=None, `_Method_`=None, `_HeaderInfo_` =None):
        arguments = {" `_Address_`":  `_Address_`, "`_SubAddress_`": `_SubAddress_`, "`_NewWindow_`": `_NewWindow_`, "`_AddHistory_`": `_AddHistory_`, "`_ExtraInfo_`": `_ExtraInfo_`, "`_Method_`": `_Method_`, "`_HeaderInfo_` ": `_HeaderInfo_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.presentation.FollowHyperlink(*args, **arguments)

    def GetWorkflowTasks(self):
        return self.presentation.GetWorkflowTasks()

    def GetWorkflowTemplates(self):
        return self.presentation.GetWorkflowTemplates()

    def LockServerFile(self):
        self.presentation.LockServerFile()

    def MergeWithBaseline(self, *args,  `_withPresentation_`=None, `_baselinePresentation_` =None):
        arguments = {" `_withPresentation_`":  `_withPresentation_`, "`_baselinePresentation_` ": `_baselinePresentation_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.presentation.MergeWithBaseline(*args, **arguments)

    def NewWindow(self):
        return self.presentation.NewWindow()

    def PrintOut(self, *args, From=None, To=None, PrintToFile=None, Copies=None, Collate=None):
        arguments = {"From": From, "To": To, "PrintToFile": PrintToFile, "Copies": Copies, "Collate": Collate}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.presentation.PrintOut(*args, **arguments)

    def PublishSlides(self, *args,  `_SlideLibraryUrl_`=None, `_Overwrite_` =None):
        arguments = {" `_SlideLibraryUrl_`":  `_SlideLibraryUrl_`, "`_Overwrite_` ": `_Overwrite_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.presentation.PublishSlides(*args, **arguments)

    def RejectAll(self):
        return self.presentation.RejectAll()

    def RemoveDocumentInformation(self, *args,  `_Type_` =None):
        arguments = {" `_Type_` ":  `_Type_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.presentation.RemoveDocumentInformation(*args, **arguments)

    def Save(self):
        self.presentation.Save()

    def SaveAs(self, *args, FileName=None, FileFormat=None, EmbedFonts=None):
        arguments = {"FileName": FileName, "FileFormat": FileFormat, "EmbedFonts": EmbedFonts}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.presentation.SaveAs(*args, **arguments)

    def SaveCopyAs(self, *args,  `_FileName_`=None, `_FileFormat_`=None, `_EmbedTrueTypeFonts_` =None):
        arguments = {" `_FileName_`":  `_FileName_`, "`_FileFormat_`": `_FileFormat_`, "`_EmbedTrueTypeFonts_` ": `_EmbedTrueTypeFonts_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.presentation.SaveCopyAs(*args, **arguments)

    def SaveCopyAs2(self, *args,  `_FileName_`=None, `_FileFormat_`=None, `_EmbedTrueTypeFonts_`=None, `_ReadOnlyRecommended_` =None):
        arguments = {" `_FileName_`":  `_FileName_`, "`_FileFormat_`": `_FileFormat_`, "`_EmbedTrueTypeFonts_`": `_EmbedTrueTypeFonts_`, "`_ReadOnlyRecommended_` ": `_ReadOnlyRecommended_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.presentation.SaveCopyAs2(*args, **arguments)

    def SendFaxOverInternet(self, *args,  `_Recipients_`=None, `_Subject_`=None, `_ShowMessage_` =None):
        arguments = {" `_Recipients_`":  `_Recipients_`, "`_Subject_`": `_Subject_`, "`_ShowMessage_` ": `_ShowMessage_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.presentation.SendFaxOverInternet(*args, **arguments)

    def SetPasswordEncryptionOptions(self, *args,  `_PasswordEncryptionProvider_`=None, `_PasswordEncryptionAlgorithm_`=None, `_PasswordEncryptionKeyLength_`=None, `_PasswordEncryptionFileProperties_` =None):
        arguments = {" `_PasswordEncryptionProvider_`":  `_PasswordEncryptionProvider_`, "`_PasswordEncryptionAlgorithm_`": `_PasswordEncryptionAlgorithm_`, "`_PasswordEncryptionKeyLength_`": `_PasswordEncryptionKeyLength_`, "`_PasswordEncryptionFileProperties_` ": `_PasswordEncryptionFileProperties_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.presentation.SetPasswordEncryptionOptions(*args, **arguments)

    def UpdateLinks(self):
        self.presentation.UpdateLinks()

class Presentations:

    def __init__(self, presentations=None):
        self.presentations = presentations

    def __call__(self, item):
        return Presentation(self.presentations(item))

    @property
    def Application(self):
        return Application(self.presentations.Application)

    @property
    def Count(self):
        return self.presentations.Count

    @property
    def Parent(self):
        return self.presentations.Parent

    def Add(self, *args, WithWindow=None):
        arguments = {"WithWindow": WithWindow}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Presentation(self.presentations.Add(*args, **arguments))

    def CanCheckOut(self, *args,  `_FileName_` =None):
        arguments = {" `_FileName_` ":  `_FileName_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.presentations.CanCheckOut(*args, **arguments)

    def CheckOut(self, *args, FileName=None):
        arguments = {"FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.presentations.CheckOut(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.presentations.Item(*args, **arguments)

    def Open(self, *args, FileName=None, ReadOnly=None, Untitled=None, WithWindow=None):
        arguments = {"FileName": FileName, "ReadOnly": ReadOnly, "Untitled": Untitled, "WithWindow": WithWindow}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Presentation(self.presentations.Open(*args, **arguments))

    def Open2007(self, *args,  `_FileName_`=None, `_ReadOnly_`=None, `_Untitled_`=None, `_WithWindow_`=None, `_OpenAndRepair_` =None):
        arguments = {" `_FileName_`":  `_FileName_`, "`_ReadOnly_`": `_ReadOnly_`, "`_Untitled_`": `_Untitled_`, "`_WithWindow_`": `_WithWindow_`, "`_OpenAndRepair_` ": `_OpenAndRepair_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Presentation(self.presentations.Open2007(*args, **arguments))

class PrintOptions:

    def __init__(self, printoptions=None):
        self.printoptions = printoptions

    @property
    def ActivePrinter(self):
        return self.printoptions.ActivePrinter

    @property
    def Application(self):
        return Application(self.printoptions.Application)

    @property
    def Collate(self):
        return self.printoptions.Collate

    @property
    def FitToPage(self):
        return self.printoptions.FitToPage

    @property
    def FrameSlides(self):
        return self.printoptions.FrameSlides

    @property
    def HandoutOrder(self):
        return self.printoptions.HandoutOrder

    @HandoutOrder.setter
    def HandoutOrder(self, value):
        self.printoptions.HandoutOrder = value

    @property
    def HighQuality(self):
        return self.printoptions.HighQuality

    @property
    def NumberOfCopies(self):
        return self.printoptions.NumberOfCopies

    @NumberOfCopies.setter
    def NumberOfCopies(self, value):
        self.printoptions.NumberOfCopies = value

    @property
    def OutputType(self):
        return self.printoptions.OutputType

    @OutputType.setter
    def OutputType(self, value):
        self.printoptions.OutputType = value

    @property
    def Parent(self):
        return self.printoptions.Parent

    @property
    def PrintColorType(self):
        return self.printoptions.PrintColorType

    @PrintColorType.setter
    def PrintColorType(self, value):
        self.printoptions.PrintColorType = value

    @property
    def PrintComments(self):
        return self.printoptions.PrintComments

    @property
    def PrintFontsAsGraphics(self):
        return self.printoptions.PrintFontsAsGraphics

    @property
    def PrintHiddenSlides(self):
        return self.printoptions.PrintHiddenSlides

    @property
    def PrintInBackground(self):
        return self.printoptions.PrintInBackground

    @property
    def Ranges(self):
        return PrintRanges(self.printoptions.Ranges)

    @property
    def RangeType(self):
        return self.printoptions.RangeType

    @RangeType.setter
    def RangeType(self, value):
        self.printoptions.RangeType = value

    @property
    def sectionIndex(self):
        return PrintOptions(self.printoptions.sectionIndex)

    @property
    def SlideShowName(self):
        return self.printoptions.SlideShowName

    @SlideShowName.setter
    def SlideShowName(self, value):
        self.printoptions.SlideShowName = value

class PrintRange:

    def __init__(self, printrange=None):
        self.printrange = printrange

    @property
    def Application(self):
        return Application(self.printrange.Application)

    @property
    def End(self):
        return self.printrange.End

    @property
    def Parent(self):
        return self.printrange.Parent

    @property
    def Start(self):
        return self.printrange.Start

    def Delete(self):
        self.printrange.Delete()

class PrintRanges:

    def __init__(self, printranges=None):
        self.printranges = printranges

    def __call__(self, item):
        return PrintRange(self.printranges(item))

    @property
    def Application(self):
        return Application(self.printranges.Application)

    @property
    def Count(self):
        return self.printranges.Count

    @property
    def Parent(self):
        return self.printranges.Parent

    def Add(self, *args, Start=None, End=None):
        arguments = {"Start": Start, "End": End}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return PrintRange(self.printranges.Add(*args, **arguments))

    def ClearAll(self):
        return self.printranges.ClearAll()

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.printranges.Item(*args, **arguments)

class PropertyEffect:

    def __init__(self, propertyeffect=None):
        self.propertyeffect = propertyeffect

    @property
    def Application(self):
        return Application(self.propertyeffect.Application)

    @property
    def From(self):
        return self.propertyeffect.From

    @property
    def Parent(self):
        return self.propertyeffect.Parent

    @property
    def Points(self):
        return AnimationPoints(self.propertyeffect.Points)

    @property
    def Property(self):
        return self.propertyeffect.Property

    @property
    def To(self):
        return self.propertyeffect.To

class ProtectedViewWindow:

    def __init__(self, protectedviewwindow=None):
        self.protectedviewwindow = protectedviewwindow

    @property
    def Active(self):
        return self.protectedviewwindow.Active

    @property
    def Application(self):
        return Application(self.protectedviewwindow.Application)

    @property
    def Caption(self):
        return self.protectedviewwindow.Caption

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
    def Parent(self):
        return self.protectedviewwindow.Parent

    @property
    def Presentation(self):
        return Presentation(self.protectedviewwindow.Presentation)

    @property
    def SourceName(self):
        return ProtectedViewWindow(self.protectedviewwindow.SourceName)

    @property
    def SourcePath(self):
        return ProtectedViewWindow(self.protectedviewwindow.SourcePath)

    @property
    def Top(self):
        return self.protectedviewwindow.Top

    @Top.setter
    def Top(self, value):
        self.protectedviewwindow.Top = value

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
        self.protectedviewwindow.Activate()

    def Close(self):
        self.protectedviewwindow.Close()

    def Edit(self, *args, ModifyPassword=None):
        arguments = {"ModifyPassword": ModifyPassword}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.protectedviewwindow.Edit(*args, **arguments)

class ProtectedViewWindows:

    def __init__(self, protectedviewwindows=None):
        self.protectedviewwindows = protectedviewwindows

    @property
    def Application(self):
        return Application(self.protectedviewwindows.Application)

    @property
    def Count(self):
        return self.protectedviewwindows.Count

    @property
    def Parent(self):
        return self.protectedviewwindows.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.protectedviewwindows.Item(*args, **arguments)

    def Open(self, *args, FileName=None, ReadPassword=None, OpenAndRepair=None):
        arguments = {"FileName": FileName, "ReadPassword": ReadPassword, "OpenAndRepair": OpenAndRepair}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.protectedviewwindows.Open(*args, **arguments)

class PublishObject:

    def __init__(self, publishobject=None):
        self.publishobject = publishobject

    @property
    def Application(self):
        return Application(self.publishobject.Application)

    @property
    def FileName(self):
        return self.publishobject.FileName

    @FileName.setter
    def FileName(self, value):
        self.publishobject.FileName = value

    @property
    def HTMLVersion(self):
        return self.publishobject.HTMLVersion

    @HTMLVersion.setter
    def HTMLVersion(self, value):
        self.publishobject.HTMLVersion = value

    @property
    def Parent(self):
        return self.publishobject.Parent

    @property
    def RangeEnd(self):
        return self.publishobject.RangeEnd

    @RangeEnd.setter
    def RangeEnd(self, value):
        self.publishobject.RangeEnd = value

    @property
    def RangeStart(self):
        return self.publishobject.RangeStart

    @RangeStart.setter
    def RangeStart(self, value):
        self.publishobject.RangeStart = value

    @property
    def SlideShowName(self):
        return self.publishobject.SlideShowName

    @SlideShowName.setter
    def SlideShowName(self, value):
        self.publishobject.SlideShowName = value

    @property
    def SourceType(self):
        return self.publishobject.SourceType

    @SourceType.setter
    def SourceType(self, value):
        self.publishobject.SourceType = value

    @property
    def SpeakerNotes(self):
        return self.publishobject.SpeakerNotes

    def Publish(self):
        self.publishobject.Publish()

class PublishObjects:

    def __init__(self, publishobjects=None):
        self.publishobjects = publishobjects

    def __call__(self, item):
        return PublishObject(self.publishobjects(item))

    @property
    def Application(self):
        return Application(self.publishobjects.Application)

    @property
    def Count(self):
        return self.publishobjects.Count

    @property
    def Parent(self):
        return self.publishobjects.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.publishobjects.Item(*args, **arguments)

class ResampleMediaTasks:

    def __init__(self, resamplemediatasks=None):
        self.resamplemediatasks = resamplemediatasks

    def __call__(self, item):
        return ResampleMediaTask(self.resamplemediatasks(item))

    @property
    def Count(self):
        return self.resamplemediatasks.Count

    @property
    def PercentComplete(self):
        return self.resamplemediatasks.PercentComplete

    def Cancel(self):
        return self.resamplemediatasks.Cancel()

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.resamplemediatasks.Item(*args, **arguments)

    def Pause(self):
        self.resamplemediatasks.Pause()

    def Resume(self):
        return self.resamplemediatasks.Resume()

class Research:

    def __init__(self, research=None):
        self.research = research

    @property
    def Application(self):
        return Application(self.research.Application)

    @property
    def Parent(self):
        return self.research.Parent

    def IsResearchService(self, *args,  `_ServiceID_` =None):
        arguments = {" `_ServiceID_` ":  `_ServiceID_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.research.IsResearchService(*args, **arguments)

    def Query(self, *args,  `_ServiceID_`=None, `_QueryString_`=None, `_QueryLanguage_`=None, `_UseSelection_`=None, `_RequeryContextXML_`=None, `_NewQueryContextXML_`=None, `_LaunchQuery_` =None):
        arguments = {" `_ServiceID_`":  `_ServiceID_`, "`_QueryString_`": `_QueryString_`, "`_QueryLanguage_`": `_QueryLanguage_`, "`_UseSelection_`": `_UseSelection_`, "`_RequeryContextXML_`": `_RequeryContextXML_`, "`_NewQueryContextXML_`": `_NewQueryContextXML_`, "`_LaunchQuery_` ": `_LaunchQuery_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.research.Query(*args, **arguments)

    def SetLanguagePair(self, *args, Language1=None, Language2=None):
        arguments = {"Language1": Language1, "Language2": Language2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.research.SetLanguagePair(*args, **arguments)

class RGBColor:

    def __init__(self, rgbcolor=None):
        self.rgbcolor = rgbcolor

    @property
    def Application(self):
        return Application(self.rgbcolor.Application)

    @property
    def Parent(self):
        return self.rgbcolor.Parent

    @property
    def RGB(self):
        return PpColorSchemeIndex(self.rgbcolor.RGB)

    @RGB.setter
    def RGB(self, value):
        self.rgbcolor.RGB = value

class RotationEffect:

    def __init__(self, rotationeffect=None):
        self.rotationeffect = rotationeffect

    @property
    def Application(self):
        return Application(self.rotationeffect.Application)

    @property
    def By(self):
        return self.rotationeffect.By

    @property
    def From(self):
        return self.rotationeffect.From

    @property
    def Parent(self):
        return self.rotationeffect.Parent

    @property
    def To(self):
        return self.rotationeffect.To

class Row:

    def __init__(self, row=None):
        self.row = row

    @property
    def Application(self):
        return Application(self.row.Application)

    def Cells(self, *args, RowIndex=None, ColumnIndex=None):
        arguments = {"RowIndex": RowIndex, "ColumnIndex": ColumnIndex}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return CellRange(self.row.Cells(*args, **arguments))

    @property
    def Height(self):
        return self.row.Height

    @Height.setter
    def Height(self, value):
        self.row.Height = value

    @property
    def Parent(self):
        return self.row.Parent

    def Delete(self):
        self.row.Delete()

    def Select(self):
        self.row.Select()

class Rows:

    def __init__(self, rows=None):
        self.rows = rows

    def __call__(self, item):
        return Row(self.rows(item))

    @property
    def Application(self):
        return Application(self.rows.Application)

    @property
    def Count(self):
        return self.rows.Count

    @property
    def Parent(self):
        return self.rows.Parent

    def Add(self, *args, BeforeRow=None):
        arguments = {"BeforeRow": BeforeRow}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Row(self.rows.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.rows.Item(*args, **arguments)

class Ruler:

    def __init__(self, ruler=None):
        self.ruler = ruler

    @property
    def Application(self):
        return Application(self.ruler.Application)

    @property
    def Levels(self):
        return RulerLevels(self.ruler.Levels)

    @property
    def Parent(self):
        return self.ruler.Parent

    @property
    def TabStops(self):
        return TabStops(self.ruler.TabStops)

class RulerLevel:

    def __init__(self, rulerlevel=None):
        self.rulerlevel = rulerlevel

    @property
    def Application(self):
        return Application(self.rulerlevel.Application)

    @property
    def FirstMargin(self):
        return self.rulerlevel.FirstMargin

    @FirstMargin.setter
    def FirstMargin(self, value):
        self.rulerlevel.FirstMargin = value

    @property
    def LeftMargin(self):
        return self.rulerlevel.LeftMargin

    @LeftMargin.setter
    def LeftMargin(self, value):
        self.rulerlevel.LeftMargin = value

    @property
    def Parent(self):
        return self.rulerlevel.Parent

class RulerLevels:

    def __init__(self, rulerlevels=None):
        self.rulerlevels = rulerlevels

    def __call__(self, item):
        return RulerLevel(self.rulerlevels(item))

    @property
    def Application(self):
        return Application(self.rulerlevels.Application)

    @property
    def Count(self):
        return self.rulerlevels.Count

    @property
    def Parent(self):
        return self.rulerlevels.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.rulerlevels.Item(*args, **arguments)

class ScaleEffect:

    def __init__(self, scaleeffect=None):
        self.scaleeffect = scaleeffect

    @property
    def Application(self):
        return Application(self.scaleeffect.Application)

    @property
    def ByX(self):
        return self.scaleeffect.ByX

    @property
    def ByY(self):
        return self.scaleeffect.ByY

    @property
    def FromX(self):
        return self.scaleeffect.FromX

    @property
    def FromY(self):
        return ScaleEffect(self.scaleeffect.FromY)

    @FromY.setter
    def FromY(self, value):
        self.scaleeffect.FromY = value

    @property
    def Parent(self):
        return self.scaleeffect.Parent

    @property
    def ToX(self):
        return self.scaleeffect.ToX

    @property
    def ToY(self):
        return ScaleEffect(self.scaleeffect.ToY)

    @ToY.setter
    def ToY(self, value):
        self.scaleeffect.ToY = value

class SectionProperties:

    def __init__(self, sectionproperties=None):
        self.sectionproperties = sectionproperties

    @property
    def Application(self):
        return Application(self.sectionproperties.Application)

    @property
    def Count(self):
        return self.sectionproperties.Count

    @property
    def Parent(self):
        return self.sectionproperties.Parent

    def AddBeforeSlide(self, *args,  `_SlideIndex_`=None, `_sectionName_` =None):
        arguments = {" `_SlideIndex_`":  `_SlideIndex_`, "`_sectionName_` ": `_sectionName_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sectionproperties.AddBeforeSlide(*args, **arguments)

    def AddSection(self, *args,  `_sectionIndex_`=None, `_sectionName_` =None):
        arguments = {" `_sectionIndex_`":  `_sectionIndex_`, "`_sectionName_` ": `_sectionName_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sectionproperties.AddSection(*args, **arguments)

    def Delete(self, *args,  `_sectionIndex_`=None, `_deleteSlides_` =None):
        arguments = {" `_sectionIndex_`":  `_sectionIndex_`, "`_deleteSlides_` ": `_deleteSlides_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sectionproperties.Delete(*args, **arguments)

    def FirstSlide(self, *args,  `_sectionIndex_` =None):
        arguments = {" `_sectionIndex_` ":  `_sectionIndex_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sectionproperties.FirstSlide(*args, **arguments)

    def Move(self, *args,  `_sectionIndex_`=None, `_toPos_` =None):
        arguments = {" `_sectionIndex_`":  `_sectionIndex_`, "`_toPos_` ": `_toPos_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sectionproperties.Move(*args, **arguments)

    def Name(self, *args,  `_sectionIndex_` =None):
        arguments = {" `_sectionIndex_` ":  `_sectionIndex_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sectionproperties.Name(*args, **arguments)

    def Rename(self, *args,  `_sectionIndex_`=None, `_sectionName_` =None):
        arguments = {" `_sectionIndex_`":  `_sectionIndex_`, "`_sectionName_` ": `_sectionName_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sectionproperties.Rename(*args, **arguments)

    def SectionID(self, *args,  `_sectionIndex_` =None):
        arguments = {" `_sectionIndex_` ":  `_sectionIndex_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sectionproperties.SectionID(*args, **arguments)

    def SlidesCount(self, *args,  `_sectionIndex_` =None):
        arguments = {" `_sectionIndex_` ":  `_sectionIndex_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sectionproperties.SlidesCount(*args, **arguments)

class Selection:

    def __init__(self, selection=None):
        self.selection = selection

    @property
    def Application(self):
        return Application(self.selection.Application)

    @property
    def ChildShapeRange(self):
        return ShapeRange(self.selection.ChildShapeRange)

    @property
    def HasChildShapeRange(self):
        return self.selection.HasChildShapeRange

    @property
    def Parent(self):
        return self.selection.Parent

    @property
    def ShapeRange(self):
        return ShapeRange(self.selection.ShapeRange)

    @property
    def SlideRange(self):
        return SlideRange(self.selection.SlideRange)

    @property
    def TextRange(self):
        return TextRange(self.selection.TextRange)

    @property
    def TextRange2(self):
        return self.selection.TextRange2

    @property
    def Type(self):
        return self.selection.Type

    def Copy(self):
        self.selection.Copy()

    def Cut(self):
        self.selection.Cut()

    def Delete(self):
        self.selection.Delete()

    def Unselect(self):
        self.selection.Unselect()

class Sequence:

    def __init__(self, sequence=None):
        self.sequence = sequence

    @property
    def Application(self):
        return Application(self.sequence.Application)

    @property
    def Count(self):
        return self.sequence.Count

    @property
    def Parent(self):
        return self.sequence.Parent

    def AddEffect(self, *args,  `_Shape_`=None, `_effectId_`=None, `_Level_`=None, `_trigger_`=None, `_Index_` =None):
        arguments = {" `_Shape_`":  `_Shape_`, "`_effectId_`": `_effectId_`, "`_Level_`": `_Level_`, "`_trigger_`": `_trigger_`, "`_Index_` ": `_Index_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sequence.AddEffect(*args, **arguments)

    def AddTriggerEffect(self, *args,  `_pShape_`=None, `_effectId_`=None, `_trigger_`=None, `_pTriggerShape_`=None, `_bookmark_`=None, `_Level_` =None):
        arguments = {" `_pShape_`":  `_pShape_`, "`_effectId_`": `_effectId_`, "`_trigger_`": `_trigger_`, "`_pTriggerShape_`": `_pTriggerShape_`, "`_bookmark_`": `_bookmark_`, "`_Level_` ": `_Level_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sequence.AddTriggerEffect(*args, **arguments)

    def Clone(self, *args,  `_Effect_`=None, `_Index_` =None):
        arguments = {" `_Effect_`":  `_Effect_`, "`_Index_` ": `_Index_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sequence.Clone(*args, **arguments)

    def ConvertToAfterEffect(self, *args,  `_Effect_`=None, `_After_`=None, `_DimColor_`=None, `_DimSchemeColor_` =None):
        arguments = {" `_Effect_`":  `_Effect_`, "`_After_`": `_After_`, "`_DimColor_`": `_DimColor_`, "`_DimSchemeColor_` ": `_DimSchemeColor_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sequence.ConvertToAfterEffect(*args, **arguments)

    def ConvertToAnimateBackground(self, *args,  `_Effect_`=None, `_AnimateBackground_` =None):
        arguments = {" `_Effect_`":  `_Effect_`, "`_AnimateBackground_` ": `_AnimateBackground_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sequence.ConvertToAnimateBackground(*args, **arguments)

    def ConvertToAnimateInReverse(self, *args,  `_Effect_`=None, `_animateInReverse_` =None):
        arguments = {" `_Effect_`":  `_Effect_`, "`_animateInReverse_` ": `_animateInReverse_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sequence.ConvertToAnimateInReverse(*args, **arguments)

    def ConvertToBuildLevel(self, *args,  `_Effect_`=None, `_Level_` =None):
        arguments = {" `_Effect_`":  `_Effect_`, "`_Level_` ": `_Level_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sequence.ConvertToBuildLevel(*args, **arguments)

    def ConvertToTextUnitEffect(self, *args,  `_Effect_`=None, `_unitEffect_` =None):
        arguments = {" `_Effect_`":  `_Effect_`, "`_unitEffect_` ": `_unitEffect_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sequence.ConvertToTextUnitEffect(*args, **arguments)

    def FindFirstAnimationFor(self, *args,  `_Shape_` =None):
        arguments = {" `_Shape_` ":  `_Shape_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sequence.FindFirstAnimationFor(*args, **arguments)

    def FindFirstAnimationForClick(self, *args,  `_click_` =None):
        arguments = {" `_click_` ":  `_click_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sequence.FindFirstAnimationForClick(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sequence.Item(*args, **arguments)

class Sequences:

    def __init__(self, sequences=None):
        self.sequences = sequences

    @property
    def Application(self):
        return Application(self.sequences.Application)

    @property
    def Count(self):
        return self.sequences.Count

    @property
    def Parent(self):
        return self.sequences.Parent

    def Add(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sequences.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sequences.Item(*args, **arguments)

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
        return XlAxisGroup(self.series.AxisGroup)

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
        return XlColorIndex(self.series.MarkerBackgroundColorIndex)

    @MarkerBackgroundColorIndex.setter
    def MarkerBackgroundColorIndex(self, value):
        self.series.MarkerBackgroundColorIndex = value

    @property
    def MarkerForegroundColor(self):
        return self.series.MarkerForegroundColor

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

    def ApplyDataLabels(self, *args, `Type`=None, `LegendKey`=None, `AutoText`=None, `HasLeaderLines`=None, `ShowSeriesName`=None, `ShowCategoryName`=None, `ShowValue`=None, `ShowPercentage`=None, `ShowBubbleSize`=None, `Separator`=None):
        arguments = {"`Type`": `Type`, "`LegendKey`": `LegendKey`, "`AutoText`": `AutoText`, "`HasLeaderLines`": `HasLeaderLines`, "`ShowSeriesName`": `ShowSeriesName`, "`ShowCategoryName`": `ShowCategoryName`, "`ShowValue`": `ShowValue`, "`ShowPercentage`": `ShowPercentage`, "`ShowBubbleSize`": `ShowBubbleSize`, "`Separator`": `Separator`}
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

    def Extend(self, *args,  `_Source_`=None, `_Rowcol_`=None, `_CategoryLabels_` =None):
        arguments = {" `_Source_`":  `_Source_`, "`_Rowcol_`": `_Rowcol_`, "`_CategoryLabels_` ": `_CategoryLabels_` }
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

class SetEffect:

    def __init__(self, seteffect=None):
        self.seteffect = seteffect

    @property
    def Application(self):
        return Application(self.seteffect.Application)

    @property
    def Parent(self):
        return self.seteffect.Parent

    @property
    def Property(self):
        return self.seteffect.Property

    @property
    def To(self):
        return self.seteffect.To

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
    def ActionSettings(self):
        return ActionSettings(self.shape.ActionSettings)

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
    def AnimationSettings(self):
        return AnimationSettings(self.shape.AnimationSettings)

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
    def BlackWhiteMode(self):
        return self.shape.BlackWhiteMode

    @BlackWhiteMode.setter
    def BlackWhiteMode(self, value):
        self.shape.BlackWhiteMode = value

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
    def Creator(self):
        return self.shape.Creator

    @property
    def CustomerData(self):
        return CustomerData(self.shape.CustomerData)

    @property
    def Decorative(self):
        return self.shape.Decorative

    @property
    def Fill(self):
        return FillFormat(self.shape.Fill)

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
    def HasTable(self):
        return self.shape.HasTable

    @property
    def HasTextFrame(self):
        return self.shape.HasTextFrame

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
    def Id(self):
        return self.shape.Id

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

    @property
    def MediaFormat(self):
        return self.shape.MediaFormat

    @property
    def MediaType(self):
        return self.shape.MediaType

    @property
    def Model3D(self):
        return Model3DFormat(self.shape.Model3D)

    @property
    def Name(self):
        return self.shape.Name

    @property
    def Nodes(self):
        return ShapeNodes(self.shape.Nodes)

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
    def PlaceholderFormat(self):
        return PlaceholderFormat(self.shape.PlaceholderFormat)

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

    @property
    def SmartArt(self):
        return Shape(self.shape.SmartArt)

    @property
    def SoftEdge(self):
        return self.shape.SoftEdge

    @property
    def Table(self):
        return Table(self.shape.Table)

    @property
    def Tags(self):
        return Tags(self.shape.Tags)

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
        return Shape(self.shape.Title)

    @property
    def Top(self):
        return self.shape.Top

    @Top.setter
    def Top(self, value):
        self.shape.Top = value

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
    def ZOrderPosition(self):
        return self.shape.ZOrderPosition

    def Apply(self):
        self.shape.Apply()

    def ApplyAnimation(self):
        self.shape.ApplyAnimation()

    def ConvertTextToSmartArt(self, *args,  `_Layout_` =None):
        arguments = {" `_Layout_` ":  `_Layout_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shape.ConvertTextToSmartArt(*args, **arguments)

    def Copy(self):
        self.shape.Copy()

    def Cut(self):
        self.shape.Cut()

    def Delete(self):
        self.shape.Delete()

    def Duplicate(self):
        return self.shape.Duplicate()

    def Flip(self, *args, FlipCmd=None):
        arguments = {"FlipCmd": FlipCmd}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shape.Flip(*args, **arguments)

    def IncrementLeft(self, *args,  `_Increment_` =None):
        arguments = {" `_Increment_` ":  `_Increment_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shape.IncrementLeft(*args, **arguments)

    def IncrementRotation(self, *args,  `_Increment_` =None):
        arguments = {" `_Increment_` ":  `_Increment_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shape.IncrementRotation(*args, **arguments)

    def IncrementTop(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shape.IncrementTop(*args, **arguments)

    def PickUp(self):
        self.shape.PickUp()

    def PickupAnimation(self):
        self.shape.PickupAnimation()

    def RerouteConnections(self):
        self.shape.RerouteConnections()

    def ScaleHeight(self, *args, Factor=None, RelativeToOriginalSize=None, fScale=None):
        arguments = {"Factor": Factor, "RelativeToOriginalSize": RelativeToOriginalSize, "fScale": fScale}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shape.ScaleHeight(*args, **arguments)

    def ScaleWidth(self, *args, Factor=None, RelativeToOriginalSize=None, fScale=None):
        arguments = {"Factor": Factor, "RelativeToOriginalSize": RelativeToOriginalSize, "fScale": fScale}
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

    def UpgradeMedia(self):
        self.shape.UpgradeMedia()

    def ZOrder(self, *args, ZOrderCmd=None):
        arguments = {"ZOrderCmd": ZOrderCmd}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shape.ZOrder(*args, **arguments)

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

class ShapeNodes:

    def __init__(self, shapenodes=None):
        self.shapenodes = shapenodes

    def __call__(self, item):
        return ShapeNode(self.shapenodes(item))

    @property
    def Application(self):
        return Application(self.shapenodes.Application)

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
        return self.shapenodes.Item(*args, **arguments)

    def SetEditingType(self, *args,  `_Index_`=None, `_EditingType_` =None):
        arguments = {" `_Index_`":  `_Index_`, "`_EditingType_` ": `_EditingType_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shapenodes.SetEditingType(*args, **arguments)

    def SetPosition(self, *args, Index=None, X1=None, Y1=None):
        arguments = {"Index": Index, "X1": X1, "Y1": Y1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shapenodes.SetPosition(*args, **arguments)

    def SetSegmentType(self, *args,  `_Index_`=None, `_SegmentType_` =None):
        arguments = {" `_Index_`":  `_Index_`, "`_SegmentType_` ": `_SegmentType_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shapenodes.SetSegmentType(*args, **arguments)

class ShapeRange:

    def __init__(self, shaperange=None):
        self.shaperange = shaperange

    @property
    def ActionSettings(self):
        return ActionSettings(self.shaperange.ActionSettings)

    @property
    def Adjustments(self):
        return Adjustments(self.shaperange.Adjustments)

    @property
    def AlternativeText(self):
        return self.shaperange.AlternativeText

    @AlternativeText.setter
    def AlternativeText(self, value):
        self.shaperange.AlternativeText = value

    @property
    def AnimationSettings(self):
        return AnimationSettings(self.shaperange.AnimationSettings)

    @property
    def Application(self):
        return Application(self.shaperange.Application)

    @property
    def AutoShapeType(self):
        return ShapeRange(self.shaperange.AutoShapeType)

    @AutoShapeType.setter
    def AutoShapeType(self, value):
        self.shaperange.AutoShapeType = value

    @property
    def BackgroundStyle(self):
        return self.shaperange.BackgroundStyle

    @property
    def BlackWhiteMode(self):
        return self.shaperange.BlackWhiteMode

    @BlackWhiteMode.setter
    def BlackWhiteMode(self, value):
        self.shaperange.BlackWhiteMode = value

    @property
    def Callout(self):
        return CalloutFormat(self.shaperange.Callout)

    @property
    def Chart(self):
        return Chart(self.shaperange.Chart)

    @property
    def Child(self):
        return self.shaperange.Child

    @property
    def ConnectionSiteCount(self):
        return self.shaperange.ConnectionSiteCount

    @property
    def Connector(self):
        return self.shaperange.Connector

    @property
    def ConnectorFormat(self):
        return ConnectorFormat(self.shaperange.ConnectorFormat)

    @property
    def Count(self):
        return self.shaperange.Count

    @property
    def Creator(self):
        return self.shaperange.Creator

    @property
    def CustomerData(self):
        return CustomerData(self.shaperange.CustomerData)

    @property
    def Decorative(self):
        return self.shaperange.Decorative

    @property
    def Fill(self):
        return FillFormat(self.shaperange.Fill)

    @property
    def Glow(self):
        return self.shaperange.Glow

    @property
    def GraphicStyle(self):
        return self.shaperange.GraphicStyle

    @GraphicStyle.setter
    def GraphicStyle(self, value):
        self.shaperange.GraphicStyle = value

    @property
    def GroupItems(self):
        return GroupShapes(self.shaperange.GroupItems)

    @property
    def HasChart(self):
        return self.shaperange.HasChart

    @property
    def HasSmartArt(self):
        return self.shaperange.HasSmartArt

    @property
    def HasTable(self):
        return self.shaperange.HasTable

    @property
    def HasTextFrame(self):
        return self.shaperange.HasTextFrame

    @property
    def Height(self):
        return self.shaperange.Height

    @Height.setter
    def Height(self, value):
        self.shaperange.Height = value

    @property
    def HorizontalFlip(self):
        return self.shaperange.HorizontalFlip

    @property
    def Id(self):
        return self.shaperange.Id

    @property
    def Left(self):
        return self.shaperange.Left

    @Left.setter
    def Left(self, value):
        self.shaperange.Left = value

    @property
    def Line(self):
        return LineFormat(self.shaperange.Line)

    @property
    def LinkFormat(self):
        return LinkFormat(self.shaperange.LinkFormat)

    @property
    def LockAspectRatio(self):
        return self.shaperange.LockAspectRatio

    @property
    def MediaFormat(self):
        return MediaFormat(self.shaperange.MediaFormat)

    @property
    def MediaType(self):
        return self.shaperange.MediaType

    @property
    def Model3D(self):
        return Model3DFormat(self.shaperange.Model3D)

    @property
    def Name(self):
        return self.shaperange.Name

    @property
    def Nodes(self):
        return ShapeNodes(self.shaperange.Nodes)

    @property
    def OLEFormat(self):
        return OLEFormat(self.shaperange.OLEFormat)

    @property
    def Parent(self):
        return self.shaperange.Parent

    @property
    def ParentGroup(self):
        return Shape(self.shaperange.ParentGroup)

    @property
    def PictureFormat(self):
        return PictureFormat(self.shaperange.PictureFormat)

    @property
    def PlaceholderFormat(self):
        return PlaceholderFormat(self.shaperange.PlaceholderFormat)

    @property
    def Reflection(self):
        return self.shaperange.Reflection

    @property
    def Rotation(self):
        return self.shaperange.Rotation

    @Rotation.setter
    def Rotation(self, value):
        self.shaperange.Rotation = value

    @property
    def Shadow(self):
        return ShadowFormat(self.shaperange.Shadow)

    @property
    def ShapeStyle(self):
        return self.shaperange.ShapeStyle

    @property
    def SmartArt(self):
        return ShapeRange(self.shaperange.SmartArt)

    @property
    def SoftEdge(self):
        return self.shaperange.SoftEdge

    @property
    def Table(self):
        return Table(self.shaperange.Table)

    @property
    def Tags(self):
        return Tags(self.shaperange.Tags)

    @property
    def TextEffect(self):
        return TextEffectFormat(self.shaperange.TextEffect)

    @property
    def TextFrame(self):
        return TextFrame(self.shaperange.TextFrame)

    @property
    def TextFrame2(self):
        return TextFrame2(self.shaperange.TextFrame2)

    @property
    def ThreeD(self):
        return ThreeDFormat(self.shaperange.ThreeD)

    @property
    def Title(self):
        return Shape(self.shaperange.Title)

    @property
    def Top(self):
        return self.shaperange.Top

    @Top.setter
    def Top(self, value):
        self.shaperange.Top = value

    @property
    def Type(self):
        return self.shaperange.Type

    @property
    def VerticalFlip(self):
        return self.shaperange.VerticalFlip

    @property
    def Vertices(self):
        return self.shaperange.Vertices

    @property
    def Visible(self):
        return self.shaperange.Visible

    @Visible.setter
    def Visible(self, value):
        self.shaperange.Visible = value

    @property
    def Width(self):
        return self.shaperange.Width

    @Width.setter
    def Width(self, value):
        self.shaperange.Width = value

    @property
    def ZOrderPosition(self):
        return self.shaperange.ZOrderPosition

    def Align(self, *args,  `_AlignCmd_`=None, `_RelativeTo_` =None):
        arguments = {" `_AlignCmd_`":  `_AlignCmd_`, "`_RelativeTo_` ": `_RelativeTo_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shaperange.Align(*args, **arguments)

    def Apply(self):
        self.shaperange.Apply()

    def ApplyAnimation(self):
        self.shaperange.ApplyAnimation()

    def ConvertTextToSmartArt(self, *args,  `_Layout_` =None):
        arguments = {" `_Layout_` ":  `_Layout_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shaperange.ConvertTextToSmartArt(*args, **arguments)

    def Copy(self):
        self.shaperange.Copy()

    def Cut(self):
        self.shaperange.Cut()

    def Delete(self):
        self.shaperange.Delete()

    def Distribute(self, *args,  `_DistributeCmd_`=None, `_RelativeTo_` =None):
        arguments = {" `_DistributeCmd_`":  `_DistributeCmd_`, "`_RelativeTo_` ": `_RelativeTo_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shaperange.Distribute(*args, **arguments)

    def Duplicate(self):
        return self.shaperange.Duplicate()

    def Flip(self, *args, FlipCmd=None):
        arguments = {"FlipCmd": FlipCmd}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shaperange.Flip(*args, **arguments)

    def Group(self):
        return self.shaperange.Group()

    def IncrementLeft(self, *args,  `_Increment_` =None):
        arguments = {" `_Increment_` ":  `_Increment_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shaperange.IncrementLeft(*args, **arguments)

    def IncrementRotation(self, *args,  `_Increment_` =None):
        arguments = {" `_Increment_` ":  `_Increment_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shaperange.IncrementRotation(*args, **arguments)

    def IncrementTop(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shaperange.IncrementTop(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shaperange.Item(*args, **arguments)

    def PickUp(self):
        self.shaperange.PickUp()

    def PickupAnimation(self):
        self.shaperange.PickupAnimation()

    def Regroup(self):
        return self.shaperange.Regroup()

    def RerouteConnections(self):
        self.shaperange.RerouteConnections()

    def ScaleHeight(self, *args,  _Factor=None, RelativeToOriginalSize=None, fScale_ =None):
        arguments = {" _Factor":  _Factor, "RelativeToOriginalSize": RelativeToOriginalSize, "fScale_ ": fScale_ }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shaperange.ScaleHeight(*args, **arguments)

    def ScaleWidth(self, *args, Factor=None, RelativeToOriginalSize=None, fScale=None):
        arguments = {"Factor": Factor, "RelativeToOriginalSize": RelativeToOriginalSize, "fScale": fScale}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shaperange.ScaleWidth(*args, **arguments)

    def Select(self, *args, Replace=None):
        arguments = {"Replace": Replace}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shaperange.Select(*args, **arguments)

    def SetShapesDefaultProperties(self):
        self.shaperange.SetShapesDefaultProperties()

    def Ungroup(self):
        return self.shaperange.Ungroup()

    def UpgradeMedia(self):
        self.shaperange.UpgradeMedia()

    def ZOrder(self, *args, ZOrderCmd=None):
        arguments = {"ZOrderCmd": ZOrderCmd}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shaperange.ZOrder(*args, **arguments)

class Shapes:

    def __init__(self, shapes=None):
        self.shapes = shapes

    def __call__(self, item):
        return Shape(self.shapes(item))

    @property
    def Application(self):
        return Application(self.shapes.Application)

    @property
    def Count(self):
        return self.shapes.Count

    @property
    def Creator(self):
        return self.shapes.Creator

    @property
    def HasTitle(self):
        return self.shapes.HasTitle

    @property
    def Parent(self):
        return self.shapes.Parent

    @property
    def Placeholders(self):
        return Placeholders(self.shapes.Placeholders)

    @property
    def Title(self):
        return Shape(self.shapes.Title)

    def Add3DModel(self, *args, FileName=None, LinkToFile=None, SaveWithDocument=None, Left=None, Top=None, Width=None, Height=None):
        arguments = {"FileName": FileName, "LinkToFile": LinkToFile, "SaveWithDocument": SaveWithDocument, "Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.Add3DModel(*args, **arguments)

    def AddCallout(self, *args,  `_Type_`=None, `_Left_`=None, `_Top_`=None, `_Width_`=None, `_Height_` =None):
        arguments = {" `_Type_`":  `_Type_`, "`_Left_`": `_Left_`, "`_Top_`": `_Top_`, "`_Width_`": `_Width_`, "`_Height_` ": `_Height_` }
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

    def AddLabel(self, *args,  `_Orientation_`=None, `_Left_`=None, `_Top_`=None, `_Width_`=None, `_Height_` =None):
        arguments = {" `_Orientation_`":  `_Orientation_`, "`_Left_`": `_Left_`, "`_Top_`": `_Top_`, "`_Width_`": `_Width_`, "`_Height_` ": `_Height_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddLabel(*args, **arguments)

    def AddLine(self, *args,  `_BeginX_`=None, `_BeginY_`=None, `_EndX_`=None, `_EndY_` =None):
        arguments = {" `_BeginX_`":  `_BeginX_`, "`_BeginY_`": `_BeginY_`, "`_EndX_`": `_EndX_`, "`_EndY_` ": `_EndY_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddLine(*args, **arguments)

    def AddMediaObject(self, *args,  `_FileName_`=None, `_Left_`=None, `_Top_`=None, `_Width_`=None, `_Height_` =None):
        arguments = {" `_FileName_`":  `_FileName_`, "`_Left_`": `_Left_`, "`_Top_`": `_Top_`, "`_Width_`": `_Width_`, "`_Height_` ": `_Height_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddMediaObject(*args, **arguments)

    def AddMediaObject2(self, *args,  `_FileName_`=None, `_LinkToFile_`=None, `_SaveWithDocument_`=None, `_Left_`=None, `_Top_`=None, `_Width_`=None, `_Height_` =None):
        arguments = {" `_FileName_`":  `_FileName_`, "`_LinkToFile_`": `_LinkToFile_`, "`_SaveWithDocument_`": `_SaveWithDocument_`, "`_Left_`": `_Left_`, "`_Top_`": `_Top_`, "`_Width_`": `_Width_`, "`_Height_` ": `_Height_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddMediaObject2(*args, **arguments)

    def AddMediaObjectFromEmbedTag(self, *args,  `_EmbedTag_`=None, `_Left_`=None, `_Top_`=None, `_Width_`=None, `_Height_` =None):
        arguments = {" `_EmbedTag_`":  `_EmbedTag_`, "`_Left_`": `_Left_`, "`_Top_`": `_Top_`, "`_Width_`": `_Width_`, "`_Height_` ": `_Height_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddMediaObjectFromEmbedTag(*args, **arguments)

    def AddOLEObject(self, *args,  `_Left_`=None, `_Top_`=None, `_Width_`=None, `_Height_`=None, `_ClassName_`=None, `_FileName_`=None, `_DisplayAsIcon_`=None, `_IconFileName_`=None, `_IconIndex_`=None, `_IconLabel_`=None, `_Link_` =None):
        arguments = {" `_Left_`":  `_Left_`, "`_Top_`": `_Top_`, "`_Width_`": `_Width_`, "`_Height_`": `_Height_`, "`_ClassName_`": `_ClassName_`, "`_FileName_`": `_FileName_`, "`_DisplayAsIcon_`": `_DisplayAsIcon_`, "`_IconFileName_`": `_IconFileName_`, "`_IconIndex_`": `_IconIndex_`, "`_IconLabel_`": `_IconLabel_`, "`_Link_` ": `_Link_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddOLEObject(*args, **arguments)

    def AddPicture(self, *args,  `_FileName_`=None, `_LinkToFile_`=None, `_SaveWithDocument_`=None, `_Left_`=None, `_Top_`=None, `_Width_`=None, `_Height_` =None):
        arguments = {" `_FileName_`":  `_FileName_`, "`_LinkToFile_`": `_LinkToFile_`, "`_SaveWithDocument_`": `_SaveWithDocument_`, "`_Left_`": `_Left_`, "`_Top_`": `_Top_`, "`_Width_`": `_Width_`, "`_Height_` ": `_Height_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddPicture(*args, **arguments)

    def AddPlaceholder(self, *args,  `_Type_`=None, `_Left_`=None, `_Top_`=None, `_Width_`=None, `_Height_` =None):
        arguments = {" `_Type_`":  `_Type_`, "`_Left_`": `_Left_`, "`_Top_`": `_Top_`, "`_Width_`": `_Width_`, "`_Height_` ": `_Height_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddPlaceholder(*args, **arguments)

    def AddPolyline(self, *args,  `_SafeArrayOfPoints_` =None):
        arguments = {" `_SafeArrayOfPoints_` ":  `_SafeArrayOfPoints_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddPolyline(*args, **arguments)

    def AddShape(self, *args,  `_Type_`=None, `_Left_`=None, `_Top_`=None, `_Width_`=None, `_Height_` =None):
        arguments = {" `_Type_`":  `_Type_`, "`_Left_`": `_Left_`, "`_Top_`": `_Top_`, "`_Width_`": `_Width_`, "`_Height_` ": `_Height_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddShape(*args, **arguments)

    def AddSmartArt(self, *args,  `_Layout_`=None, `_Left_`=None, `_Top_`=None, `_Width_`=None, `_Height_` =None):
        arguments = {" `_Layout_`":  `_Layout_`, "`_Left_`": `_Left_`, "`_Top_`": `_Top_`, "`_Width_`": `_Width_`, "`_Height_` ": `_Height_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddSmartArt(*args, **arguments)

    def AddTable(self, *args,  `_NumRows_`=None, `_NumColumns_`=None, `_Left_`=None, `_Top_`=None, `_Width_`=None, `_Height_` =None):
        arguments = {" `_NumRows_`":  `_NumRows_`, "`_NumColumns_`": `_NumColumns_`, "`_Left_`": `_Left_`, "`_Top_`": `_Top_`, "`_Width_`": `_Width_`, "`_Height_` ": `_Height_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddTable(*args, **arguments)

    def AddTextbox(self, *args,  `_Orientation_`=None, `_Left_`=None, `_Top_`=None, `_Width_`=None, `_Height_` =None):
        arguments = {" `_Orientation_`":  `_Orientation_`, "`_Left_`": `_Left_`, "`_Top_`": `_Top_`, "`_Width_`": `_Width_`, "`_Height_` ": `_Height_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddTextbox(*args, **arguments)

    def AddTextEffect(self, *args,  `_PresetTextEffect_`=None, `_Text_`=None, `_FontName_`=None, `_FontSize_`=None, `_FontBold_`=None, `_FontItalic_`=None, `_Left_`=None, `_Top_` =None):
        arguments = {" `_PresetTextEffect_`":  `_PresetTextEffect_`, "`_Text_`": `_Text_`, "`_FontName_`": `_FontName_`, "`_FontSize_`": `_FontSize_`, "`_FontBold_`": `_FontBold_`, "`_FontItalic_`": `_FontItalic_`, "`_Left_`": `_Left_`, "`_Top_` ": `_Top_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.AddTextEffect(*args, **arguments)

    def AddTitle(self):
        return self.shapes.AddTitle()

    def BuildFreeform(self, *args,  `_EditingType_`=None, `_X1_`=None, `_Y1_` =None):
        arguments = {" `_EditingType_`":  `_EditingType_`, "`_X1_`": `_X1_`, "`_Y1_` ": `_Y1_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.BuildFreeform(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.Item(*args, **arguments)

    def Paste(self):
        return self.shapes.Paste()

    def PasteSpecial(self, *args, DataType=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Link=None):
        arguments = {"DataType": DataType, "DisplayAsIcon": DisplayAsIcon, "IconFileName": IconFileName, "IconIndex": IconIndex, "IconLabel": IconLabel, "Link": Link}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.PasteSpecial(*args, **arguments)

    def Range(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shapes.Range(*args, **arguments)

    def SelectAll(self):
        self.shapes.SelectAll()

class Slide:

    def __init__(self, slide=None):
        self.slide = slide

    @property
    def Application(self):
        return Application(self.slide.Application)

    @property
    def Background(self):
        return ShapeRange(self.slide.Background)

    @property
    def BackgroundStyle(self):
        return self.slide.BackgroundStyle

    @property
    def ColorScheme(self):
        return ColorScheme(self.slide.ColorScheme)

    @ColorScheme.setter
    def ColorScheme(self, value):
        self.slide.ColorScheme = value

    @property
    def Comments(self):
        return Comments(self.slide.Comments)

    @property
    def CustomerData(self):
        return CustomerData(self.slide.CustomerData)

    @property
    def CustomLayout(self):
        return CustomLayout(self.slide.CustomLayout)

    @property
    def Design(self):
        return Design(self.slide.Design)

    @property
    def DisplayMasterShapes(self):
        return self.slide.DisplayMasterShapes

    @property
    def FollowMasterBackground(self):
        return self.slide.FollowMasterBackground

    @property
    def HasNotesPage(self):
        return self.slide.HasNotesPage

    @property
    def HeadersFooters(self):
        return HeadersFooters(self.slide.HeadersFooters)

    @property
    def Hyperlinks(self):
        return Hyperlinks(self.slide.Hyperlinks)

    @property
    def Layout(self):
        return PpSlideLayout(self.slide.Layout)

    @Layout.setter
    def Layout(self, value):
        self.slide.Layout = value

    @property
    def Master(self):
        return Master(self.slide.Master)

    @property
    def Name(self):
        return self.slide.Name

    @property
    def NotesPage(self):
        return SlideRange(self.slide.NotesPage)

    @property
    def Parent(self):
        return self.slide.Parent

    @property
    def PrintSteps(self):
        return self.slide.PrintSteps

    @property
    def sectionIndex(self):
        return Slide(self.slide.sectionIndex)

    @property
    def Shapes(self):
        return Shapes(self.slide.Shapes)

    @property
    def SlideID(self):
        return self.slide.SlideID

    @property
    def SlideIndex(self):
        return Slides(self.slide.SlideIndex)

    @property
    def SlideNumber(self):
        return self.slide.SlideNumber

    @property
    def SlideShowTransition(self):
        return SlideShowTransition(self.slide.SlideShowTransition)

    @property
    def Tags(self):
        return Tags(self.slide.Tags)

    @property
    def ThemeColorScheme(self):
        return self.slide.ThemeColorScheme

    @property
    def TimeLine(self):
        return TimeLine(self.slide.TimeLine)

    def ApplyTemplate(self, *args,  `_FileName_` =None):
        arguments = {" `_FileName_` ":  `_FileName_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.slide.ApplyTemplate(*args, **arguments)

    def ApplyTheme(self, *args,  `_themeName_` =None):
        arguments = {" `_themeName_` ":  `_themeName_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.slide.ApplyTheme(*args, **arguments)

    def ApplyThemeColorScheme(self, *args,  `_themeColorSchemeName_` =None):
        arguments = {" `_themeColorSchemeName_` ":  `_themeColorSchemeName_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.slide.ApplyThemeColorScheme(*args, **arguments)

    def Copy(self):
        self.slide.Copy()

    def Cut(self):
        self.slide.Cut()

    def Delete(self):
        self.slide.Delete()

    def Duplicate(self):
        return self.slide.Duplicate()

    def Export(self, *args, FileName=None, FilterName=None, ScaleWidth=None, ScaleHeight=None):
        arguments = {"FileName": FileName, "FilterName": FilterName, "ScaleWidth": ScaleWidth, "ScaleHeight": ScaleHeight}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.slide.Export(*args, **arguments)

    def MoveTo(self, *args,  `_toPos_` =None):
        arguments = {" `_toPos_` ":  `_toPos_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.slide.MoveTo(*args, **arguments)

    def MoveToSectionStart(self, *args,  `_toSection_` =None):
        arguments = {" `_toSection_` ":  `_toSection_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.slide.MoveToSectionStart(*args, **arguments)

    def PublishSlides(self, *args,  `_SlideLibraryUrl_`=None, `_Overwrite_`=None, `_UseSlideOrder_` =None):
        arguments = {" `_SlideLibraryUrl_`":  `_SlideLibraryUrl_`, "`_Overwrite_`": `_Overwrite_`, "`_UseSlideOrder_` ": `_UseSlideOrder_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.slide.PublishSlides(*args, **arguments)

    def Select(self):
        self.slide.Select()

class SlideRange:

    def __init__(self, sliderange=None):
        self.sliderange = sliderange

    def __call__(self, item):
        return SlideRang(self.sliderange(item))

    @property
    def Application(self):
        return Application(self.sliderange.Application)

    @property
    def Background(self):
        return ShapeRange(self.sliderange.Background)

    @property
    def BackgroundStyle(self):
        return self.sliderange.BackgroundStyle

    @property
    def ColorScheme(self):
        return ColorScheme(self.sliderange.ColorScheme)

    @ColorScheme.setter
    def ColorScheme(self, value):
        self.sliderange.ColorScheme = value

    @property
    def Comments(self):
        return Comments(self.sliderange.Comments)

    @property
    def Count(self):
        return self.sliderange.Count

    @property
    def CustomerData(self):
        return CustomerData(self.sliderange.CustomerData)

    @property
    def CustomLayout(self):
        return CustomLayout(self.sliderange.CustomLayout)

    @property
    def Design(self):
        return Design(self.sliderange.Design)

    @property
    def DisplayMasterShapes(self):
        return self.sliderange.DisplayMasterShapes

    @property
    def FollowMasterBackground(self):
        return self.sliderange.FollowMasterBackground

    @property
    def HasNotesPage(self):
        return self.sliderange.HasNotesPage

    @property
    def HeadersFooters(self):
        return HeadersFooters(self.sliderange.HeadersFooters)

    @property
    def Hyperlinks(self):
        return Hyperlinks(self.sliderange.Hyperlinks)

    @property
    def Layout(self):
        return PpSlideLayout(self.sliderange.Layout)

    @Layout.setter
    def Layout(self, value):
        self.sliderange.Layout = value

    @property
    def Master(self):
        return Master(self.sliderange.Master)

    @property
    def Name(self):
        return self.sliderange.Name

    @property
    def NotesPage(self):
        return SlideRange(self.sliderange.NotesPage)

    @property
    def Parent(self):
        return self.sliderange.Parent

    @property
    def PrintSteps(self):
        return self.sliderange.PrintSteps

    @property
    def sectionIndex(self):
        return SlideRange(self.sliderange.sectionIndex)

    @property
    def Shapes(self):
        return Shapes(self.sliderange.Shapes)

    @property
    def SlideID(self):
        return self.sliderange.SlideID

    @property
    def SlideIndex(self):
        return Slides(self.sliderange.SlideIndex)

    @property
    def SlideNumber(self):
        return self.sliderange.SlideNumber

    @property
    def SlideShowTransition(self):
        return SlideShowTransition(self.sliderange.SlideShowTransition)

    @property
    def Tags(self):
        return Tags(self.sliderange.Tags)

    @property
    def ThemeColorScheme(self):
        return self.sliderange.ThemeColorScheme

    @property
    def TimeLine(self):
        return TimeLine(self.sliderange.TimeLine)

    def ApplyTemplate(self, *args,  `_FileName_` =None):
        arguments = {" `_FileName_` ":  `_FileName_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sliderange.ApplyTemplate(*args, **arguments)

    def ApplyTheme(self, *args,  `_themeName_` =None):
        arguments = {" `_themeName_` ":  `_themeName_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sliderange.ApplyTheme(*args, **arguments)

    def ApplyThemeColorScheme(self, *args,  `_themeColorSchemeName_` =None):
        arguments = {" `_themeColorSchemeName_` ":  `_themeColorSchemeName_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sliderange.ApplyThemeColorScheme(*args, **arguments)

    def Copy(self):
        self.sliderange.Copy()

    def Cut(self):
        self.sliderange.Cut()

    def Delete(self):
        self.sliderange.Delete()

    def Duplicate(self):
        return self.sliderange.Duplicate()

    def Export(self, *args, FileName=None, FilterName=None, ScaleWidth=None, ScaleHeight=None):
        arguments = {"FileName": FileName, "FilterName": FilterName, "ScaleWidth": ScaleWidth, "ScaleHeight": ScaleHeight}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sliderange.Export(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sliderange.Item(*args, **arguments)

    def MoveTo(self, *args,  `_toPos_` =None):
        arguments = {" `_toPos_` ":  `_toPos_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sliderange.MoveTo(*args, **arguments)

    def MoveToSectionStart(self, *args,  `_toSection_` =None):
        arguments = {" `_toSection_` ":  `_toSection_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sliderange.MoveToSectionStart(*args, **arguments)

    def PublishSlides(self, *args,  `_SlideLibraryUrl_`=None, `_Overwrite_` =None):
        arguments = {" `_SlideLibraryUrl_`":  `_SlideLibraryUrl_`, "`_Overwrite_` ": `_Overwrite_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sliderange.PublishSlides(*args, **arguments)

    def Select(self):
        self.sliderange.Select()

class Slides:

    def __init__(self, slides=None):
        self.slides = slides

    def __call__(self, item):
        return Slide(self.slides(item))

    @property
    def Application(self):
        return Application(self.slides.Application)

    @property
    def Count(self):
        return self.slides.Count

    @property
    def Parent(self):
        return self.slides.Parent

    def AddSlide(self, *args,  `_Index_`=None, `_pCustomLayout_` =None):
        arguments = {" `_Index_`":  `_Index_`, "`_pCustomLayout_` ": `_pCustomLayout_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.slides.AddSlide(*args, **arguments)

    def FindBySlideID(self, *args,  `_SlideID_` =None):
        arguments = {" `_SlideID_` ":  `_SlideID_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.slides.FindBySlideID(*args, **arguments)

    def InsertFromFile(self, *args,  `_FileName_`=None, `_Index_`=None, `_SlideStart_`=None, `_SlideEnd_` =None):
        arguments = {" `_FileName_`":  `_FileName_`, "`_Index_`": `_Index_`, "`_SlideStart_`": `_SlideStart_`, "`_SlideEnd_` ": `_SlideEnd_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.slides.InsertFromFile(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.slides.Item(*args, **arguments)

    def Paste(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.slides.Paste(*args, **arguments)

    def Range(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.slides.Range(*args, **arguments)

class SlideShowSettings:

    def __init__(self, slideshowsettings=None):
        self.slideshowsettings = slideshowsettings

    @property
    def AdvanceMode(self):
        return self.slideshowsettings.AdvanceMode

    @AdvanceMode.setter
    def AdvanceMode(self, value):
        self.slideshowsettings.AdvanceMode = value

    @property
    def Application(self):
        return Application(self.slideshowsettings.Application)

    @property
    def EndingSlide(self):
        return self.slideshowsettings.EndingSlide

    @EndingSlide.setter
    def EndingSlide(self, value):
        self.slideshowsettings.EndingSlide = value

    @property
    def LoopUntilStopped(self):
        return self.slideshowsettings.LoopUntilStopped

    @property
    def NamedSlideShows(self):
        return NamedSlideShows(self.slideshowsettings.NamedSlideShows)

    @property
    def Parent(self):
        return self.slideshowsettings.Parent

    @property
    def PointerColor(self):
        return ColorFormat(self.slideshowsettings.PointerColor)

    @property
    def RangeType(self):
        return self.slideshowsettings.RangeType

    @RangeType.setter
    def RangeType(self, value):
        self.slideshowsettings.RangeType = value

    @property
    def ShowMediaControls(self):
        return self.slideshowsettings.ShowMediaControls

    @property
    def ShowPresenterView(self):
        return SlideShowSettings(self.slideshowsettings.ShowPresenterView)

    @property
    def ShowScrollbar(self):
        return self.slideshowsettings.ShowScrollbar

    @property
    def ShowType(self):
        return self.slideshowsettings.ShowType

    @ShowType.setter
    def ShowType(self, value):
        self.slideshowsettings.ShowType = value

    @property
    def ShowWithAnimation(self):
        return self.slideshowsettings.ShowWithAnimation

    @property
    def ShowWithNarration(self):
        return self.slideshowsettings.ShowWithNarration

    @property
    def SlideShowName(self):
        return self.slideshowsettings.SlideShowName

    @SlideShowName.setter
    def SlideShowName(self, value):
        self.slideshowsettings.SlideShowName = value

    @property
    def StartingSlide(self):
        return self.slideshowsettings.StartingSlide

    @StartingSlide.setter
    def StartingSlide(self, value):
        self.slideshowsettings.StartingSlide = value

    def Run(self):
        return self.slideshowsettings.Run()

class SlideShowTransition:

    def __init__(self, slideshowtransition=None):
        self.slideshowtransition = slideshowtransition

    @property
    def AdvanceOnClick(self):
        return self.slideshowtransition.AdvanceOnClick

    @property
    def AdvanceOnTime(self):
        return self.slideshowtransition.AdvanceOnTime

    @property
    def AdvanceTime(self):
        return self.slideshowtransition.AdvanceTime

    @AdvanceTime.setter
    def AdvanceTime(self, value):
        self.slideshowtransition.AdvanceTime = value

    @property
    def Application(self):
        return Application(self.slideshowtransition.Application)

    @property
    def Duration(self):
        return self.slideshowtransition.Duration

    @Duration.setter
    def Duration(self, value):
        self.slideshowtransition.Duration = value

    @property
    def EntryEffect(self):
        return self.slideshowtransition.EntryEffect

    @EntryEffect.setter
    def EntryEffect(self, value):
        self.slideshowtransition.EntryEffect = value

    @property
    def Hidden(self):
        return self.slideshowtransition.Hidden

    @property
    def LoopSoundUntilNext(self):
        return self.slideshowtransition.LoopSoundUntilNext

    @property
    def Parent(self):
        return self.slideshowtransition.Parent

    @property
    def SoundEffect(self):
        return SoundEffect(self.slideshowtransition.SoundEffect)

    @property
    def Speed(self):
        return self.slideshowtransition.Speed

class SlideShowView:

    def __init__(self, slideshowview=None):
        self.slideshowview = slideshowview

    @property
    def AcceleratorsEnabled(self):
        return self.slideshowview.AcceleratorsEnabled

    @property
    def AdvanceMode(self):
        return self.slideshowview.AdvanceMode

    @property
    def Application(self):
        return Application(self.slideshowview.Application)

    @property
    def CurrentShowPosition(self):
        return self.slideshowview.CurrentShowPosition

    @property
    def IsNamedShow(self):
        return self.slideshowview.IsNamedShow

    @property
    def LastSlideViewed(self):
        return Slide(self.slideshowview.LastSlideViewed)

    @property
    def MediaControlsHeight(self):
        return self.slideshowview.MediaControlsHeight

    @property
    def MediaControlsLeft(self):
        return Slide(self.slideshowview.MediaControlsLeft)

    @property
    def MediaControlsTop(self):
        return Slide(self.slideshowview.MediaControlsTop)

    @property
    def MediaControlsVisible(self):
        return self.slideshowview.MediaControlsVisible

    @property
    def MediaControlsWidth(self):
        return self.slideshowview.MediaControlsWidth

    @property
    def Parent(self):
        return self.slideshowview.Parent

    @property
    def PointerColor(self):
        return ColorFormat(self.slideshowview.PointerColor)

    @property
    def PointerType(self):
        return self.slideshowview.PointerType

    @PointerType.setter
    def PointerType(self, value):
        self.slideshowview.PointerType = value

    @property
    def PresentationElapsedTime(self):
        return self.slideshowview.PresentationElapsedTime

    @property
    def Slide(self):
        return Slide(self.slideshowview.Slide)

    @property
    def SlideElapsedTime(self):
        return self.slideshowview.SlideElapsedTime

    @property
    def SlideShowName(self):
        return self.slideshowview.SlideShowName

    @property
    def State(self):
        return self.slideshowview.State

    @State.setter
    def State(self, value):
        self.slideshowview.State = value

    @property
    def Zoom(self):
        return self.slideshowview.Zoom

    def DrawLine(self, *args,  `_BeginX_`=None, `_BeginY_`=None, `_EndX_`=None, `_EndY_` =None):
        arguments = {" `_BeginX_`":  `_BeginX_`, "`_BeginY_`": `_BeginY_`, "`_EndX_`": `_EndX_`, "`_EndY_` ": `_EndY_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.slideshowview.DrawLine(*args, **arguments)

    def EndNamedShow(self):
        self.slideshowview.EndNamedShow()

    def EraseDrawing(self):
        self.slideshowview.EraseDrawing()

    def Exit(self):
        self.slideshowview.Exit()

    def First(self):
        return self.slideshowview.First()

    def FirstAnimationIsAutomatic(self):
        return self.slideshowview.FirstAnimationIsAutomatic()

    def GetClickCount(self):
        return self.slideshowview.GetClickCount()

    def GetClickIndex(self):
        return self.slideshowview.GetClickIndex()

    def GotoClick(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.slideshowview.GotoClick(*args, **arguments)

    def GotoNamedShow(self, *args,  `_SlideShowName_` =None):
        arguments = {" `_SlideShowName_` ":  `_SlideShowName_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.slideshowview.GotoNamedShow(*args, **arguments)

    def GotoSlide(self, *args,  `_Index_`=None, `_ResetSlide_` =None):
        arguments = {" `_Index_`":  `_Index_`, "`_ResetSlide_` ": `_ResetSlide_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.slideshowview.GotoSlide(*args, **arguments)

    def Last(self):
        self.slideshowview.Last()

    def Next(self):
        self.slideshowview.Next()

    def Player(self, *args,  `_ShapeId_` =None):
        arguments = {" `_ShapeId_` ":  `_ShapeId_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.slideshowview.Player(*args, **arguments)

    def Previous(self):
        self.slideshowview.Previous()

    def ResetSlideTime(self):
        self.slideshowview.ResetSlideTime()

class SlideShowWindow:

    def __init__(self, slideshowwindow=None):
        self.slideshowwindow = slideshowwindow

    @property
    def Active(self):
        return self.slideshowwindow.Active

    @property
    def Application(self):
        return Application(self.slideshowwindow.Application)

    @property
    def Height(self):
        return self.slideshowwindow.Height

    @Height.setter
    def Height(self, value):
        self.slideshowwindow.Height = value

    @property
    def IsFullScreen(self):
        return self.slideshowwindow.IsFullScreen

    @property
    def Left(self):
        return self.slideshowwindow.Left

    @Left.setter
    def Left(self, value):
        self.slideshowwindow.Left = value

    @property
    def Parent(self):
        return self.slideshowwindow.Parent

    @property
    def Presentation(self):
        return Presentation(self.slideshowwindow.Presentation)

    @property
    def Top(self):
        return self.slideshowwindow.Top

    @Top.setter
    def Top(self, value):
        self.slideshowwindow.Top = value

    @property
    def View(self):
        return SlideShowView(self.slideshowwindow.View)

    @property
    def Width(self):
        return self.slideshowwindow.Width

    @Width.setter
    def Width(self, value):
        self.slideshowwindow.Width = value

    def Activate(self):
        self.slideshowwindow.Activate()

class SlideShowWindows:

    def __init__(self, slideshowwindows=None):
        self.slideshowwindows = slideshowwindows

    def __call__(self, item):
        return SlideShowWindow(self.slideshowwindows(item))

    @property
    def Application(self):
        return Application(self.slideshowwindows.Application)

    @property
    def Count(self):
        return self.slideshowwindows.Count

    @property
    def Parent(self):
        return self.slideshowwindows.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.slideshowwindows.Item(*args, **arguments)

class SoundEffect:

    def __init__(self, soundeffect=None):
        self.soundeffect = soundeffect

    @property
    def Application(self):
        return Application(self.soundeffect.Application)

    @property
    def Name(self):
        return self.soundeffect.Name

    @Name.setter
    def Name(self, value):
        self.soundeffect.Name = value

    @property
    def Parent(self):
        return self.soundeffect.Parent

    @property
    def Type(self):
        return self.soundeffect.Type

    def ImportFromFile(self, *args,  `_FullName_` =None):
        arguments = {" `_FullName_` ":  `_FullName_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.soundeffect.ImportFromFile(*args, **arguments)

    def Play(self):
        self.soundeffect.Play()

class Table:

    def __init__(self, table=None):
        self.table = table

    @property
    def AlternativeText(self):
        return self.table.AlternativeText

    @AlternativeText.setter
    def AlternativeText(self, value):
        self.table.AlternativeText = value

    @property
    def Application(self):
        return Application(self.table.Application)

    @property
    def Background(self):
        return TableBackground(self.table.Background)

    @property
    def Columns(self):
        return Columns(self.table.Columns)

    @property
    def FirstCol(self):
        return self.table.FirstCol

    @property
    def FirstRow(self):
        return self.table.FirstRow

    @property
    def HorizBanding(self):
        return self.table.HorizBanding

    @property
    def LastCol(self):
        return self.table.LastCol

    @property
    def LastRow(self):
        return self.table.LastRow

    @property
    def Parent(self):
        return self.table.Parent

    @property
    def Rows(self):
        return Rows(self.table.Rows)

    @property
    def Style(self):
        return TableStyle(self.table.Style)

    @property
    def TableDirection(self):
        return self.table.TableDirection

    @TableDirection.setter
    def TableDirection(self, value):
        self.table.TableDirection = value

    @property
    def Title(self):
        return Table(self.table.Title)

    @Title.setter
    def Title(self, value):
        self.table.Title = value

    @property
    def VertBanding(self):
        return self.table.VertBanding

    def ApplyStyle(self, *args,  `_StyleID_`=None, `_SaveFormatting_` =None):
        arguments = {" `_StyleID_`":  `_StyleID_`, "`_SaveFormatting_` ": `_SaveFormatting_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.table.ApplyStyle(*args, **arguments)

    def Cell(self, *args,  `_Row_`=None, `_Column_` =None):
        arguments = {" `_Row_`":  `_Row_`, "`_Column_` ": `_Column_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.table.Cell(*args, **arguments)

    def ScaleProportionally(self, *args,  `_scale_` =None):
        arguments = {" `_scale_` ":  `_scale_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.table.ScaleProportionally(*args, **arguments)

class TableBackground:

    def __init__(self, tablebackground=None):
        self.tablebackground = tablebackground

    @property
    def Fill(self):
        return FillFormat(self.tablebackground.Fill)

    @property
    def Picture(self):
        return PictureFormat(self.tablebackground.Picture)

    @property
    def Reflection(self):
        return self.tablebackground.Reflection

    @property
    def Shadow(self):
        return ShadowFormat(self.tablebackground.Shadow)

class TableStyle:

    def __init__(self, tablestyle=None):
        self.tablestyle = tablestyle

    @property
    def Id(self):
        return self.tablestyle.Id

    @property
    def Name(self):
        return self.tablestyle.Name

class TabStop:

    def __init__(self, tabstop=None):
        self.tabstop = tabstop

    @property
    def Application(self):
        return Application(self.tabstop.Application)

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
    def Type(self):
        return self.tabstop.Type

    def Clear(self):
        self.tabstop.Clear()

class TabStops:

    def __init__(self, tabstops=None):
        self.tabstops = tabstops

    def __call__(self, item):
        return TabStop(self.tabstops(item))

    @property
    def Application(self):
        return Application(self.tabstops.Application)

    @property
    def Count(self):
        return self.tabstops.Count

    @property
    def DefaultSpacing(self):
        return self.tabstops.DefaultSpacing

    @DefaultSpacing.setter
    def DefaultSpacing(self, value):
        self.tabstops.DefaultSpacing = value

    @property
    def Parent(self):
        return self.tabstops.Parent

    def Add(self, *args, Type=None, Position=None):
        arguments = {"Type": Type, "Position": Position}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return TabStop(self.tabstops.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.tabstops.Item(*args, **arguments)

class Tags:

    def __init__(self, tags=None):
        self.tags = tags

    @property
    def Application(self):
        return Application(self.tags.Application)

    @property
    def Count(self):
        return self.tags.Count

    @property
    def Parent(self):
        return self.tags.Parent

    def Add(self, *args, Name=None, Value=None):
        arguments = {"Name": Name, "Value": Value}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.tags.Add(*args, **arguments)

    def Delete(self, *args,  `_Name_` =None):
        arguments = {" `_Name_` ":  `_Name_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.tags.Delete(*args, **arguments)

    def Item(self, *args, Name=None):
        arguments = {"Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.tags.Item(*args, **arguments)

    def Name(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.tags.Name(*args, **arguments)

    def Value(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.tags.Value(*args, **arguments)

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
    def Orientation(self):
        return self.textframe.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.textframe.Orientation = value

    @property
    def Parent(self):
        return self.textframe.Parent

    @property
    def Ruler(self):
        return Ruler(self.textframe.Ruler)

    @property
    def TextRange(self):
        return TextRange(self.textframe.TextRange)

    @property
    def VerticalAnchor(self):
        return self.textframe.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.textframe.VerticalAnchor = value

    @property
    def WordWrap(self):
        return self.textframe.WordWrap

    def DeleteText(self):
        self.textframe.DeleteText()

class TextFrame2:

    def __init__(self, textframe2=None):
        self.textframe2 = textframe2

    @property
    def Application(self):
        return Application(self.textframe2.Application)

    @property
    def AutoSize(self):
        return self.textframe2.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.textframe2.AutoSize = value

    @property
    def Column(self):
        return Column(self.textframe2.Column)

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
    def WordArtFormat(self):
        return self.textframe2.WordArtFormat

    @WordArtFormat.setter
    def WordArtFormat(self, value):
        self.textframe2.WordArtFormat = value

    @property
    def WordWrap(self):
        return self.textframe2.WordWrap

    def DeleteText(self):
        return self.textframe2.DeleteText()

class TextRange:

    def __init__(self, textrange=None):
        self.textrange = textrange

    @property
    def ActionSettings(self):
        return ActionSettings(self.textrange.ActionSettings)

    @property
    def Application(self):
        return Application(self.textrange.Application)

    @property
    def BoundHeight(self):
        return self.textrange.BoundHeight

    @property
    def BoundLeft(self):
        return self.textrange.BoundLeft

    @property
    def BoundTop(self):
        return self.textrange.BoundTop

    @property
    def BoundWidth(self):
        return self.textrange.BoundWidth

    @property
    def Count(self):
        return self.textrange.Count

    @property
    def Font(self):
        return Font(self.textrange.Font)

    @property
    def IndentLevel(self):
        return self.textrange.IndentLevel

    @IndentLevel.setter
    def IndentLevel(self, value):
        self.textrange.IndentLevel = value

    @property
    def LanguageID(self):
        return self.textrange.LanguageID

    @LanguageID.setter
    def LanguageID(self, value):
        self.textrange.LanguageID = value

    @property
    def Length(self):
        return self.textrange.Length

    @property
    def ParagraphFormat(self):
        return ParagraphFormat(self.textrange.ParagraphFormat)

    @property
    def Parent(self):
        return self.textrange.Parent

    @property
    def Start(self):
        return self.textrange.Start

    @property
    def Text(self):
        return self.textrange.Text

    @Text.setter
    def Text(self, value):
        self.textrange.Text = value

    def AddPeriods(self):
        self.textrange.AddPeriods()

    def ChangeCase(self, *args,  `_Type_` =None):
        arguments = {" `_Type_` ":  `_Type_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.textrange.ChangeCase(*args, **arguments)

    def Characters(self, *args,  `_Start_`=None, `_Length_` =None):
        arguments = {" `_Start_`":  `_Start_`, "`_Length_` ": `_Length_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.textrange.Characters(*args, **arguments)

    def Copy(self):
        self.textrange.Copy()

    def Cut(self):
        self.textrange.Cut()

    def Delete(self):
        self.textrange.Delete()

    def Find(self, *args, FindWhat=None, After=None, MatchCase=None, WholeWords=None):
        arguments = {"FindWhat": FindWhat, "After": After, "MatchCase": MatchCase, "WholeWords": WholeWords}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.textrange.Find(*args, **arguments)

    def InsertAfter(self, *args,  `_NewText_` =None):
        arguments = {" `_NewText_` ":  `_NewText_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.textrange.InsertAfter(*args, **arguments)

    def InsertBefore(self, *args,  `_NewText_` =None):
        arguments = {" `_NewText_` ":  `_NewText_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.textrange.InsertBefore(*args, **arguments)

    def InsertDateTime(self, *args,  _DateTimeFormat=None, InsertAsField_ =None):
        arguments = {" _DateTimeFormat":  _DateTimeFormat, "InsertAsField_ ": InsertAsField_ }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.textrange.InsertDateTime(*args, **arguments)

    def InsertSlideNumber(self):
        return self.textrange.InsertSlideNumber()

    def InsertSymbol(self, *args,  `_FontName_`=None, `_CharNumber_`=None, `_UniCode_` =None):
        arguments = {" `_FontName_`":  `_FontName_`, "`_CharNumber_`": `_CharNumber_`, "`_UniCode_` ": `_UniCode_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.textrange.InsertSymbol(*args, **arguments)

    def Lines(self, *args,  `_Start_`=None, `_Length_` =None):
        arguments = {" `_Start_`":  `_Start_`, "`_Length_` ": `_Length_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.textrange.Lines(*args, **arguments)

    def LtrRun(self):
        self.textrange.LtrRun()

    def Paragraphs(self, *args,  `_Start_`=None, `_Length_` =None):
        arguments = {" `_Start_`":  `_Start_`, "`_Length_` ": `_Length_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.textrange.Paragraphs(*args, **arguments)

    def Paste(self):
        return self.textrange.Paste()

    def PasteSpecial(self, *args, DataType=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Link=None):
        arguments = {"DataType": DataType, "DisplayAsIcon": DisplayAsIcon, "IconFileName": IconFileName, "IconIndex": IconIndex, "IconLabel": IconLabel, "Link": Link}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.textrange.PasteSpecial(*args, **arguments)

    def RemovePeriods(self):
        self.textrange.RemovePeriods()

    def Replace(self, *args, FindWhat=None, ReplaceWhat=None, After=None, MatchCase=None, WholeWords=None):
        arguments = {"FindWhat": FindWhat, "ReplaceWhat": ReplaceWhat, "After": After, "MatchCase": MatchCase, "WholeWords": WholeWords}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.textrange.Replace(*args, **arguments)

    def RotatedBounds(self, *args,  `_X1_`=None, `_Y1_`=None, `_X2_`=None, `_Y2_`=None, `_X3_`=None, `_Y3_`=None, `_X4_`=None, `_Y4_` =None):
        arguments = {" `_X1_`":  `_X1_`, "`_Y1_`": `_Y1_`, "`_X2_`": `_X2_`, "`_Y2_`": `_Y2_`, "`_X3_`": `_X3_`, "`_Y3_`": `_Y3_`, "`_X4_`": `_X4_`, "`_Y4_` ": `_Y4_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.textrange.RotatedBounds(*args, **arguments)

    def RtlRun(self):
        self.textrange.RtlRun()

    def Runs(self, *args,  `_Start_`=None, `_Length_` =None):
        arguments = {" `_Start_`":  `_Start_`, "`_Length_` ": `_Length_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.textrange.Runs(*args, **arguments)

    def Select(self):
        self.textrange.Select()

    def Sentences(self, *args,  `_Start_`=None, `_Length_` =None):
        arguments = {" `_Start_`":  `_Start_`, "`_Length_` ": `_Length_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.textrange.Sentences(*args, **arguments)

    def TrimText(self):
        return self.textrange.TrimText()

    def Words(self, *args,  `_Start_`=None, `_Length_` =None):
        arguments = {" `_Start_`":  `_Start_`, "`_Length_` ": `_Length_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.textrange.Words(*args, **arguments)

class TextStyle:

    def __init__(self, textstyle=None):
        self.textstyle = textstyle

    @property
    def Application(self):
        return Application(self.textstyle.Application)

    @property
    def Levels(self):
        return TextStyleLevels(self.textstyle.Levels)

    @property
    def Parent(self):
        return self.textstyle.Parent

    @property
    def Ruler(self):
        return Ruler(self.textstyle.Ruler)

    @property
    def TextFrame(self):
        return TextFrame(self.textstyle.TextFrame)

class TextStyleLevel:

    def __init__(self, textstylelevel=None):
        self.textstylelevel = textstylelevel

    @property
    def Application(self):
        return Application(self.textstylelevel.Application)

    @property
    def Font(self):
        return Font(self.textstylelevel.Font)

    @property
    def ParagraphFormat(self):
        return ParagraphFormat(self.textstylelevel.ParagraphFormat)

    @property
    def Parent(self):
        return self.textstylelevel.Parent

class TextStyleLevels:

    def __init__(self, textstylelevels=None):
        self.textstylelevels = textstylelevels

    def __call__(self, item):
        return TextStyleLevel(self.textstylelevels(item))

    @property
    def Application(self):
        return Application(self.textstylelevels.Application)

    @property
    def Count(self):
        return self.textstylelevels.Count

    @property
    def Parent(self):
        return self.textstylelevels.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.textstylelevels.Item(*args, **arguments)

class TextStyles:

    def __init__(self, textstyles=None):
        self.textstyles = textstyles

    def __call__(self, item):
        return TextStyle(self.textstyles(item))

    @property
    def Application(self):
        return Application(self.textstyles.Application)

    @property
    def Count(self):
        return self.textstyles.Count

    @property
    def Parent(self):
        return self.textstyles.Parent

    def Item(self, *args, Type=None):
        arguments = {"Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.textstyles.Item(*args, **arguments)

class ThreeDFormat:

    def __init__(self, threedformat=None):
        self.threedformat = threedformat

    @property
    def Application(self):
        return Application(self.threedformat.Application)

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
        return self.threedformat.LightAngle

    @property
    def Parent(self):
        return self.threedformat.Parent

    @property
    def Perspective(self):
        return self.threedformat.Perspective

    @property
    def PresetCamera(self):
        return ThreeDFormat(self.threedformat.PresetCamera)

    @property
    def PresetExtrusionDirection(self):
        return self.threedformat.PresetExtrusionDirection

    @property
    def PresetLighting(self):
        return ThreeDFormat(self.threedformat.PresetLighting)

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
        return ThreeDFormat(self.threedformat.RotationZ)

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

    def IncrementRotationHorizontal(self, *args,  `_Increment_` =None):
        arguments = {" `_Increment_` ":  `_Increment_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.threedformat.IncrementRotationHorizontal(*args, **arguments)

    def IncrementRotationVertical(self, *args,  `_Increment_` =None):
        arguments = {" `_Increment_` ":  `_Increment_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.threedformat.IncrementRotationVertical(*args, **arguments)

    def IncrementRotationX(self, *args,  `_Increment_` =None):
        arguments = {" `_Increment_` ":  `_Increment_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.threedformat.IncrementRotationX(*args, **arguments)

    def IncrementRotationY(self, *args,  `_Increment_` =None):
        arguments = {" `_Increment_` ":  `_Increment_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.threedformat.IncrementRotationY(*args, **arguments)

    def IncrementRotationZ(self, *args,  `_Increment_` =None):
        arguments = {" `_Increment_` ":  `_Increment_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.threedformat.IncrementRotationZ(*args, **arguments)

    def ResetRotation(self):
        self.threedformat.ResetRotation()

    def SetExtrusionDirection(self, *args,  `_PresetExtrusionDirection_` =None):
        arguments = {" `_PresetExtrusionDirection_` ":  `_PresetExtrusionDirection_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.threedformat.SetExtrusionDirection(*args, **arguments)

    def SetPresetCamera(self, *args,  `_PresetCamera_` =None):
        arguments = {" `_PresetCamera_` ":  `_PresetCamera_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.threedformat.SetPresetCamera(*args, **arguments)

    def SetThreeDFormat(self, *args,  `_PresetThreeDFormat_` =None):
        arguments = {" `_PresetThreeDFormat_` ":  `_PresetThreeDFormat_` }
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

class TimeLine:

    def __init__(self, timeline=None):
        self.timeline = timeline

    @property
    def Application(self):
        return Application(self.timeline.Application)

    @property
    def InteractiveSequences(self):
        return Sequences(self.timeline.InteractiveSequences)

    @property
    def MainSequence(self):
        return Sequence(self.timeline.MainSequence)

    @property
    def Parent(self):
        return self.timeline.Parent

class Timing:

    def __init__(self, timing=None):
        self.timing = timing

    @property
    def Accelerate(self):
        return self.timing.Accelerate

    @Accelerate.setter
    def Accelerate(self, value):
        self.timing.Accelerate = value

    @property
    def Application(self):
        return Application(self.timing.Application)

    @property
    def AutoReverse(self):
        return self.timing.AutoReverse

    @property
    def BounceEnd(self):
        return self.timing.BounceEnd

    @property
    def BounceEndIntensity(self):
        return self.timing.BounceEndIntensity

    @property
    def Decelerate(self):
        return self.timing.Decelerate

    @property
    def Parent(self):
        return self.timing.Parent

    @property
    def RepeatCount(self):
        return self.timing.RepeatCount

    @property
    def RepeatDuration(self):
        return self.timing.RepeatDuration

    @property
    def Restart(self):
        return self.timing.Restart

    @property
    def RewindAtEnd(self):
        return self.timing.RewindAtEnd

    @property
    def SmoothEnd(self):
        return self.timing.SmoothEnd

    @property
    def SmoothStart(self):
        return self.timing.SmoothStart

    @property
    def Speed(self):
        return self.timing.Speed

    @Speed.setter
    def Speed(self, value):
        self.timing.Speed = value

    @property
    def triggerBookmark(self):
        return self.timing.triggerBookmark

    @property
    def TriggerDelayTime(self):
        return self.timing.TriggerDelayTime

    @property
    def TriggerShape(self):
        return self.timing.TriggerShape

    @property
    def TriggerType(self):
        return self.timing.TriggerType

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
        return XlTrendlineType(self.trendline.Type)

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
    def Fill(self):
        return FillFormat(self.upbars.Fill)

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
        self.upbars.Delete()

    def Select(self):
        self.upbars.Select()

class View:

    def __init__(self, view=None):
        self.view = view

    @property
    def Application(self):
        return Application(self.view.Application)

    @property
    def DisplaySlideMiniature(self):
        return self.view.DisplaySlideMiniature

    @property
    def MediaControlsHeight(self):
        return self.view.MediaControlsHeight

    @property
    def MediaControlsLeft(self):
        return self.view.MediaControlsLeft

    @property
    def MediaControlsTop(self):
        return self.view.MediaControlsTop

    @property
    def MediaControlsVisible(self):
        return self.view.MediaControlsVisible

    @property
    def MediaControlsWidth(self):
        return self.view.MediaControlsWidth

    @property
    def Parent(self):
        return self.view.Parent

    @property
    def PrintOptions(self):
        return PrintOptions(self.view.PrintOptions)

    @property
    def Slide(self):
        return Slide(self.view.Slide)

    @Slide.setter
    def Slide(self, value):
        self.view.Slide = value

    @property
    def Type(self):
        return self.view.Type

    @property
    def Zoom(self):
        return self.view.Zoom

    @Zoom.setter
    def Zoom(self, value):
        self.view.Zoom = value

    @property
    def ZoomToFit(self):
        return self.view.ZoomToFit

    def GotoSlide(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.view.GotoSlide(*args, **arguments)

    def Paste(self):
        self.view.Paste()

    def PasteSpecial(self, *args, DataType=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Link=None):
        arguments = {"DataType": DataType, "DisplayAsIcon": DisplayAsIcon, "IconFileName": IconFileName, "IconIndex": IconIndex, "IconLabel": IconLabel, "Link": Link}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.view.PasteSpecial(*args, **arguments)

    def Player(self, *args,  `_ShapeId_` =None):
        arguments = {" `_ShapeId_` ":  `_ShapeId_` }
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.view.Player(*args, **arguments)

    def PrintOut(self, *args, From=None, To=None, PrintToFile=None, Copies=None, Collate=None):
        arguments = {"From": From, "To": To, "PrintToFile": PrintToFile, "Copies": Copies, "Collate": Collate}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.view.PrintOut(*args, **arguments)

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
