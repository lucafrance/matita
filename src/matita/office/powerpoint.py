import win32com.client
import pythoncom

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

    @AnimateAction.setter
    def AnimateAction(self, value):
        self.actionsetting.AnimateAction = value

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

    @ShowAndReturn.setter
    def ShowAndReturn(self, value):
        self.actionsetting.ShowAndReturn = value

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

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.actionsettings.Item(*params)


class AddIn:

    def __init__(self, addin=None):
        self.addin = addin

    @property
    def Application(self):
        return Application(self.addin.Application)

    @property
    def AutoLoad(self):
        return self.addin.AutoLoad

    @AutoLoad.setter
    def AutoLoad(self, value):
        self.addin.AutoLoad = value

    @property
    def FullName(self):
        return self.addin.FullName

    @property
    def Loaded(self):
        return self.addin.Loaded

    @Loaded.setter
    def Loaded(self, value):
        self.addin.Loaded = value

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

    @Registered.setter
    def Registered(self, value):
        self.addin.Registered = value


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

    def Add(self, FileName=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
        ]
        return AddIn(self.addins.Add(*params))

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.addins.Item(*params)

    def Remove(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        self.addins.Remove(*params)


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

    @Accumulate.setter
    def Accumulate(self, value):
        self.animationbehavior.Accumulate = value

    @property
    def Additive(self):
        return self.animationbehavior.Additive

    @Additive.setter
    def Additive(self, value):
        self.animationbehavior.Additive = value

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

    @Type.setter
    def Type(self, value):
        self.animationbehavior.Type = value

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

    def Add(self, Type=None, Index=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.animationbehaviors.Add(*params)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.animationbehaviors.Item(*params)


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

    @Time.setter
    def Time(self, value):
        self.animationpoint.Time = value

    @property
    def Value(self):
        return self.animationpoint.Value

    @Value.setter
    def Value(self, value):
        self.animationpoint.Value = value

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

    @Smooth.setter
    def Smooth(self, value):
        self.animationpoints.Smooth = value

    def Add(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.animationpoints.Add(*params)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.animationpoints.Item(*params)


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

    @Animate.setter
    def Animate(self, value):
        self.animationsettings.Animate = value

    @property
    def AnimateBackground(self):
        return self.animationsettings.AnimateBackground

    @AnimateBackground.setter
    def AnimateBackground(self, value):
        self.animationsettings.AnimateBackground = value

    @property
    def AnimateTextInReverse(self):
        return self.animationsettings.AnimateTextInReverse

    @AnimateTextInReverse.setter
    def AnimateTextInReverse(self, value):
        self.animationsettings.AnimateTextInReverse = value

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

    @EntryEffect.setter
    def EntryEffect(self, value):
        self.animationsettings.EntryEffect = value

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

    @TextLevelEffect.setter
    def TextLevelEffect(self, value):
        self.animationsettings.TextLevelEffect = value

    @property
    def TextUnitEffect(self):
        return self.animationsettings.TextUnitEffect

    @TextUnitEffect.setter
    def TextUnitEffect(self, value):
        self.animationsettings.TextUnitEffect = value


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

    @AutomationSecurity.setter
    def AutomationSecurity(self, value):
        self.application.AutomationSecurity = value

    @property
    def Build(self):
        return self.application.Build

    @property
    def Caption(self):
        return self.application.Caption

    @Caption.setter
    def Caption(self, value):
        self.application.Caption = value

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

    @DisplayAlerts.setter
    def DisplayAlerts(self, value):
        self.application.DisplayAlerts = value

    @property
    def DisplayDocumentInformationPanel(self):
        return self.application.DisplayDocumentInformationPanel

    @DisplayDocumentInformationPanel.setter
    def DisplayDocumentInformationPanel(self, value):
        self.application.DisplayDocumentInformationPanel = value

    @property
    def DisplayGridLines(self):
        return self.application.DisplayGridLines

    @DisplayGridLines.setter
    def DisplayGridLines(self, value):
        self.application.DisplayGridLines = value

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

    def FileDialog(self, Type=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
        ]
        if callable(self.application.FileDialog):
            return self.application.FileDialog(*params)
        else:
            return self.application.GetFileDialog(*params)

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

    @ShowStartupDialog.setter
    def ShowStartupDialog(self, value):
        self.application.ShowStartupDialog = value

    @property
    def ShowWindowsInTaskbar(self):
        return self.application.ShowWindowsInTaskbar

    @ShowWindowsInTaskbar.setter
    def ShowWindowsInTaskbar(self, value):
        self.application.ShowWindowsInTaskbar = value

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

    def Help(self, HelpFile=None, ContextID=None):
        params = [
            HelpFile if HelpFile is not None else pythoncom.Missing,
            ContextID if ContextID is not None else pythoncom.Missing,
        ]
        self.application.Help(*params)

    def Quit(self):
        self.application.Quit()

    def Run(self, MacroName=None, safeArrayOfParams=None):
        params = [
            MacroName if MacroName is not None else pythoncom.Missing,
            safeArrayOfParams if safeArrayOfParams is not None else pythoncom.Missing,
        ]
        return self.application.Run(*params)

    def StartNewUndoEntry(self):
        self.application.StartNewUndoEntry()


class AutoCorrect:

    def __init__(self, autocorrect=None):
        self.autocorrect = autocorrect

    @property
    def DisplayAutoCorrectOptions(self):
        return self.autocorrect.DisplayAutoCorrectOptions

    @DisplayAutoCorrectOptions.setter
    def DisplayAutoCorrectOptions(self, value):
        self.autocorrect.DisplayAutoCorrectOptions = value

    @property
    def DisplayAutoLayoutOptions(self):
        return self.autocorrect.DisplayAutoLayoutOptions

    @DisplayAutoLayoutOptions.setter
    def DisplayAutoLayoutOptions(self, value):
        self.autocorrect.DisplayAutoLayoutOptions = value


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
        params = [
            Type if Type is not None else pythoncom.Missing,
            AxisGroup if AxisGroup is not None else pythoncom.Missing,
        ]
        self.axes.Item(*params)


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
        params = [
            Start if Start is not None else pythoncom.Missing,
            Length if Length is not None else pythoncom.Missing,
        ]
        if callable(self.axistitle.Characters):
            return ChartCharacters(self.axistitle.Characters(*params))
        else:
            return ChartCharacters(self.axistitle.GetCharacters(*params))

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

    def Item(self, BorderType=None):
        params = [
            BorderType if BorderType is not None else pythoncom.Missing,
        ]
        return self.borders.Item(*params)


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

    def Start(self, serverUrl=None):
        params = [
            serverUrl if serverUrl is not None else pythoncom.Missing,
        ]
        self.broadcast.Start(*params)


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

    @Type.setter
    def Type(self, value):
        self.bulletformat.Type = value

    @property
    def UseTextColor(self):
        return self.bulletformat.UseTextColor

    @UseTextColor.setter
    def UseTextColor(self, value):
        self.bulletformat.UseTextColor = value

    @property
    def UseTextFont(self):
        return self.bulletformat.UseTextFont

    @UseTextFont.setter
    def UseTextFont(self, value):
        self.bulletformat.UseTextFont = value

    def Picture(self):
        self.bulletformat.Picture()


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
    def AutoAttach(self):
        return self.calloutformat.AutoAttach

    @AutoAttach.setter
    def AutoAttach(self, value):
        self.calloutformat.AutoAttach = value

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

    def AutomaticLength(self):
        self.calloutformat.AutomaticLength()

    def CustomDrop(self, Drop=None):
        params = [
            Drop if Drop is not None else pythoncom.Missing,
        ]
        return self.calloutformat.CustomDrop(*params)

    def CustomLength(self, Length=None):
        params = [
            Length if Length is not None else pythoncom.Missing,
        ]
        return self.calloutformat.CustomLength(*params)

    def PresetDrop(self, DropType=None):
        params = [
            DropType if DropType is not None else pythoncom.Missing,
        ]
        self.calloutformat.PresetDrop(*params)


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

    def Merge(self, MergeTo=None):
        params = [
            MergeTo if MergeTo is not None else pythoncom.Missing,
        ]
        self.cell.Merge(*params)

    def Select(self):
        self.cell.Select()

    def Split(self, NumRows=None, NumColumns=None):
        params = [
            NumRows if NumRows is not None else pythoncom.Missing,
            NumColumns if NumColumns is not None else pythoncom.Missing,
        ]
        self.cell.Split(*params)


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

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.cellrange.Item(*params)


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
    def Name(self):
        return self.chart.Name

    @Name.setter
    def Name(self, value):
        self.chart.Name = value

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

    @Title.setter
    def Title(self, value):
        self.chart.Title = value

    @property
    def Walls(self):
        return Walls(self.chart.Walls)

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
        self.chart.ChartGroups(*params)

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

    def Export(self, FileName=None, FilterName=None, Interactive=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            FilterName if FilterName is not None else pythoncom.Missing,
            Interactive if Interactive is not None else pythoncom.Missing,
        ]
        self.chart.Export(*params)

    def GetChartElement(self, x=None, y=None, ElementID=None, Arg1=None, Arg2=None):
        params = [
            x if x is not None else pythoncom.Missing,
            y if y is not None else pythoncom.Missing,
            ElementID if ElementID is not None else pythoncom.Missing,
            Arg1 if Arg1 is not None else pythoncom.Missing,
            Arg2 if Arg2 is not None else pythoncom.Missing,
        ]
        self.chart.GetChartElement(*params)

    def Paste(self, Type=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
        ]
        self.chart.Paste(*params)

    def Refresh(self):
        self.chart.Refresh()

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
        return SeriesCollection(self.chart.SeriesCollection(*params))

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
        self.chart.SetElement(*params)

    def SetSourceData(self, Source=None, PlotBy=None):
        params = [
            Source if Source is not None else pythoncom.Missing,
            PlotBy if PlotBy is not None else pythoncom.Missing,
        ]
        self.chart.SetSourceData(*params)


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

    def Insert(self, String=None):
        params = [
            String if String is not None else pythoncom.Missing,
        ]
        self.chartcharacters.Insert(*params)


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
        return SeriesCollection(self.chartgroup.SeriesCollection(*params))


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
            return ChartCharacters(self.charttitle.Characters(*params))
        else:
            return ChartCharacters(self.charttitle.GetCharacters(*params))

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

    @FavorServerEditsDuringMerge.setter
    def FavorServerEditsDuringMerge(self, value):
        self.coauthoring.FavorServerEditsDuringMerge = value

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

    @To.setter
    def To(self, value):
        self.coloreffect.To = value


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

    @TintAndShade.setter
    def TintAndShade(self, value):
        self.colorformat.TintAndShade = value

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

    def Colors(self, SchemeColor=None):
        params = [
            SchemeColor if SchemeColor is not None else pythoncom.Missing,
        ]
        return self.colorscheme.Colors(*params)

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

    def Add(self, Scheme=None):
        params = [
            Scheme if Scheme is not None else pythoncom.Missing,
        ]
        return ColorScheme(self.colorschemes.Add(*params))

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.colorschemes.Item(*params)


class Column:

    def __init__(self, column=None):
        self.column = column

    @property
    def Application(self):
        return Application(self.column.Application)

    def Cells(self, RowIndex=None, ColumnIndex=None):
        params = [
            RowIndex if RowIndex is not None else pythoncom.Missing,
            ColumnIndex if ColumnIndex is not None else pythoncom.Missing,
        ]
        if callable(self.column.Cells):
            return CellRange(self.column.Cells(*params))
        else:
            return CellRange(self.column.GetCells(*params))

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

    def Add(self, BeforeColumn=None):
        params = [
            BeforeColumn if BeforeColumn is not None else pythoncom.Missing,
        ]
        return Column(self.columns.Add(*params))

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.columns.Item(*params)


class CommandEffect:

    def __init__(self, commandeffect=None):
        self.commandeffect = commandeffect

    @property
    def Application(self):
        return Application(self.commandeffect.Application)

    @property
    def Bookmark(self):
        return self.commandeffect.Bookmark

    @Bookmark.setter
    def Bookmark(self, value):
        self.commandeffect.Bookmark = value

    @property
    def Command(self):
        return self.commandeffect.Command

    @Command.setter
    def Command(self, value):
        self.commandeffect.Command = value

    @property
    def Parent(self):
        return self.commandeffect.Parent

    @property
    def Type(self):
        return self.commandeffect.Type

    @Type.setter
    def Type(self, value):
        self.commandeffect.Type = value


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

    def Add(self, Left=None, Top=None, Author=None, AuthorInitials=None, Text=None):
        params = [
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            Author if Author is not None else pythoncom.Missing,
            AuthorInitials if AuthorInitials is not None else pythoncom.Missing,
            Text if Text is not None else pythoncom.Missing,
        ]
        return self.comments.Add(*params)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.comments.Item(*params)


class ConnectorFormat:

    def __init__(self, connectorformat=None):
        self.connectorformat = connectorformat

    @property
    def Application(self):
        return Application(self.connectorformat.Application)

    @property
    def BeginConnected(self):
        return self.connectorformat.BeginConnected

    @BeginConnected.setter
    def BeginConnected(self, value):
        self.connectorformat.BeginConnected = value

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

    def Delete(self, Id=None):
        params = [
            Id if Id is not None else pythoncom.Missing,
        ]
        self.customerdata.Delete(*params)

    def Item(self, Id=None):
        params = [
            Id if Id is not None else pythoncom.Missing,
        ]
        return self.customerdata.Item(*params)


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

    @DisplayMasterShapes.setter
    def DisplayMasterShapes(self, value):
        self.customlayout.DisplayMasterShapes = value

    @property
    def FollowMasterBackground(self):
        return self.customlayout.FollowMasterBackground

    @FollowMasterBackground.setter
    def FollowMasterBackground(self, value):
        self.customlayout.FollowMasterBackground = value

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

    @MatchingName.setter
    def MatchingName(self, value):
        self.customlayout.MatchingName = value

    @property
    def Name(self):
        return self.customlayout.Name

    @Name.setter
    def Name(self, value):
        self.customlayout.Name = value

    @property
    def Parent(self):
        return CustomLayout(self.customlayout.Parent)

    @property
    def Preserved(self):
        return self.customlayout.Preserved

    @Preserved.setter
    def Preserved(self, value):
        self.customlayout.Preserved = value

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

    def MoveTo(self, toPos=None):
        params = [
            toPos if toPos is not None else pythoncom.Missing,
        ]
        self.customlayout.MoveTo(*params)

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

    def Add(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.customlayouts.Add(*params)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.customlayouts.Item(*params)

    def Paste(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.customlayouts.Paste(*params)


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
            return ChartCharacters(self.datalabel.Characters(*params))
        else:
            return ChartCharacters(self.datalabel.GetCharacters(*params))

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
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return DataLabel(self.datalabels.Item(*params))

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

    @Preserved.setter
    def Preserved(self, value):
        self.design.Preserved = value

    @property
    def SlideMaster(self):
        return Master(self.design.SlideMaster)

    def Delete(self):
        self.design.Delete()

    def MoveTo(self, toPos=None):
        params = [
            toPos if toPos is not None else pythoncom.Missing,
        ]
        self.design.MoveTo(*params)


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

    def Add(self, designName=None, Index=None):
        params = [
            designName if designName is not None else pythoncom.Missing,
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.designs.Add(*params)

    def Clone(self, pOriginal=None, Index=None):
        params = [
            pOriginal if pOriginal is not None else pythoncom.Missing,
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.designs.Clone(*params)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.designs.Item(*params)

    def Load(self, TemplateName=None, Index=None):
        params = [
            TemplateName if TemplateName is not None else pythoncom.Missing,
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.designs.Load(*params)


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
            return ChartCharacters(self.displayunitlabel.Characters(*params))
        else:
            return ChartCharacters(self.displayunitlabel.GetCharacters(*params))

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

    @BlackAndWhite.setter
    def BlackAndWhite(self, value):
        self.documentwindow.BlackAndWhite = value

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

    def ExpandSection(self, sectionIndex=None, Expand=None):
        params = [
            sectionIndex if sectionIndex is not None else pythoncom.Missing,
            Expand if Expand is not None else pythoncom.Missing,
        ]
        self.documentwindow.ExpandSection(*params)

    def FitToPage(self):
        self.documentwindow.FitToPage()

    def IsSectionExpanded(self, sectionIndex=None):
        params = [
            sectionIndex if sectionIndex is not None else pythoncom.Missing,
        ]
        return self.documentwindow.IsSectionExpanded(*params)

    def LargeScroll(self, Down=None, Up=None, ToRight=None, ToLeft=None):
        params = [
            Down if Down is not None else pythoncom.Missing,
            Up if Up is not None else pythoncom.Missing,
            ToRight if ToRight is not None else pythoncom.Missing,
            ToLeft if ToLeft is not None else pythoncom.Missing,
        ]
        self.documentwindow.LargeScroll(*params)

    def NewWindow(self):
        return self.documentwindow.NewWindow()

    def PointsToScreenPixelsX(self, Points=None):
        params = [
            Points if Points is not None else pythoncom.Missing,
        ]
        return self.documentwindow.PointsToScreenPixelsX(*params)

    def PointsToScreenPixelsY(self, Points=None):
        params = [
            Points if Points is not None else pythoncom.Missing,
        ]
        return self.documentwindow.PointsToScreenPixelsY(*params)

    def RangeFromPoint(self, x=None, y=None):
        params = [
            x if x is not None else pythoncom.Missing,
            y if y is not None else pythoncom.Missing,
        ]
        self.documentwindow.RangeFromPoint(*params)

    def ScrollIntoView(self, Left=None, Top=None, Width=None, Height=None, Start=None):
        params = [
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            Width if Width is not None else pythoncom.Missing,
            Height if Height is not None else pythoncom.Missing,
            Start if Start is not None else pythoncom.Missing,
        ]
        self.documentwindow.ScrollIntoView(*params)

    def SmallScroll(self, Down=None, Up=None, ToRight=None, ToLeft=None):
        params = [
            Down if Down is not None else pythoncom.Missing,
            Up if Up is not None else pythoncom.Missing,
            ToRight if ToRight is not None else pythoncom.Missing,
            ToLeft if ToLeft is not None else pythoncom.Missing,
        ]
        self.documentwindow.SmallScroll(*params)


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

    def Arrange(self, arrangeStyle=None):
        params = [
            arrangeStyle if arrangeStyle is not None else pythoncom.Missing,
        ]
        return self.documentwindows.Arrange(*params)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.documentwindows.Item(*params)


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

    @EffectType.setter
    def EffectType(self, value):
        self.effect.EffectType = value

    @property
    def Exit(self):
        return self.effect.Exit

    @Exit.setter
    def Exit(self, value):
        self.effect.Exit = value

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

    def MoveAfter(self, Effect=None):
        params = [
            Effect if Effect is not None else pythoncom.Missing,
        ]
        self.effect.MoveAfter(*params)

    def MoveBefore(self, Effect=None):
        params = [
            Effect if Effect is not None else pythoncom.Missing,
        ]
        self.effect.MoveBefore(*params)

    def MoveTo(self, toPos=None):
        params = [
            toPos if toPos is not None else pythoncom.Missing,
        ]
        self.effect.MoveTo(*params)


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

    @AnimateTextInReverse.setter
    def AnimateTextInReverse(self, value):
        self.effectinformation.AnimateTextInReverse = value

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

    @Direction.setter
    def Direction(self, value):
        self.effectparameters.Direction = value

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

    @Relative.setter
    def Relative(self, value):
        self.effectparameters.Relative = value

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

    def Add(self, Type=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
        ]
        self.extracolors.Add(*params)

    def Clear(self):
        self.extracolors.Clear()

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return MsoThemeColorSchemeIndex(self.extracolors.Item(*params))


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

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.fileconverters.Item(*params)


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

    @Reveal.setter
    def Reveal(self, value):
        self.filtereffect.Reveal = value

    @property
    def Subtype(self):
        return self.filtereffect.Subtype

    @Subtype.setter
    def Subtype(self, value):
        self.filtereffect.Subtype = value

    @property
    def Type(self):
        return self.filtereffect.Type

    @Type.setter
    def Type(self, value):
        self.filtereffect.Type = value


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

    @Bold.setter
    def Bold(self, value):
        self.font.Bold = value

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

    @Emboss.setter
    def Emboss(self, value):
        self.font.Emboss = value

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
    def Underline(self):
        return self.font.Underline

    @Underline.setter
    def Underline(self, value):
        self.font.Underline = value


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

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.fonts.Item(*params)

    def Replace(self, Original=None, Replacement=None):
        params = [
            Original if Original is not None else pythoncom.Missing,
            Replacement if Replacement is not None else pythoncom.Missing,
        ]
        self.fonts.Replace(*params)


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

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.groupshapes.Item(*params)

    def Range(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.groupshapes.Range(*params)


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

    @UseFormat.setter
    def UseFormat(self, value):
        self.headerfooter.UseFormat = value

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

    @DisplayOnTitleSlide.setter
    def DisplayOnTitleSlide(self, value):
        self.headersfooters.DisplayOnTitleSlide = value

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

    @ShowAndReturn.setter
    def ShowAndReturn(self, value):
        self.hyperlink.ShowAndReturn = value

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
        return self.hyperlink.CreateNewDocument(*params)

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

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.hyperlinks.Item(*params)


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

    @BackgroundStyle.setter
    def BackgroundStyle(self, value):
        self.master.BackgroundStyle = value

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

    def ApplyTheme(self, themeName=None):
        params = [
            themeName if themeName is not None else pythoncom.Missing,
        ]
        self.master.ApplyTheme(*params)

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

    def Add(self, Position=None, Name=None):
        params = [
            Position if Position is not None else pythoncom.Missing,
            Name if Name is not None else pythoncom.Missing,
        ]
        return MediaBookmark(self.mediabookmarks.Add(*params))

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.mediabookmarks.Item(*params)


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

    @EndPoint.setter
    def EndPoint(self, value):
        self.mediaformat.EndPoint = value

    @property
    def FadeInDuration(self):
        return self.mediaformat.FadeInDuration

    @FadeInDuration.setter
    def FadeInDuration(self, value):
        self.mediaformat.FadeInDuration = value

    @property
    def FadeOutDuration(self):
        return self.mediaformat.FadeOutDuration

    @FadeOutDuration.setter
    def FadeOutDuration(self, value):
        self.mediaformat.FadeOutDuration = value

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

    @Muted.setter
    def Muted(self, value):
        self.mediaformat.Muted = value

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

    @StartPoint.setter
    def StartPoint(self, value):
        self.mediaformat.StartPoint = value

    @property
    def VideoCompressionType(self):
        return self.mediaformat.VideoCompressionType

    @property
    def VideoFrameRate(self):
        return self.mediaformat.VideoFrameRate

    @property
    def Volume(self):
        return self.mediaformat.Volume

    @Volume.setter
    def Volume(self, value):
        self.mediaformat.Volume = value

    def Resample(self, Trim=None, SampleHeight=None, SampleWidth=None, VideoFrameRate=None, AudioSamplingRate=None, VideoBitRate=None):
        params = [
            Trim if Trim is not None else pythoncom.Missing,
            SampleHeight if SampleHeight is not None else pythoncom.Missing,
            SampleWidth if SampleWidth is not None else pythoncom.Missing,
            VideoFrameRate if VideoFrameRate is not None else pythoncom.Missing,
            AudioSamplingRate if AudioSamplingRate is not None else pythoncom.Missing,
            VideoBitRate if VideoBitRate is not None else pythoncom.Missing,
        ]
        return self.mediaformat.Resample(*params)

    def ResampleFromProfile(self, profile=None):
        params = [
            profile if profile is not None else pythoncom.Missing,
        ]
        return self.mediaformat.ResampleFromProfile(*params)

    def SetDisplayPicture(self, Position=None):
        params = [
            Position if Position is not None else pythoncom.Missing,
        ]
        return self.mediaformat.SetDisplayPicture(*params)

    def SetDisplayPictureFromFile(self, FilePath=None):
        params = [
            FilePath if FilePath is not None else pythoncom.Missing,
        ]
        return self.mediaformat.SetDisplayPictureFromFile(*params)


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


class MotionEffect:

    def __init__(self, motioneffect=None):
        self.motioneffect = motioneffect

    @property
    def Application(self):
        return Application(self.motioneffect.Application)

    @property
    def ByX(self):
        return self.motioneffect.ByX

    @ByX.setter
    def ByX(self, value):
        self.motioneffect.ByX = value

    @property
    def ByY(self):
        return self.motioneffect.ByY

    @ByY.setter
    def ByY(self, value):
        self.motioneffect.ByY = value

    @property
    def FromX(self):
        return self.motioneffect.FromX

    @FromX.setter
    def FromX(self, value):
        self.motioneffect.FromX = value

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

    @Path.setter
    def Path(self, value):
        self.motioneffect.Path = value

    @property
    def ToX(self):
        return self.motioneffect.ToX

    @ToX.setter
    def ToX(self, value):
        self.motioneffect.ToX = value

    @property
    def ToY(self):
        return MotionEffect(self.motioneffect.ToY)

    @ToY.setter
    def ToY(self, value):
        self.motioneffect.ToY = value


# MsoAnimAccumulate enumeration
msoAnimAccumulateAlways = 2
msoAnimAccumulateNone = 1

# MsoAnimAdditive enumeration
msoAnimAdditiveAddBase = 1
msoAnimAdditiveAddSum = 2

# MsoAnimAfterEffect enumeration
msoAnimAfterEffectDim = 1
msoAnimAfterEffectHide = 2
msoAnimAfterEffectHideOnNextClick = 3
msoAnimAfterEffectMixed = -1
msoAnimAfterEffectNone = 0

# MsoAnimateByLevel enumeration
msoAnimateChartAllAtOnce = 7
msoAnimateChartByCategory = 8
msoAnimateChartByCategoryElements = 9
msoAnimateChartBySeries = 10
msoAnimateChartBySeriesElements = 11
msoAnimateDiagramAllAtOnce = 12
msoAnimateDiagramBreadthByLevel = 16
msoAnimateDiagramBreadthByNode = 15
msoAnimateDiagramClockwise = 17
msoAnimateDiagramClockwiseIn = 18
msoAnimateDiagramClockwiseOut = 19
msoAnimateDiagramCounterClockwise = 20
msoAnimateDiagramCounterClockwiseIn = 21
msoAnimateDiagramCounterClockwiseOut = 22
msoAnimateDiagramDepthByBranch = 14
msoAnimateDiagramDepthByNode = 13
msoAnimateDiagramDown = 26
msoAnimateDiagramInByRing = 23
msoAnimateDiagramOutByRing = 24
msoAnimateDiagramUp = 25
msoAnimateLevelMixed = -1
msoAnimateLevelNone = 0
msoAnimateTextByAllLevels = 1
msoAnimateTextByFifthLevel = 6
msoAnimateTextByFirstLevel = 2
msoAnimateTextByFourthLevel = 5
msoAnimateTextBySecondLevel = 3
msoAnimateTextByThirdLevel = 4

# MsoAnimCommandType enumeration
msoAnimCommandTypeCall = 1
msoAnimCommandTypeEvent = 0
msoAnimCommandTypeVerb = 2

# MsoAnimDirection enumeration
msoAnimDirectionAcross = 18
msoAnimDirectionBottom = 11
msoAnimDirectionBottomLeft = 15
msoAnimDirectionBottomRight = 14
msoAnimDirectionCenter = 28
msoAnimDirectionClockwise = 21
msoAnimDirectionCounterclockwise = 22
msoAnimDirectionCycleClockwise = 43
msoAnimDirectionCycleCounterclockwise = 44
msoAnimDirectionDown = 3
msoAnimDirectionDownLeft = 9
msoAnimDirectionDownRight = 8
msoAnimDirectionFontAllCaps = 40
msoAnimDirectionFontBold = 35
msoAnimDirectionFontItalic = 36
msoAnimDirectionFontShadow = 39
msoAnimDirectionFontStrikethrough = 38
msoAnimDirectionFontUnderline = 37
msoAnimDirectionGradual = 42
msoAnimDirectionHorizontal = 16
msoAnimDirectionHorizontalIn = 23
msoAnimDirectionHorizontalOut = 24
msoAnimDirectionIn = 19
msoAnimDirectionInBottom = 31
msoAnimDirectionInCenter = 30
msoAnimDirectionInSlightly = 29
msoAnimDirectionInstant = 41
msoAnimDirectionLeft = 4
msoAnimDirectionNone = 0
msoAnimDirectionOrdinalMask = 5
msoAnimDirectionOut = 20
msoAnimDirectionOutBottom = 34
msoAnimDirectionOutCenter = 33
msoAnimDirectionOutSlightly = 32
msoAnimDirectionRight = 2
msoAnimDirectionSlightly = 27
msoAnimDirectionTop = 10
msoAnimDirectionTopLeft = 12
msoAnimDirectionTopRight = 13
msoAnimDirectionUp = 1
msoAnimDirectionUpLeft = 6
msoAnimDirectionUpRight = 7
msoAnimDirectionVertical = 17
msoAnimDirectionVerticalIn = 25
msoAnimDirectionVerticalOut = 26

# MsoAnimEffect enumeration
msoAnimEffectAppear = 1
msoAnimEffectArcUp = 47
msoAnimEffectAscend = 39
msoAnimEffectBlast = 64
msoAnimEffectBlinds = 3
msoAnimEffectBoldFlash = 63
msoAnimEffectBoldReveal = 65
msoAnimEffectBoomerang = 25
msoAnimEffectBounce = 26
msoAnimEffectBox = 4
msoAnimEffectBrushOnColor = 66
msoAnimEffectBrushOnUnderline = 67
msoAnimEffectCenterRevolve = 40
msoAnimEffectChangeFillColor = 54
msoAnimEffectChangeFont = 55
msoAnimEffectChangeFontColor = 56
msoAnimEffectChangeFontSize = 57
msoAnimEffectChangeFontStyle = 58
msoAnimEffectChangeLineColor = 60
msoAnimEffectCheckerboard = 5
msoAnimEffectCircle = 6
msoAnimEffectColorBlend = 68
msoAnimEffectColorReveal = 27
msoAnimEffectColorWave = 69
msoAnimEffectComplementaryColor = 70
msoAnimEffectComplementaryColor2 = 71
msoAnimEffectContrastingColor = 72
msoAnimEffectCrawl = 7
msoAnimEffectCredits = 28
msoAnimEffectCustom = 0
msoAnimEffectDarken = 73
msoAnimEffectDesaturate = 74
msoAnimEffectDescend = 42
msoAnimEffectDiamond = 8
msoAnimEffectDissolve = 9
msoAnimEffectEaseIn = 29
msoAnimEffectExpand = 50
msoAnimEffectFade = 10
msoAnimEffectFadedSwivel = 41
msoAnimEffectFadedZoom = 48
msoAnimEffectFlashBulb = 75
msoAnimEffectFlashOnce = 11
msoAnimEffectFlicker = 76
msoAnimEffectFlip = 51
msoAnimEffectFloat = 30
msoAnimEffectFly = 2
msoAnimEffectFold = 53
msoAnimEffectGlide = 49
msoAnimEffectGrowAndTurn = 31
msoAnimEffectGrowShrink = 59
msoAnimEffectGrowWithColor = 77
msoAnimEffectLighten = 78
msoAnimEffectLightSpeed = 32
msoAnimEffectMediaPause = 84
msoAnimEffectMediaPlay = 83
msoAnimEffectMediaStop = 85
msoAnimEffectPath4PointStar = 101
msoAnimEffectPath5PointStar = 90
msoAnimEffectPath6PointStar = 96
msoAnimEffectPath8PointStar = 102
msoAnimEffectPathArcDown = 122
msoAnimEffectPathArcLeft = 136
msoAnimEffectPathArcRight = 143
msoAnimEffectPathArcUp = 129
msoAnimEffectPathBean = 116
msoAnimEffectPathBounceLeft = 126
msoAnimEffectPathBounceRight = 139
msoAnimEffectPathBuzzsaw = 110
msoAnimEffectPathCircle = 86
msoAnimEffectPathCrescentMoon = 91
msoAnimEffectPathCurvedSquare = 105
msoAnimEffectPathCurvedX = 106
msoAnimEffectPathCurvyLeft = 133
msoAnimEffectPathCurvyRight = 146
msoAnimEffectPathCurvyStar = 108
msoAnimEffectPathDecayingWave = 145
msoAnimEffectPathDiagonalDownRight = 134
msoAnimEffectPathDiagonalUpRight = 141
msoAnimEffectPathDiamond = 88
msoAnimEffectPathDown = 127
msoAnimEffectPathEqualTriangle = 98
msoAnimEffectPathFigure8Four = 113
msoAnimEffectPathFootball = 97
msoAnimEffectPathFunnel = 137
msoAnimEffectPathHeart = 94
msoAnimEffectPathHeartbeat = 130
msoAnimEffectPathHexagon = 89
msoAnimEffectPathHorizontalFigure8 = 111
msoAnimEffectPathInvertedSquare = 119
msoAnimEffectPathInvertedTriangle = 118
msoAnimEffectPathLeft = 120
msoAnimEffectPathLoopdeLoop = 109
msoAnimEffectPathNeutron = 114
msoAnimEffectPathOctagon = 95
msoAnimEffectPathParallelogram = 99
msoAnimEffectPathPeanut = 112
msoAnimEffectPathPentagon = 100
msoAnimEffectPathPlus = 117
msoAnimEffectPathPointyStar = 104
msoAnimEffectPathRight = 149
msoAnimEffectPathRightTriangle = 87
msoAnimEffectPathSCurve1 = 144
msoAnimEffectPathSCurve2 = 124
msoAnimEffectPathSineWave = 125
msoAnimEffectPathSpiralLeft = 140
msoAnimEffectPathSpiralRight = 131
msoAnimEffectPathSpring = 138
msoAnimEffectPathSquare = 92
msoAnimEffectPathStairsDown = 147
msoAnimEffectPathSwoosh = 115
msoAnimEffectPathTeardrop = 103
msoAnimEffectPathTrapezoid = 93
msoAnimEffectPathTurnDown = 135
msoAnimEffectPathTurnRight = 121
msoAnimEffectPathTurnUp = 128
msoAnimEffectPathTurnUpRight = 142
msoAnimEffectPathUp = 148
msoAnimEffectPathVerticalFigure8 = 107
msoAnimEffectPathWave = 132
msoAnimEffectPathZigzag = 123
msoAnimEffectPeek = 12
msoAnimEffectPinwheel = 33
msoAnimEffectPlus = 13
msoAnimEffectRandomBars = 14
msoAnimEffectRandomEffects = 24
msoAnimEffectRiseUp = 34
msoAnimEffectShimmer = 52
msoAnimEffectSling = 43
msoAnimEffectSpin = 61
msoAnimEffectSpinner = 44
msoAnimEffectSpiral = 15
msoAnimEffectSplit = 16
msoAnimEffectStretch = 17
msoAnimEffectStretchy = 45
msoAnimEffectStrips = 18
msoAnimEffectStyleEmphasis = 79
msoAnimEffectSwish = 35
msoAnimEffectSwivel = 19
msoAnimEffectTeeter = 80
msoAnimEffectThinLine = 36
msoAnimEffectTransparency = 62
msoAnimEffectUnfold = 37
msoAnimEffectVerticalGrow = 81
msoAnimEffectWave = 82
msoAnimEffectWedge = 20
msoAnimEffectWheel = 21
msoAnimEffectWhip = 38
msoAnimEffectWipe = 22
msoAnimEffectZip = 46
msoAnimEffectZoom = 23

# MsoAnimEffectAfter enumeration
msoAnimEffectAfterFreeze = 1
msoAnimEffectAfterHold = 3
msoAnimEffectAfterRemove = 2
msoAnimEffectAfterTransition = 4

# MsoAnimEffectRestart enumeration
msoAnimEffectRestartAlways = 1
msoAnimEffectRestartNever = 3
msoAnimEffectRestartWhenOff = 2

# MsoAnimFilterEffectSubtype enumeration
msoAnimFilterEffectSubtypeAcross = 9
msoAnimFilterEffectSubtypeDown = 25
msoAnimFilterEffectSubtypeDownLeft = 14
msoAnimFilterEffectSubtypeDownRight = 16
msoAnimFilterEffectSubtypeFromBottom = 13
msoAnimFilterEffectSubtypeFromLeft = 10
msoAnimFilterEffectSubtypeFromRight = 11
msoAnimFilterEffectSubtypeFromTop = 12
msoAnimFilterEffectSubtypeHorizontal = 5
msoAnimFilterEffectSubtypeIn = 7
msoAnimFilterEffectSubtypeInHorizontal = 3
msoAnimFilterEffectSubtypeInVertical = 1
msoAnimFilterEffectSubtypeLeft = 23
msoAnimFilterEffectSubtypeNone = 0
msoAnimFilterEffectSubtypeOut = 8
msoAnimFilterEffectSubtypeOutHorizontal = 4
msoAnimFilterEffectSubtypeOutVertical = 2
msoAnimFilterEffectSubtypeRight = 24
msoAnimFilterEffectSubtypeSpokes1 = 18
msoAnimFilterEffectSubtypeSpokes2 = 19
msoAnimFilterEffectSubtypeSpokes3 = 20
msoAnimFilterEffectSubtypeSpokes4 = 21
msoAnimFilterEffectSubtypeSpokes8 = 22
msoAnimFilterEffectSubtypeUp = 26
msoAnimFilterEffectSubtypeUpLeft = 15
msoAnimFilterEffectSubtypeUpRight = 17
msoAnimFilterEffectSubtypeVertical = 6

# MsoAnimFilterEffectType enumeration
msoAnimFilterEffectTypeBarn = 1
msoAnimFilterEffectTypeBlinds = 2
msoAnimFilterEffectTypeBox = 3
msoAnimFilterEffectTypeCheckerboard = 4
msoAnimFilterEffectTypeCircle = 5
msoAnimFilterEffectTypeDiamond = 6
msoAnimFilterEffectTypeDissolve = 7
msoAnimFilterEffectTypeFade = 8
msoAnimFilterEffectTypeImage = 9
msoAnimFilterEffectTypeNone = 0
msoAnimFilterEffectTypePixelate = 10
msoAnimFilterEffectTypePlus = 11
msoAnimFilterEffectTypeRandomBar = 12
msoAnimFilterEffectTypeSlide = 13
msoAnimFilterEffectTypeStretch = 14
msoAnimFilterEffectTypeStrips = 15
msoAnimFilterEffectTypeWedge = 16
msoAnimFilterEffectTypeWheel = 17
msoAnimFilterEffectTypeWipe = 18

# MsoAnimProperty enumeration
msoAnimColor = 7
msoAnimHeight = 4
msoAnimNone = 0
msoAnimOpacity = 5
msoAnimRotation = 6
msoAnimShapeFillBackColor = 1007
msoAnimShapeFillColor = 1005
msoAnimShapeFillOn = 1004
msoAnimShapeFillOpacity = 1006
msoAnimShapeLineColor = 1009
msoAnimShapeLineOn = 1008
msoAnimShapePictureBrightness = 1001
msoAnimShapePictureContrast = 1000
msoAnimShapePictureGamma = 1002
msoAnimShapePictureGrayscale = 1003
msoAnimShapeShadowColor = 1012
msoAnimShapeShadowOffsetX = 1014
msoAnimShapeShadowOffsetY = 1015
msoAnimShapeShadowOn = 1010
msoAnimShapeShadowOpacity = 1013
msoAnimShapeShadowType = 1011
msoAnimTextBulletCharacter = 111
msoAnimTextBulletColor = 114
msoAnimTextBulletFontName = 112
msoAnimTextBulletNumber = 113
msoAnimTextBulletRelativeSize = 115
msoAnimTextBulletStyle = 116
msoAnimTextBulletType = 117
msoAnimTextFontBold = 100
msoAnimTextFontColor = 101
msoAnimTextFontEmboss = 102
msoAnimTextFontItalic = 103
msoAnimTextFontName = 104
msoAnimTextFontShadow = 105
msoAnimTextFontSize = 106
msoAnimTextFontStrikeThrough = 110
msoAnimTextFontSubscript = 107
msoAnimTextFontSuperscript = 108
msoAnimTextFontUnderline = 109
msoAnimVisibility = 8
msoAnimWidth = 3
msoAnimX = 1
msoAnimY = 2

# MsoAnimTextUnitEffect enumeration
msoAnimTextUnitEffectByCharacter = 1
msoAnimTextUnitEffectByParagraph = 0
msoAnimTextUnitEffectByWord = 2
msoAnimTextUnitEffectMixed = -1

# MsoAnimTriggerType enumeration
msoAnimTriggerAfterPrevious = 3
msoAnimTriggerMixed = -1
msoAnimTriggerNone = 0
msoAnimTriggerOnPageClick = 1
msoAnimTriggerOnShapeClick = 4
msoAnimTriggerWithPrevious = 2
msoAnimTriggerOnMediaBookmark = 5

# MsoAnimType enumeration
msoAnimTypeColor = 2
msoAnimTypeCommand = 6
msoAnimTypeFilter = 7
msoAnimTypeMixed = -2
msoAnimTypeMotion = 1
msoAnimTypeNone = 0
msoAnimTypeProperty = 5
msoAnimTypeRotation = 4
msoAnimTypeScale = 3
msoAnimTypeSet = 8

# MsoClickState enumeration
msoClickStateAfterAllAnimations = -2
msoClickStateBeforeAutomaticAnimations = -1

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

    def Add(self, Name=None, SafeArrayOfSlideIDs=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
            SafeArrayOfSlideIDs if SafeArrayOfSlideIDs is not None else pythoncom.Missing,
        ]
        return NamedSlideShow(self.namedslideshows.Add(*params))

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.namedslideshows.Item(*params)


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

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.objectverbs.Item(*params)


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

    def DoVerb(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        self.oleformat.DoVerb(*params)


class Options:

    def __init__(self, options=None):
        self.options = options

    @property
    def DisplayPasteOptions(self):
        return self.options.DisplayPasteOptions

    @DisplayPasteOptions.setter
    def DisplayPasteOptions(self, value):
        self.options.DisplayPasteOptions = value

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

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.panes.Item(*params)


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

    @LineRuleAfter.setter
    def LineRuleAfter(self, value):
        self.paragraphformat.LineRuleAfter = value

    @property
    def LineRuleBefore(self):
        return self.paragraphformat.LineRuleBefore

    @LineRuleBefore.setter
    def LineRuleBefore(self, value):
        self.paragraphformat.LineRuleBefore = value

    @property
    def LineRuleWithin(self):
        return self.paragraphformat.LineRuleWithin

    @LineRuleWithin.setter
    def LineRuleWithin(self, value):
        self.paragraphformat.LineRuleWithin = value

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

    @WordWrap.setter
    def WordWrap(self, value):
        self.paragraphformat.WordWrap = value


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
        params = [
            Increment if Increment is not None else pythoncom.Missing,
        ]
        self.pictureformat.IncrementBrightness(*params)

    def IncrementContrast(self, Increment=None):
        params = [
            Increment if Increment is not None else pythoncom.Missing,
        ]
        self.pictureformat.IncrementContrast(*params)


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

    def FindByName(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.placeholders.FindByName(*params)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.placeholders.Item(*params)


class Player:

    def __init__(self, player=None):
        self.player = player

    @property
    def Application(self):
        return Application(self.player.Application)

    @property
    def CurrentPosition(self):
        return self.player.CurrentPosition

    @CurrentPosition.setter
    def CurrentPosition(self, value):
        self.player.CurrentPosition = value

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

    @HideWhileNotPlaying.setter
    def HideWhileNotPlaying(self, value):
        self.playsettings.HideWhileNotPlaying = value

    @property
    def LoopUntilStopped(self):
        return self.playsettings.LoopUntilStopped

    @LoopUntilStopped.setter
    def LoopUntilStopped(self, value):
        self.playsettings.LoopUntilStopped = value

    @property
    def Parent(self):
        return self.playsettings.Parent

    @property
    def PauseAnimation(self):
        return self.playsettings.PauseAnimation

    @PauseAnimation.setter
    def PauseAnimation(self, value):
        self.playsettings.PauseAnimation = value

    @property
    def PlayOnEntry(self):
        return self.playsettings.PlayOnEntry

    @PlayOnEntry.setter
    def PlayOnEntry(self, value):
        self.playsettings.PlayOnEntry = value

    @property
    def RewindMovie(self):
        return self.playsettings.RewindMovie

    @RewindMovie.setter
    def RewindMovie(self, value):
        self.playsettings.RewindMovie = value

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
        self.point.ClearFormats()

    def Copy(self):
        self.point.Copy()

    def Delete(self):
        self.point.Delete()

    def Paste(self):
        self.point.Paste()

    def PieSliceLocation(self, loc=None, Index=None):
        params = [
            loc if loc is not None else pythoncom.Missing,
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.point.PieSliceLocation(*params)

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
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return Point(self.points.Item(*params))


# PpActionType enumeration
ppActionEndShow = 6
ppActionFirstSlide = 3
ppActionHyperlink = 7
ppActionLastSlide = 4
ppActionLastSlideViewed = 5
ppActionMixed = -2
ppActionNamedSlideShow = 10
ppActionNextSlide = 1
ppActionNone = 0
ppActionOLEVerb = 11
ppActionPlay = 12
ppActionPreviousSlide = 2
ppActionRunMacro = 8
ppActionRunProgram = 9

# PpAdvanceMode enumeration
ppAdvanceModeMixed = -2
ppAdvanceOnClick = 1
ppAdvanceOnTime = 2

# PpAfterEffect enumeration
ppAfterEffectDim = 2
ppAfterEffectHide = 1
ppAfterEffectHideOnClick = 3
ppAfterEffectMixed = -2
ppAfterEffectNothing = 0

# PpAlertLevel enumeration
ppAlertsAll = 2
ppAlertsNone = 1

# PpArrangeStyle enumeration
ppArrangeCascade = 2
ppArrangeTiled = 1

# PpAutoSize enumeration
ppAutoSizeMixed = -2
ppAutoSizeNone = 0
ppAutoSizeShapeToFitText = 1

# PpBaselineAlignment enumeration
ppBaselineAlignBaseline = 1
ppBaselineAlignCenter = 3
ppBaselineAlignFarEast50 = 4
ppBaselineAlignMixed = -2
ppBaselineAlignTop = 2

# PpBorderType enumeration
ppBorderBottom = 3
ppBorderDiagonalDown = 5
ppBorderDiagonalUp = 6
ppBorderLeft = 2
ppBorderRight = 4
ppBorderTop = 1

# PpBulletType enumeration
ppBulletMixed = -2
ppBulletNone = 0
ppBulletNumbered = 2
ppBulletPicture = 3
ppBulletUnnumbered = 1

# PpChangeCase enumeration
ppCaseLower = 2
ppCaseSentence = 1
ppCaseTitle = 4
ppCaseToggle = 5
ppCaseUpper = 3

# PpChartUnitEffect enumeration
ppAnimateByCategory = 2
ppAnimateByCategoryElements = 4
ppAnimateBySeries = 1
ppAnimateBySeriesElements = 3
ppAnimateChartAllAtOnce = 5
ppAnimateChartMixed = -2

# PpCheckInVersionType enumeration
ppCheckInMajorVersion = 1
ppCheckInMinorVersion = 0
ppCheckInOverwriteVersion = 2

# PpColorSchemeIndex enumeration
ppAccent1 = 6
ppAccent2 = 7
ppAccent3 = 8
ppBackground = 1
ppFill = 5
ppForeground = 2
ppNotSchemeColor = 0
ppSchemeColorMixed = -2
ppShadow = 3
ppTitle = 4

# PpDateTimeFormat enumeration
ppDateTimeddddMMMMddyyyy = 2
ppDateTimedMMMMyyyy = 3
ppDateTimedMMMyy = 5
ppDateTimeFigureOut = 14
ppDateTimeFormatMixed = -2
ppDateTimeHmm = 10
ppDateTimehmmAMPM = 12
ppDateTimeHmmss = 11
ppDateTimehmmssAMPM = 13
ppDateTimeMdyy = 1
ppDateTimeMMddyyHmm = 8
ppDateTimeMMddyyhmmAMPM = 9
ppDateTimeMMMMdyyyy = 4
ppDateTimeMMMMyy = 6
ppDateTimeMMyy = 7
ppDateTimeUAQ1 = 15
ppDateTimeUAQ2 = 16
ppDateTimeUAQ3 = 17
ppDateTimeUAQ4 = 18
ppDateTimeUAQ5 = 19
ppDateTimeUAQ6 = 20
ppDateTimeUAQ7 = 21

# PpDirection enumeration
ppDirectionLeftToRight = 1
ppDirectionMixed = -2
ppDirectionRightToLeft = 2

# PpEntryEffect enumeration
ppEffectAppear = 3844
ppEffectBlindsHorizontal = 769
ppEffectBlindsVertical = 770
ppEffectBoxDown = 3925
ppEffectBoxIn = 3074
ppEffectBoxLeft = 3922
ppEffectBoxOut = 3073
ppEffectBoxRight = 3924
ppEffectBoxUp = 3923
ppEffectCheckerboardAcross = 1025
ppEffectCheckerboardDown = 1026
ppEffectCircleOut = 3845
ppEffectCombHorizontal = 3847
ppEffectCombVertical = 3848
ppEffectConveyorLeft = 3882
ppEffectConveyorRight = 3883
ppEffectCoverDown = 1284
ppEffectCoverLeft = 1281
ppEffectCoverLeftDown = 1287
ppEffectCoverLeftUp = 1285
ppEffectCoverRight = 1283
ppEffectCoverRightDown = 1288
ppEffectCoverRightUp = 1286
ppEffectCoverUp = 1282
ppEffectCrawlFromDown = 3344
ppEffectCrawlFromLeft = 3341
ppEffectCrawlFromRight = 3343
ppEffectCrawlFromUp = 3342
ppEffectCubeDown = 3917
ppEffectCubeLeft = 3914
ppEffectCubeRight = 3916
ppEffectCubeUp = 3915
ppEffectCut = 257
ppEffectCutThroughBlack = 258
ppEffectDiamondOut = 3846
ppEffectDissolve = 1537
ppEffectDoorsHorizontal = 3885
ppEffectDoorsVertical = 3884
ppEffectFade = 1793
ppEffectFadeSmoothly = 3849
ppEffectFerrisWheelLeft = 3899
ppEffectFerrisWheelRight = 3900
ppEffectFlashbulb = 3909
ppEffectFlashOnceFast = 3841
ppEffectFlashOnceMedium = 3842
ppEffectFlashOnceSlow = 3843
ppEffectFlipDown = 3908
ppEffectFlipLeft = 3905
ppEffectFlipRight = 3907
ppEffectFlipUp = 3906
ppEffectFlyFromBottom = 3332
ppEffectFlyFromBottomLeft = 3335
ppEffectFlyFromBottomRight = 3336
ppEffectFlyFromLeft = 3329
ppEffectFlyFromRight = 3331
ppEffectFlyFromTop = 3330
ppEffectFlyFromTopLeft = 3333
ppEffectFlyFromTopRight = 3334
ppEffectFlyThroughIn = 3890
ppEffectFlyThroughInBounce = 3892
ppEffectFlyThroughOut = 3891
ppEffectFlyThroughOutBounce = 3893
ppEffectGalleryLeft = 3880
ppEffectGalleryRight = 3881
ppEffectGlitterDiamondDown = 3875
ppEffectGlitterDiamondLeft = 3872
ppEffectGlitterDiamondRight = 3874
ppEffectGlitterDiamondUp = 3873
ppEffectGlitterHexagonDown = 3879
ppEffectGlitterHexagonLeft = 3876
ppEffectGlitterHexagonRight = 3878
ppEffectGlitterHexagonUp = 3877
ppEffectHoneycomb = 3898
ppEffectMixed = -2
ppEffectNewsflash = 3850
ppEffectNone = 0
ppEffectOrbitDown = 3929
ppEffectOrbitLeft = 3926
ppEffectOrbitRight = 3928
ppEffectOrbitUp = 3927
ppEffectPanDown = 3933
ppEffectPanLeft = 3930
ppEffectPanRight = 3932
ppEffectPanUp = 3931
ppEffectPeekFromDown = 3338
ppEffectPeekFromLeft = 3337
ppEffectPeekFromRight = 3339
ppEffectPeekFromUp = 3340
ppEffectPlusOut = 3851
ppEffectPushDown = 3852
ppEffectPushLeft = 3853
ppEffectPushRight = 3854
ppEffectPushUp = 3855
ppEffectRandom = 513
ppEffectRandomBarsHorizontal = 2305
ppEffectRandomBarsVertical = 2306
ppEffectRevealBlackLeft = 3896
ppEffectRevealBlackRight = 3897
ppEffectRevealSmoothLeft = 3894
ppEffectRevealSmoothRight = 3895
ppEffectRippleCenter = 3867
ppEffectRippleLeftDown = 3870
ppEffectRippleLeftUp = 3869
ppEffectRippleRightDown = 3871
ppEffectRippleRightUp = 3868
ppEffectRotateDown = 3921
ppEffectRotateLeft = 3918
ppEffectRotateRight = 3920
ppEffectRotateUp = 3919
ppEffectShredRectangleIn = 3912
ppEffectShredRectangleOut = 3913
ppEffectShredStripsIn = 3910
ppEffectShredStripsOut = 3911
ppEffectSpiral = 3357
ppEffectSplitHorizontalIn = 3586
ppEffectSplitHorizontalOut = 3585
ppEffectSplitVerticalIn = 3588
ppEffectSplitVerticalOut = 3587
ppEffectStretchAcross = 3351
ppEffectStretchDown = 3355
ppEffectStretchLeft = 3352
ppEffectStretchRight = 3354
ppEffectStretchUp = 3353
ppEffectStripsDownLeft = 2563
ppEffectStripsDownRight = 2564
ppEffectStripsLeftDown = 2567
ppEffectStripsLeftUp = 2565
ppEffectStripsRightDown = 2568
ppEffectStripsRightUp = 2566
ppEffectStripsUpLeft = 2561
ppEffectStripsUpRight = 2562
ppEffectSwitchDown = 3904
ppEffectSwitchLeft = 3901
ppEffectSwitchRight = 3903
ppEffectSwitchUp = 3902
ppEffectSwivel = 3356
ppEffectUncoverDown = 2052
ppEffectUncoverLeft = 2049
ppEffectUncoverLeftDown = 2055
ppEffectUncoverLeftUp = 2053
ppEffectUncoverRight = 2051
ppEffectUncoverRightDown = 2056
ppEffectUncoverRightUp = 2054
ppEffectUncoverUp = 2050
ppEffectVortexDown = 3866
ppEffectVortexLeft = 3863
ppEffectVortexRight = 3865
ppEffectVortexUp = 3864
ppEffectWarpIn = 3888
ppEffectWarpOut = 3889
ppEffectWedge = 3856
ppEffectWheel1Spoke = 3857
ppEffectWheel2Spokes = 3858
ppEffectWheel3Spokes = 3859
ppEffectWheel4Spokes = 3860
ppEffectWheel8Spokes = 3861
ppEffectWheelReverse1Spoke = 3862
ppEffectWindowHorizontal = 3887
ppEffectWindowVertical = 3886
ppEffectWipeDown = 2820
ppEffectWipeLeft = 2817
ppEffectWipeRight = 2819
ppEffectWipeUp = 2818
ppEffectZoomBottom = 3350
ppEffectZoomCenter = 3349
ppEffectZoomIn = 3345
ppEffectZoomInSlightly = 3346
ppEffectZoomOut = 3347
ppEffectZoomOutSlightly = 3348

# PpFarEastLineBreakLevel enumeration
ppFarEastLineBreakLevelCustom = 3
ppFarEastLineBreakLevelNormal = 1
ppFarEastLineBreakLevelStrict = 2

# PpFixedFormatIntent enumeration
ppFixedFormatIntentPrint = 2
ppFixedFormatIntentScreen = 1

# PpFixedFormatType enumeration
ppFixedFormatTypePDF = 2
ppFixedFormatTypeXPS = 1

# PpFollowColors enumeration
ppFollowColorsMixed = -2
ppFollowColorsNone = 0
ppFollowColorsScheme = 1
ppFollowColorsTextAndBackground = 2

# PpFrameColors enumeration
ppFrameColorsBlackTextOnWhite = 5
ppFrameColorsBrowserColors = 1
ppFrameColorsPresentationSchemeAccentColor = 3
ppFrameColorsPresentationSchemeTextColor = 2
ppFrameColorsWhiteTextOnBlack = 4

# PpHTMLVersion enumeration
ppHTMLAutodetect = 4
ppHTMLDual = 3
ppHTMLv3 = 1
ppHTMLv4 = 2

# PpIndentControl enumeration
ppIndentControlMixed = -2
ppIndentKeepAttr = 2
ppIndentReplaceAttr = 1

# PpMediaType enumeration
ppMediaTypeMixed = -2
ppMediaTypeMovie = 3
ppMediaTypeOther = 1
ppMediaTypeSound = 2

# PpMouseActivation enumeration
ppMouseClick = 1
ppMouseOver = 2

# PpNumberedBulletStyle enumeration
ppBulletAlphaLCParenBoth = 8
ppBulletAlphaLCParenRight = 9
ppBulletAlphaLCPeriod = 0
ppBulletAlphaUCParenBoth = 10
ppBulletAlphaUCParenRight = 11
ppBulletAlphaUCPeriod = 1
ppBulletArabicAbjadDash = 24
ppBulletArabicAlphaDash = 23
ppBulletArabicDBPeriod = 29
ppBulletArabicDBPlain = 28
ppBulletArabicParenBoth = 12
ppBulletArabicParenRight = 2
ppBulletArabicPeriod = 3
ppBulletArabicPlain = 13
ppBulletCircleNumDBPlain = 18
ppBulletCircleNumWDBlackPlain = 20
ppBulletCircleNumWDWhitePlain = 19
ppBulletHebrewAlphaDash = 25
ppBulletHindiAlpha1Period = 40
ppBulletHindiAlphaPeriod = 36
ppBulletHindiNumParenRight = 39
ppBulletHindiNumPeriod = 37
ppBulletKanjiKoreanPeriod = 27
ppBulletKanjiKoreanPlain = 26
ppBulletKanjiSimpChinDBPeriod = 38
ppBulletRomanLCParenBoth = 4
ppBulletRomanLCParenRight = 5
ppBulletRomanLCPeriod = 6
ppBulletRomanUCParenBoth = 14
ppBulletRomanUCParenRight = 15
ppBulletRomanUCPeriod = 7
ppBulletSimpChinPeriod = 17
ppBulletSimpChinPlain = 16
ppBulletStyleMixed = -2
ppBulletThaiAlphaParenBoth = 32
ppBulletThaiAlphaParenRight = 31
ppBulletThaiAlphaPeriod = 30
ppBulletThaiNumParenBoth = 35
ppBulletThaiNumParenRight = 34
ppBulletThaiNumPeriod = 33
ppBulletTradChinPeriod = 22
ppBulletTradChinPlain = 21

# PpParagraphAlignment enumeration
ppAlignCenter = 2
ppAlignDistribute = 5
ppAlignJustify = 4
ppAlignJustifyLow = 7
ppAlignLeft = 1
ppAlignmentMixed = -2
ppAlignRight = 3
ppAlignThaiDistribute = 6

# PpPasteDataType enumeration
ppPasteBitmap = 1
ppPasteDefault = 0
ppPasteEnhancedMetafile = 2
ppPasteGIF = 4
ppPasteHTML = 8
ppPasteJPG = 5
ppPasteMetafilePicture = 3
ppPasteOLEObject = 10
ppPastePNG = 6
ppPasteRTF = 9
ppPasteShape = 11
ppPasteText = 7

# PpPlaceholderType enumeration
ppPlaceholderMixed = -2
ppPlaceholderTitle = 1
ppPlaceholderBody = 2
ppPlaceholderCenterTitle = 3
ppPlaceholderSubtitle = 4
ppPlaceholderVerticalTitle = 5
ppPlaceholderVerticalBody = 6
ppPlaceholderObject = 7
ppPlaceholderChart = 8
ppPlaceholderBitmap = 9
ppPlaceholderMediaClip = 10
ppPlaceholderOrgChart = 11
ppPlaceholderTable = 12
ppPlaceholderSlideNumber = 13
ppPlaceholderHeader = 14
ppPlaceholderFooter = 15
ppPlaceholderDate = 16
ppPlaceholderVerticalObject = 17
ppPlaceholderPicture = 18
ppPlaceholderCameo = 19

# PpPrintColorType enumeration
ppPrintBlackAndWhite = 2
ppPrintColor = 1
ppPrintPureBlackAndWhite = 3

# PpPrintHandoutOrder enumeration
ppPrintHandoutHorizontalFirst = 2
ppPrintHandoutVerticalFirst = 1

# PpPrintOutputType enumeration
ppPrintOutputBuildSlides = 7
ppPrintOutputFourSlideHandouts = 8
ppPrintOutputNineSlideHandouts = 9
ppPrintOutputNotesPages = 5
ppPrintOutputOneSlideHandouts = 10
ppPrintOutputOutline = 6
ppPrintOutputSixSlideHandouts = 4
ppPrintOutputSlides = 1
ppPrintOutputThreeSlideHandouts = 3
ppPrintOutputTwoSlideHandouts = 2

# PpPrintRangeType enumeration
ppPrintAll = 1
ppPrintCurrent = 3
ppPrintNamedSlideShow = 5
ppPrintSelection = 2
ppPrintSlideRange = 4

# PpPublishSourceType enumeration
ppPublishAll = 1
ppPublishNamedSlideShow = 3
ppPublishSlideRange = 2

# PpRemoveDocInfoType enumeration
ppRDIAll = 99
ppRDIAtMentions = 18
ppRDIComments = 1
ppRDIContentType = 16
ppRDIDocumentManagementPolicy = 15
ppRDIDocumentProperties = 8
ppRDIDocumentServerProperties = 14
ppRDIDocumentWorkspace = 10
ppRDIInkAnnotations = 11
ppRDIPublishPath = 13
ppRDIRemovePersonalInformation = 4
ppRDISlideUpdateInformation = 17

# PpRevisionInfo enumeration
ppRevisionInfoBaseline = 1
ppRevisionInfoMerged = 2
ppRevisionInfoNone = 0

# PpSelectionType enumeration
ppSelectionNone = 0
ppSelectionShapes = 2
ppSelectionSlides = 1
ppSelectionText = 3

# PpSlideLayout enumeration
ppLayoutBlank = 12
ppLayoutChart = 8
ppLayoutChartAndText = 6
ppLayoutClipArtAndText = 10
ppLayoutClipArtAndVerticalText = 26
ppLayoutComparison = 34
ppLayoutContentWithCaption = 35
ppLayoutCustom = 32
ppLayoutFourObjects = 24
ppLayoutLargeObject = 15
ppLayoutMediaClipAndText = 18
ppLayoutMixed = -2
ppLayoutObject = 16
ppLayoutObjectAndText = 14
ppLayoutObjectAndTwoObjects = 30
ppLayoutObjectOverText = 19
ppLayoutOrgchart = 7
ppLayoutPictureWithCaption = 36
ppLayoutSectionHeader = 33
ppLayoutTable = 4
ppLayoutText = 2
ppLayoutTextAndChart = 5
ppLayoutTextAndClipArt = 9
ppLayoutTextAndMediaClip = 17
ppLayoutTextAndObject = 13
ppLayoutTextAndTwoObjects = 21
ppLayoutTextOverObject = 20
ppLayoutTitle = 1
ppLayoutTitleOnly = 11
ppLayoutTwoColumnText = 3
ppLayoutTwoObjects = 29
ppLayoutTwoObjectsAndObject = 31
ppLayoutTwoObjectsAndText = 22
ppLayoutTwoObjectsOverText = 23
ppLayoutVerticalText = 25
ppLayoutVerticalTitleAndText = 27
ppLayoutVerticalTitleAndTextOverChart = 28

# PpSlideShowAdvanceMode enumeration
ppSlideShowManualAdvance = 1
ppSlideShowRehearseNewTimings = 3
ppSlideShowUseSlideTimings = 2

# PpSlideShowPointerType enumeration
ppSlideShowPointerAlwaysHidden = 3
ppSlideShowPointerArrow = 1
ppSlideShowPointerAutoArrow = 4
ppSlideShowPointerEraser = 5
ppSlideShowPointerNone = 0
ppSlideShowPointerPen = 2

# PpSlideShowRangeType enumeration
ppShowAll = 1
ppShowNamedSlideShow = 3
ppShowSlideRange = 2

# PpSlideShowState enumeration
ppSlideShowBlackScreen = 3
ppSlideShowDone = 5
ppSlideShowPaused = 2
ppSlideShowRunning = 1
ppSlideShowWhiteScreen = 4

# PpSlideShowType enumeration
ppShowTypeKiosk = 3
ppShowTypeSpeaker = 1
ppShowTypeWindow = 2
ppShowTypeWindow2 = 4

# PpSlideSizeType enumeration
ppSlideSize35MM = 4
ppSlideSizeA3Paper = 9
ppSlideSizeA4Paper = 3
ppSlideSizeB4ISOPaper = 10
ppSlideSizeB4JISPaper = 12
ppSlideSizeB5ISOPaper = 11
ppSlideSizeB5JISPaper = 13
ppSlideSizeBanner = 6
ppSlideSizeCustom = 7
ppSlideSizeHagakiCard = 14
ppSlideSizeLedgerPaper = 8
ppSlideSizeLetterPaper = 2
ppSlideSizeOnScreen = 1
ppSlideSizeOverhead = 5

# PpSoundEffectType enumeration
ppSoundEffectsMixed = -2
ppSoundFile = 2
ppSoundNone = 0
ppSoundStopPrevious = 1

# PpSoundFormatType enumeration
ppSoundFormatCDAudio = 3
ppSoundFormatMIDI = 2
ppSoundFormatMixed = -2
ppSoundFormatNone = 0
ppSoundFormatWAV = 1

# PpTabStopType enumeration
ppTabStopCenter = 2
ppTabStopDecimal = 4
ppTabStopLeft = 1
ppTabStopMixed = -2
ppTabStopRight = 3

# PpTextLevelEffect enumeration
ppAnimateByAllLevels = 16
ppAnimateByFifthLevel = 5
ppAnimateByFirstLevel = 1
ppAnimateByFourthLevel = 4
ppAnimateBySecondLevel = 2
ppAnimateByThirdLevel = 3
ppAnimateLevelMixed = -2
ppAnimateLevelNone = 0

# PpTextStyleType enumeration
ppBodyStyle = 3
ppDefaultStyle = 1
ppTitleStyle = 2

# PpTextUnitEffect enumeration
ppAnimateByCharacter = 2
ppAnimateByParagraph = 0
ppAnimateByWord = 1
ppAnimateUnitMixed = -2

# PpTransitionSpeed enumeration
ppTransitionSpeedFast = 3
ppTransitionSpeedMedium = 2
ppTransitionSpeedMixed = -2
ppTransitionSpeedSlow = 1

# PpUpdateOption enumeration
ppUpdateOptionAutomatic = 2
ppUpdateOptionManual = 1
ppUpdateOptionMixed = -2

# PpViewType enumeration
ppViewHandoutMaster = 4
ppViewMasterThumbnails = 12
ppViewNormal = 9
ppViewNotesMaster = 5
ppViewNotesPage = 3
ppViewOutline = 6
ppViewPrintPreview = 10
ppViewSlide = 1
ppViewSlideMaster = 2
ppViewSlideSorter = 7
ppViewThumbnails = 11
ppViewTitleMaster = 8

# PpWindowState enumeration
ppWindowMaximized = 3
ppWindowMinimized = 2
ppWindowNormal = 1

class Presentation:

    def __init__(self, presentation=None):
        self.presentation = presentation

    @property
    def Application(self):
        return Application(self.presentation.Application)

    @property
    def AutoSaveOn(self):
        return self.presentation.AutoSaveOn

    @AutoSaveOn.setter
    def AutoSaveOn(self, value):
        self.presentation.AutoSaveOn = value

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

    @DisplayComments.setter
    def DisplayComments(self, value):
        self.presentation.DisplayComments = value

    @property
    def DocumentInspectors(self):
        return self.presentation.DocumentInspectors

    @property
    def DocumentLibraryVersions(self):
        return self.presentation.DocumentLibraryVersions

    @property
    def EncryptionProvider(self):
        return self.presentation.EncryptionProvider

    @EncryptionProvider.setter
    def EncryptionProvider(self, value):
        self.presentation.EncryptionProvider = value

    @property
    def EnvelopeVisible(self):
        return self.presentation.EnvelopeVisible

    @EnvelopeVisible.setter
    def EnvelopeVisible(self, value):
        self.presentation.EnvelopeVisible = value

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

    @Final.setter
    def Final(self, value):
        self.presentation.Final = value

    @property
    def Fonts(self):
        return Fonts(self.presentation.Fonts)

    @property
    def FullName(self):
        return self.presentation.FullName

    @property
    def GridDistance(self):
        return self.presentation.GridDistance

    @GridDistance.setter
    def GridDistance(self, value):
        self.presentation.GridDistance = value

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

    @RemovePersonalInformation.setter
    def RemovePersonalInformation(self, value):
        self.presentation.RemovePersonalInformation = value

    @property
    def Research(self):
        return Research(self.presentation.Research)

    @property
    def Saved(self):
        return self.presentation.Saved

    @Saved.setter
    def Saved(self, value):
        self.presentation.Saved = value

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

    @SnapToGrid.setter
    def SnapToGrid(self, value):
        self.presentation.SnapToGrid = value

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

    @WritePassword.setter
    def WritePassword(self, value):
        self.presentation.WritePassword = value

    def AcceptAll(self):
        return self.presentation.AcceptAll()

    def AddTitleMaster(self):
        return self.presentation.AddTitleMaster()

    def AddToFavorites(self):
        self.presentation.AddToFavorites()

    def ApplyTemplate(self, FileName=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
        ]
        self.presentation.ApplyTemplate(*params)

    def ApplyTheme(self, themeName=None):
        params = [
            themeName if themeName is not None else pythoncom.Missing,
        ]
        self.presentation.ApplyTheme(*params)

    def CanCheckIn(self):
        return self.presentation.CanCheckIn()

    def CheckIn(self, SaveChanges=None, Comments=None, MakePublic=None):
        params = [
            SaveChanges if SaveChanges is not None else pythoncom.Missing,
            Comments if Comments is not None else pythoncom.Missing,
            MakePublic if MakePublic is not None else pythoncom.Missing,
        ]
        self.presentation.CheckIn(*params)

    def CheckInWithVersion(self, SaveChanges=None, Comments=None, MakePublic=None, VersionType=None):
        params = [
            SaveChanges if SaveChanges is not None else pythoncom.Missing,
            Comments if Comments is not None else pythoncom.Missing,
            MakePublic if MakePublic is not None else pythoncom.Missing,
            VersionType if VersionType is not None else pythoncom.Missing,
        ]
        self.presentation.CheckInWithVersion(*params)

    def Close(self):
        self.presentation.Close()

    def Convert2(self, FileName=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
        ]
        self.presentation.Convert2(*params)

    def CreateVideo(self, FileName=None, UseTimingsAndNarrations=None, DefaultSlideDuration=None, VertResolution=None, FramesPerSecond=None, Quality=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            UseTimingsAndNarrations if UseTimingsAndNarrations is not None else pythoncom.Missing,
            DefaultSlideDuration if DefaultSlideDuration is not None else pythoncom.Missing,
            VertResolution if VertResolution is not None else pythoncom.Missing,
            FramesPerSecond if FramesPerSecond is not None else pythoncom.Missing,
            Quality if Quality is not None else pythoncom.Missing,
        ]
        self.presentation.CreateVideo(*params)

    def EndReview(self):
        return self.presentation.EndReview()

    def EnsureAllMediaUpgraded(self):
        self.presentation.EnsureAllMediaUpgraded()

    def Export(self, Path=None, FilterName=None, ScaleWidth=None, ScaleHeight=None):
        params = [
            Path if Path is not None else pythoncom.Missing,
            FilterName if FilterName is not None else pythoncom.Missing,
            ScaleWidth if ScaleWidth is not None else pythoncom.Missing,
            ScaleHeight if ScaleHeight is not None else pythoncom.Missing,
        ]
        self.presentation.Export(*params)

    def ExportAsFixedFormat(self, Path=None, FixedFormatType=None, Intent=None, FrameSlides=None, HandoutOrder=None, OutputType=None, PrintHiddenSlides=None, PrintRange=None, RangeType=None, SlideShowName=None, IncludeDocProperties=None, KeepIRMSettings=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005_1=None, ExternalExporter=None):
        params = [
            Path if Path is not None else pythoncom.Missing,
            FixedFormatType if FixedFormatType is not None else pythoncom.Missing,
            Intent if Intent is not None else pythoncom.Missing,
            FrameSlides if FrameSlides is not None else pythoncom.Missing,
            HandoutOrder if HandoutOrder is not None else pythoncom.Missing,
            OutputType if OutputType is not None else pythoncom.Missing,
            PrintHiddenSlides if PrintHiddenSlides is not None else pythoncom.Missing,
            PrintRange if PrintRange is not None else pythoncom.Missing,
            RangeType if RangeType is not None else pythoncom.Missing,
            SlideShowName if SlideShowName is not None else pythoncom.Missing,
            IncludeDocProperties if IncludeDocProperties is not None else pythoncom.Missing,
            KeepIRMSettings if KeepIRMSettings is not None else pythoncom.Missing,
            DocStructureTags if DocStructureTags is not None else pythoncom.Missing,
            BitmapMissingFonts if BitmapMissingFonts is not None else pythoncom.Missing,
            UseISO19005_1 if UseISO19005_1 is not None else pythoncom.Missing,
            ExternalExporter if ExternalExporter is not None else pythoncom.Missing,
        ]
        self.presentation.ExportAsFixedFormat(*params)

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
        return self.presentation.FollowHyperlink(*params)

    def GetWorkflowTasks(self):
        return self.presentation.GetWorkflowTasks()

    def GetWorkflowTemplates(self):
        return self.presentation.GetWorkflowTemplates()

    def LockServerFile(self):
        self.presentation.LockServerFile()

    def MergeWithBaseline(self, withPresentation=None, baselinePresentation=None):
        params = [
            withPresentation if withPresentation is not None else pythoncom.Missing,
            baselinePresentation if baselinePresentation is not None else pythoncom.Missing,
        ]
        return self.presentation.MergeWithBaseline(*params)

    def NewWindow(self):
        return self.presentation.NewWindow()

    def PrintOut(self, From=None, To=None, PrintToFile=None, Copies=None, Collate=None):
        params = [
            From if From is not None else pythoncom.Missing,
            To if To is not None else pythoncom.Missing,
            PrintToFile if PrintToFile is not None else pythoncom.Missing,
            Copies if Copies is not None else pythoncom.Missing,
            Collate if Collate is not None else pythoncom.Missing,
        ]
        self.presentation.PrintOut(*params)

    def PublishSlides(self, SlideLibraryUrl=None, Overwrite=None):
        params = [
            SlideLibraryUrl if SlideLibraryUrl is not None else pythoncom.Missing,
            Overwrite if Overwrite is not None else pythoncom.Missing,
        ]
        self.presentation.PublishSlides(*params)

    def RejectAll(self):
        return self.presentation.RejectAll()

    def RemoveDocumentInformation(self, Type=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
        ]
        self.presentation.RemoveDocumentInformation(*params)

    def Save(self):
        self.presentation.Save()

    def SaveAs(self, FileName=None, FileFormat=None, EmbedFonts=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            FileFormat if FileFormat is not None else pythoncom.Missing,
            EmbedFonts if EmbedFonts is not None else pythoncom.Missing,
        ]
        self.presentation.SaveAs(*params)

    def SaveCopyAs(self, FileName=None, FileFormat=None, EmbedTrueTypeFonts=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            FileFormat if FileFormat is not None else pythoncom.Missing,
            EmbedTrueTypeFonts if EmbedTrueTypeFonts is not None else pythoncom.Missing,
        ]
        self.presentation.SaveCopyAs(*params)

    def SaveCopyAs2(self, FileName=None, FileFormat=None, EmbedTrueTypeFonts=None, ReadOnlyRecommended=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            FileFormat if FileFormat is not None else pythoncom.Missing,
            EmbedTrueTypeFonts if EmbedTrueTypeFonts is not None else pythoncom.Missing,
            ReadOnlyRecommended if ReadOnlyRecommended is not None else pythoncom.Missing,
        ]
        self.presentation.SaveCopyAs2(*params)

    def SendFaxOverInternet(self, Recipients=None, Subject=None, ShowMessage=None):
        params = [
            Recipients if Recipients is not None else pythoncom.Missing,
            Subject if Subject is not None else pythoncom.Missing,
            ShowMessage if ShowMessage is not None else pythoncom.Missing,
        ]
        self.presentation.SendFaxOverInternet(*params)

    def SetPasswordEncryptionOptions(self, PasswordEncryptionProvider=None, PasswordEncryptionAlgorithm=None, PasswordEncryptionKeyLength=None, PasswordEncryptionFileProperties=None):
        params = [
            PasswordEncryptionProvider if PasswordEncryptionProvider is not None else pythoncom.Missing,
            PasswordEncryptionAlgorithm if PasswordEncryptionAlgorithm is not None else pythoncom.Missing,
            PasswordEncryptionKeyLength if PasswordEncryptionKeyLength is not None else pythoncom.Missing,
            PasswordEncryptionFileProperties if PasswordEncryptionFileProperties is not None else pythoncom.Missing,
        ]
        self.presentation.SetPasswordEncryptionOptions(*params)

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

    def Add(self, WithWindow=None):
        params = [
            WithWindow if WithWindow is not None else pythoncom.Missing,
        ]
        return Presentation(self.presentations.Add(*params))

    def CanCheckOut(self, FileName=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
        ]
        return self.presentations.CanCheckOut(*params)

    def CheckOut(self, FileName=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
        ]
        return self.presentations.CheckOut(*params)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.presentations.Item(*params)

    def Open(self, FileName=None, ReadOnly=None, Untitled=None, WithWindow=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            ReadOnly if ReadOnly is not None else pythoncom.Missing,
            Untitled if Untitled is not None else pythoncom.Missing,
            WithWindow if WithWindow is not None else pythoncom.Missing,
        ]
        return Presentation(self.presentations.Open(*params))

    def Open2007(self, FileName=None, ReadOnly=None, Untitled=None, WithWindow=None, OpenAndRepair=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            ReadOnly if ReadOnly is not None else pythoncom.Missing,
            Untitled if Untitled is not None else pythoncom.Missing,
            WithWindow if WithWindow is not None else pythoncom.Missing,
            OpenAndRepair if OpenAndRepair is not None else pythoncom.Missing,
        ]
        return Presentation(self.presentations.Open2007(*params))


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

    @Collate.setter
    def Collate(self, value):
        self.printoptions.Collate = value

    @property
    def FitToPage(self):
        return self.printoptions.FitToPage

    @FitToPage.setter
    def FitToPage(self, value):
        self.printoptions.FitToPage = value

    @property
    def FrameSlides(self):
        return self.printoptions.FrameSlides

    @FrameSlides.setter
    def FrameSlides(self, value):
        self.printoptions.FrameSlides = value

    @property
    def HandoutOrder(self):
        return self.printoptions.HandoutOrder

    @HandoutOrder.setter
    def HandoutOrder(self, value):
        self.printoptions.HandoutOrder = value

    @property
    def HighQuality(self):
        return self.printoptions.HighQuality

    @HighQuality.setter
    def HighQuality(self, value):
        self.printoptions.HighQuality = value

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

    @PrintComments.setter
    def PrintComments(self, value):
        self.printoptions.PrintComments = value

    @property
    def PrintFontsAsGraphics(self):
        return self.printoptions.PrintFontsAsGraphics

    @PrintFontsAsGraphics.setter
    def PrintFontsAsGraphics(self, value):
        self.printoptions.PrintFontsAsGraphics = value

    @property
    def PrintHiddenSlides(self):
        return self.printoptions.PrintHiddenSlides

    @PrintHiddenSlides.setter
    def PrintHiddenSlides(self, value):
        self.printoptions.PrintHiddenSlides = value

    @property
    def PrintInBackground(self):
        return self.printoptions.PrintInBackground

    @PrintInBackground.setter
    def PrintInBackground(self, value):
        self.printoptions.PrintInBackground = value

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

    @sectionIndex.setter
    def sectionIndex(self, value):
        self.printoptions.sectionIndex = value

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

    def Add(self, Start=None, End=None):
        params = [
            Start if Start is not None else pythoncom.Missing,
            End if End is not None else pythoncom.Missing,
        ]
        return PrintRange(self.printranges.Add(*params))

    def ClearAll(self):
        return self.printranges.ClearAll()

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.printranges.Item(*params)


class PropertyEffect:

    def __init__(self, propertyeffect=None):
        self.propertyeffect = propertyeffect

    @property
    def Application(self):
        return Application(self.propertyeffect.Application)

    @property
    def From(self):
        return self.propertyeffect.From

    @From.setter
    def From(self, value):
        self.propertyeffect.From = value

    @property
    def Parent(self):
        return self.propertyeffect.Parent

    @property
    def Points(self):
        return AnimationPoints(self.propertyeffect.Points)

    @property
    def Property(self):
        return self.propertyeffect.Property

    @Property.setter
    def Property(self, value):
        self.propertyeffect.Property = value

    @property
    def To(self):
        return self.propertyeffect.To

    @To.setter
    def To(self, value):
        self.propertyeffect.To = value


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

    def Edit(self, ModifyPassword=None):
        params = [
            ModifyPassword if ModifyPassword is not None else pythoncom.Missing,
        ]
        return self.protectedviewwindow.Edit(*params)


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

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.protectedviewwindows.Item(*params)

    def Open(self, FileName=None, ReadPassword=None, OpenAndRepair=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            ReadPassword if ReadPassword is not None else pythoncom.Missing,
            OpenAndRepair if OpenAndRepair is not None else pythoncom.Missing,
        ]
        return self.protectedviewwindows.Open(*params)


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

    @SpeakerNotes.setter
    def SpeakerNotes(self, value):
        self.publishobject.SpeakerNotes = value

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

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.publishobjects.Item(*params)


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

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.resamplemediatasks.Item(*params)

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
        self.research.Query(*params)

    def SetLanguagePair(self, Language1=None, Language2=None):
        params = [
            Language1 if Language1 is not None else pythoncom.Missing,
            Language2 if Language2 is not None else pythoncom.Missing,
        ]
        self.research.SetLanguagePair(*params)


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

    @By.setter
    def By(self, value):
        self.rotationeffect.By = value

    @property
    def From(self):
        return self.rotationeffect.From

    @From.setter
    def From(self, value):
        self.rotationeffect.From = value

    @property
    def Parent(self):
        return self.rotationeffect.Parent

    @property
    def To(self):
        return self.rotationeffect.To

    @To.setter
    def To(self, value):
        self.rotationeffect.To = value


class Row:

    def __init__(self, row=None):
        self.row = row

    @property
    def Application(self):
        return Application(self.row.Application)

    def Cells(self, RowIndex=None, ColumnIndex=None):
        params = [
            RowIndex if RowIndex is not None else pythoncom.Missing,
            ColumnIndex if ColumnIndex is not None else pythoncom.Missing,
        ]
        if callable(self.row.Cells):
            return CellRange(self.row.Cells(*params))
        else:
            return CellRange(self.row.GetCells(*params))

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

    def Add(self, BeforeRow=None):
        params = [
            BeforeRow if BeforeRow is not None else pythoncom.Missing,
        ]
        return Row(self.rows.Add(*params))

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.rows.Item(*params)


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

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.rulerlevels.Item(*params)


class ScaleEffect:

    def __init__(self, scaleeffect=None):
        self.scaleeffect = scaleeffect

    @property
    def Application(self):
        return Application(self.scaleeffect.Application)

    @property
    def ByX(self):
        return self.scaleeffect.ByX

    @ByX.setter
    def ByX(self, value):
        self.scaleeffect.ByX = value

    @property
    def ByY(self):
        return self.scaleeffect.ByY

    @ByY.setter
    def ByY(self, value):
        self.scaleeffect.ByY = value

    @property
    def FromX(self):
        return self.scaleeffect.FromX

    @FromX.setter
    def FromX(self, value):
        self.scaleeffect.FromX = value

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

    @ToX.setter
    def ToX(self, value):
        self.scaleeffect.ToX = value

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

    def AddBeforeSlide(self, SlideIndex=None, sectionName=None):
        params = [
            SlideIndex if SlideIndex is not None else pythoncom.Missing,
            sectionName if sectionName is not None else pythoncom.Missing,
        ]
        return self.sectionproperties.AddBeforeSlide(*params)

    def AddSection(self, sectionIndex=None, sectionName=None):
        params = [
            sectionIndex if sectionIndex is not None else pythoncom.Missing,
            sectionName if sectionName is not None else pythoncom.Missing,
        ]
        return self.sectionproperties.AddSection(*params)

    def Delete(self, sectionIndex=None, deleteSlides=None):
        params = [
            sectionIndex if sectionIndex is not None else pythoncom.Missing,
            deleteSlides if deleteSlides is not None else pythoncom.Missing,
        ]
        self.sectionproperties.Delete(*params)

    def FirstSlide(self, sectionIndex=None):
        params = [
            sectionIndex if sectionIndex is not None else pythoncom.Missing,
        ]
        return self.sectionproperties.FirstSlide(*params)

    def Move(self, sectionIndex=None, toPos=None):
        params = [
            sectionIndex if sectionIndex is not None else pythoncom.Missing,
            toPos if toPos is not None else pythoncom.Missing,
        ]
        self.sectionproperties.Move(*params)

    def Name(self, sectionIndex=None):
        params = [
            sectionIndex if sectionIndex is not None else pythoncom.Missing,
        ]
        return self.sectionproperties.Name(*params)

    def Rename(self, sectionIndex=None, sectionName=None):
        params = [
            sectionIndex if sectionIndex is not None else pythoncom.Missing,
            sectionName if sectionName is not None else pythoncom.Missing,
        ]
        self.sectionproperties.Rename(*params)

    def SectionID(self, sectionIndex=None):
        params = [
            sectionIndex if sectionIndex is not None else pythoncom.Missing,
        ]
        return self.sectionproperties.SectionID(*params)

    def SlidesCount(self, sectionIndex=None):
        params = [
            sectionIndex if sectionIndex is not None else pythoncom.Missing,
        ]
        return self.sectionproperties.SlidesCount(*params)


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

    def AddEffect(self, Shape=None, effectId=None, Level=None, trigger=None, Index=None):
        params = [
            Shape if Shape is not None else pythoncom.Missing,
            effectId if effectId is not None else pythoncom.Missing,
            Level if Level is not None else pythoncom.Missing,
            trigger if trigger is not None else pythoncom.Missing,
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.sequence.AddEffect(*params)

    def AddTriggerEffect(self, pShape=None, effectId=None, trigger=None, pTriggerShape=None, bookmark=None, Level=None):
        params = [
            pShape if pShape is not None else pythoncom.Missing,
            effectId if effectId is not None else pythoncom.Missing,
            trigger if trigger is not None else pythoncom.Missing,
            pTriggerShape if pTriggerShape is not None else pythoncom.Missing,
            bookmark if bookmark is not None else pythoncom.Missing,
            Level if Level is not None else pythoncom.Missing,
        ]
        return self.sequence.AddTriggerEffect(*params)

    def Clone(self, Effect=None, Index=None):
        params = [
            Effect if Effect is not None else pythoncom.Missing,
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.sequence.Clone(*params)

    def ConvertToAfterEffect(self, Effect=None, After=None, DimColor=None, DimSchemeColor=None):
        params = [
            Effect if Effect is not None else pythoncom.Missing,
            After if After is not None else pythoncom.Missing,
            DimColor if DimColor is not None else pythoncom.Missing,
            DimSchemeColor if DimSchemeColor is not None else pythoncom.Missing,
        ]
        return self.sequence.ConvertToAfterEffect(*params)

    def ConvertToAnimateBackground(self, Effect=None, AnimateBackground=None):
        params = [
            Effect if Effect is not None else pythoncom.Missing,
            AnimateBackground if AnimateBackground is not None else pythoncom.Missing,
        ]
        return self.sequence.ConvertToAnimateBackground(*params)

    def ConvertToAnimateInReverse(self, Effect=None, animateInReverse=None):
        params = [
            Effect if Effect is not None else pythoncom.Missing,
            animateInReverse if animateInReverse is not None else pythoncom.Missing,
        ]
        return self.sequence.ConvertToAnimateInReverse(*params)

    def ConvertToBuildLevel(self, Effect=None, Level=None):
        params = [
            Effect if Effect is not None else pythoncom.Missing,
            Level if Level is not None else pythoncom.Missing,
        ]
        return self.sequence.ConvertToBuildLevel(*params)

    def ConvertToTextUnitEffect(self, Effect=None, unitEffect=None):
        params = [
            Effect if Effect is not None else pythoncom.Missing,
            unitEffect if unitEffect is not None else pythoncom.Missing,
        ]
        return self.sequence.ConvertToTextUnitEffect(*params)

    def FindFirstAnimationFor(self, Shape=None):
        params = [
            Shape if Shape is not None else pythoncom.Missing,
        ]
        return self.sequence.FindFirstAnimationFor(*params)

    def FindFirstAnimationForClick(self, click=None):
        params = [
            click if click is not None else pythoncom.Missing,
        ]
        return self.sequence.FindFirstAnimationForClick(*params)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.sequence.Item(*params)


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

    def Add(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.sequences.Add(*params)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.sequences.Item(*params)


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
        return XlAxisGroup(self.series.AxisGroup)

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
        self.series.ClearFormats()

    def Copy(self):
        self.series.Copy()

    def DataLabels(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.series.DataLabels(*params)

    def Delete(self):
        self.series.Delete()

    def ErrorBar(self, Direction=None, Include=None, Type=None, Amount=None, MinusValues=None):
        params = [
            Direction if Direction is not None else pythoncom.Missing,
            Include if Include is not None else pythoncom.Missing,
            Type if Type is not None else pythoncom.Missing,
            Amount if Amount is not None else pythoncom.Missing,
            MinusValues if MinusValues is not None else pythoncom.Missing,
        ]
        self.series.ErrorBar(*params)

    def Paste(self):
        self.series.Paste()

    def Points(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return Points(self.series.Points(*params))

    def Select(self):
        self.series.Select()

    def Trendlines(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return Trendlines(self.series.Trendlines(*params))


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
        params = [
            Source if Source is not None else pythoncom.Missing,
            Rowcol if Rowcol is not None else pythoncom.Missing,
            SeriesLabels if SeriesLabels is not None else pythoncom.Missing,
            CategoryLabels if CategoryLabels is not None else pythoncom.Missing,
            Replace if Replace is not None else pythoncom.Missing,
        ]
        return Series(self.seriescollection.Add(*params))

    def Extend(self, Source=None, Rowcol=None, CategoryLabels=None):
        params = [
            Source if Source is not None else pythoncom.Missing,
            Rowcol if Rowcol is not None else pythoncom.Missing,
            CategoryLabels if CategoryLabels is not None else pythoncom.Missing,
        ]
        self.seriescollection.Extend(*params)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return Series(self.seriescollection.Item(*params))

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

    @Property.setter
    def Property(self, value):
        self.seteffect.Property = value

    @property
    def To(self):
        return self.seteffect.To

    @To.setter
    def To(self, value):
        self.seteffect.To = value


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

    @Decorative.setter
    def Decorative(self, value):
        self.shape.Decorative = value

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

    @LockAspectRatio.setter
    def LockAspectRatio(self, value):
        self.shape.LockAspectRatio = value

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

    @ShapeStyle.setter
    def ShapeStyle(self, value):
        self.shape.ShapeStyle = value

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

    def ConvertTextToSmartArt(self, Layout=None):
        params = [
            Layout if Layout is not None else pythoncom.Missing,
        ]
        self.shape.ConvertTextToSmartArt(*params)

    def Copy(self):
        self.shape.Copy()

    def Cut(self):
        self.shape.Cut()

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

    def PickupAnimation(self):
        self.shape.PickupAnimation()

    def RerouteConnections(self):
        self.shape.RerouteConnections()

    def ScaleHeight(self, Factor=None, RelativeToOriginalSize=None, fScale=None):
        params = [
            Factor if Factor is not None else pythoncom.Missing,
            RelativeToOriginalSize if RelativeToOriginalSize is not None else pythoncom.Missing,
            fScale if fScale is not None else pythoncom.Missing,
        ]
        self.shape.ScaleHeight(*params)

    def ScaleWidth(self, Factor=None, RelativeToOriginalSize=None, fScale=None):
        params = [
            Factor if Factor is not None else pythoncom.Missing,
            RelativeToOriginalSize if RelativeToOriginalSize is not None else pythoncom.Missing,
            fScale if fScale is not None else pythoncom.Missing,
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
        return self.shape.Ungroup()

    def UpgradeMedia(self):
        self.shape.UpgradeMedia()

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
        return self.shapenodes.Item(*params)

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

    @BackgroundStyle.setter
    def BackgroundStyle(self, value):
        self.shaperange.BackgroundStyle = value

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

    @Decorative.setter
    def Decorative(self, value):
        self.shaperange.Decorative = value

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

    @LockAspectRatio.setter
    def LockAspectRatio(self, value):
        self.shaperange.LockAspectRatio = value

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

    @Name.setter
    def Name(self, value):
        self.shaperange.Name = value

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

    def Align(self, AlignCmd=None, RelativeTo=None):
        params = [
            AlignCmd if AlignCmd is not None else pythoncom.Missing,
            RelativeTo if RelativeTo is not None else pythoncom.Missing,
        ]
        self.shaperange.Align(*params)

    def Apply(self):
        self.shaperange.Apply()

    def ApplyAnimation(self):
        self.shaperange.ApplyAnimation()

    def ConvertTextToSmartArt(self, Layout=None):
        params = [
            Layout if Layout is not None else pythoncom.Missing,
        ]
        return self.shaperange.ConvertTextToSmartArt(*params)

    def Copy(self):
        self.shaperange.Copy()

    def Cut(self):
        self.shaperange.Cut()

    def Delete(self):
        self.shaperange.Delete()

    def Distribute(self, DistributeCmd=None, RelativeTo=None):
        params = [
            DistributeCmd if DistributeCmd is not None else pythoncom.Missing,
            RelativeTo if RelativeTo is not None else pythoncom.Missing,
        ]
        return self.shaperange.Distribute(*params)

    def Duplicate(self):
        return self.shaperange.Duplicate()

    def Flip(self, FlipCmd=None):
        params = [
            FlipCmd if FlipCmd is not None else pythoncom.Missing,
        ]
        self.shaperange.Flip(*params)

    def Group(self):
        return self.shaperange.Group()

    def IncrementLeft(self, Increment=None):
        params = [
            Increment if Increment is not None else pythoncom.Missing,
        ]
        self.shaperange.IncrementLeft(*params)

    def IncrementRotation(self, Increment=None):
        params = [
            Increment if Increment is not None else pythoncom.Missing,
        ]
        self.shaperange.IncrementRotation(*params)

    def IncrementTop(self, Increment=None):
        params = [
            Increment if Increment is not None else pythoncom.Missing,
        ]
        self.shaperange.IncrementTop(*params)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.shaperange.Item(*params)

    def PickUp(self):
        self.shaperange.PickUp()

    def PickupAnimation(self):
        self.shaperange.PickupAnimation()

    def Regroup(self):
        return self.shaperange.Regroup()

    def RerouteConnections(self):
        self.shaperange.RerouteConnections()

    def ScaleHeight(self, Factor=None, RelativeToOriginalSize=None, fScale=None):
        params = [
            Factor if Factor is not None else pythoncom.Missing,
            RelativeToOriginalSize if RelativeToOriginalSize is not None else pythoncom.Missing,
            fScale if fScale is not None else pythoncom.Missing,
        ]
        return self.shaperange.ScaleHeight(*params)

    def ScaleWidth(self, Factor=None, RelativeToOriginalSize=None, fScale=None):
        params = [
            Factor if Factor is not None else pythoncom.Missing,
            RelativeToOriginalSize if RelativeToOriginalSize is not None else pythoncom.Missing,
            fScale if fScale is not None else pythoncom.Missing,
        ]
        self.shaperange.ScaleWidth(*params)

    def Select(self, Replace=None):
        params = [
            Replace if Replace is not None else pythoncom.Missing,
        ]
        self.shaperange.Select(*params)

    def SetShapesDefaultProperties(self):
        self.shaperange.SetShapesDefaultProperties()

    def Ungroup(self):
        return self.shaperange.Ungroup()

    def UpgradeMedia(self):
        self.shaperange.UpgradeMedia()

    def ZOrder(self, ZOrderCmd=None):
        params = [
            ZOrderCmd if ZOrderCmd is not None else pythoncom.Missing,
        ]
        self.shaperange.ZOrder(*params)


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

    def AddMediaObject(self, FileName=None, Left=None, Top=None, Width=None, Height=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            Width if Width is not None else pythoncom.Missing,
            Height if Height is not None else pythoncom.Missing,
        ]
        return self.shapes.AddMediaObject(*params)

    def AddMediaObject2(self, FileName=None, LinkToFile=None, SaveWithDocument=None, Left=None, Top=None, Width=None, Height=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            LinkToFile if LinkToFile is not None else pythoncom.Missing,
            SaveWithDocument if SaveWithDocument is not None else pythoncom.Missing,
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            Width if Width is not None else pythoncom.Missing,
            Height if Height is not None else pythoncom.Missing,
        ]
        return self.shapes.AddMediaObject2(*params)

    def AddMediaObjectFromEmbedTag(self, EmbedTag=None, Left=None, Top=None, Width=None, Height=None):
        params = [
            EmbedTag if EmbedTag is not None else pythoncom.Missing,
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            Width if Width is not None else pythoncom.Missing,
            Height if Height is not None else pythoncom.Missing,
        ]
        return self.shapes.AddMediaObjectFromEmbedTag(*params)

    def AddOLEObject(self, Left=None, Top=None, Width=None, Height=None, ClassName=None, FileName=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Link=None):
        params = [
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            Width if Width is not None else pythoncom.Missing,
            Height if Height is not None else pythoncom.Missing,
            ClassName if ClassName is not None else pythoncom.Missing,
            FileName if FileName is not None else pythoncom.Missing,
            DisplayAsIcon if DisplayAsIcon is not None else pythoncom.Missing,
            IconFileName if IconFileName is not None else pythoncom.Missing,
            IconIndex if IconIndex is not None else pythoncom.Missing,
            IconLabel if IconLabel is not None else pythoncom.Missing,
            Link if Link is not None else pythoncom.Missing,
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

    def AddPlaceholder(self, Type=None, Left=None, Top=None, Width=None, Height=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            Width if Width is not None else pythoncom.Missing,
            Height if Height is not None else pythoncom.Missing,
        ]
        return self.shapes.AddPlaceholder(*params)

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

    def AddTable(self, NumRows=None, NumColumns=None, Left=None, Top=None, Width=None, Height=None):
        params = [
            NumRows if NumRows is not None else pythoncom.Missing,
            NumColumns if NumColumns is not None else pythoncom.Missing,
            Left if Left is not None else pythoncom.Missing,
            Top if Top is not None else pythoncom.Missing,
            Width if Width is not None else pythoncom.Missing,
            Height if Height is not None else pythoncom.Missing,
        ]
        return self.shapes.AddTable(*params)

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

    def AddTitle(self):
        return self.shapes.AddTitle()

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
        return self.shapes.Item(*params)

    def Paste(self):
        return self.shapes.Paste()

    def PasteSpecial(self, DataType=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Link=None):
        params = [
            DataType if DataType is not None else pythoncom.Missing,
            DisplayAsIcon if DisplayAsIcon is not None else pythoncom.Missing,
            IconFileName if IconFileName is not None else pythoncom.Missing,
            IconIndex if IconIndex is not None else pythoncom.Missing,
            IconLabel if IconLabel is not None else pythoncom.Missing,
            Link if Link is not None else pythoncom.Missing,
        ]
        return self.shapes.PasteSpecial(*params)

    def Range(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.shapes.Range(*params)

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

    @BackgroundStyle.setter
    def BackgroundStyle(self, value):
        self.slide.BackgroundStyle = value

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

    @DisplayMasterShapes.setter
    def DisplayMasterShapes(self, value):
        self.slide.DisplayMasterShapes = value

    @property
    def FollowMasterBackground(self):
        return self.slide.FollowMasterBackground

    @FollowMasterBackground.setter
    def FollowMasterBackground(self, value):
        self.slide.FollowMasterBackground = value

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

    def ApplyTemplate(self, FileName=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
        ]
        self.slide.ApplyTemplate(*params)

    def ApplyTheme(self, themeName=None):
        params = [
            themeName if themeName is not None else pythoncom.Missing,
        ]
        self.slide.ApplyTheme(*params)

    def ApplyThemeColorScheme(self, themeColorSchemeName=None):
        params = [
            themeColorSchemeName if themeColorSchemeName is not None else pythoncom.Missing,
        ]
        self.slide.ApplyThemeColorScheme(*params)

    def Copy(self):
        self.slide.Copy()

    def Cut(self):
        self.slide.Cut()

    def Delete(self):
        self.slide.Delete()

    def Duplicate(self):
        return self.slide.Duplicate()

    def Export(self, FileName=None, FilterName=None, ScaleWidth=None, ScaleHeight=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            FilterName if FilterName is not None else pythoncom.Missing,
            ScaleWidth if ScaleWidth is not None else pythoncom.Missing,
            ScaleHeight if ScaleHeight is not None else pythoncom.Missing,
        ]
        self.slide.Export(*params)

    def MoveTo(self, toPos=None):
        params = [
            toPos if toPos is not None else pythoncom.Missing,
        ]
        self.slide.MoveTo(*params)

    def MoveToSectionStart(self, toSection=None):
        params = [
            toSection if toSection is not None else pythoncom.Missing,
        ]
        self.slide.MoveToSectionStart(*params)

    def PublishSlides(self, SlideLibraryUrl=None, Overwrite=None, UseSlideOrder=None):
        params = [
            SlideLibraryUrl if SlideLibraryUrl is not None else pythoncom.Missing,
            Overwrite if Overwrite is not None else pythoncom.Missing,
            UseSlideOrder if UseSlideOrder is not None else pythoncom.Missing,
        ]
        return self.slide.PublishSlides(*params)

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

    @BackgroundStyle.setter
    def BackgroundStyle(self, value):
        self.sliderange.BackgroundStyle = value

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

    @DisplayMasterShapes.setter
    def DisplayMasterShapes(self, value):
        self.sliderange.DisplayMasterShapes = value

    @property
    def FollowMasterBackground(self):
        return self.sliderange.FollowMasterBackground

    @FollowMasterBackground.setter
    def FollowMasterBackground(self, value):
        self.sliderange.FollowMasterBackground = value

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

    @Name.setter
    def Name(self, value):
        self.sliderange.Name = value

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

    def ApplyTemplate(self, FileName=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
        ]
        self.sliderange.ApplyTemplate(*params)

    def ApplyTheme(self, themeName=None):
        params = [
            themeName if themeName is not None else pythoncom.Missing,
        ]
        self.sliderange.ApplyTheme(*params)

    def ApplyThemeColorScheme(self, themeColorSchemeName=None):
        params = [
            themeColorSchemeName if themeColorSchemeName is not None else pythoncom.Missing,
        ]
        self.sliderange.ApplyThemeColorScheme(*params)

    def Copy(self):
        self.sliderange.Copy()

    def Cut(self):
        self.sliderange.Cut()

    def Delete(self):
        self.sliderange.Delete()

    def Duplicate(self):
        return self.sliderange.Duplicate()

    def Export(self, FileName=None, FilterName=None, ScaleWidth=None, ScaleHeight=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            FilterName if FilterName is not None else pythoncom.Missing,
            ScaleWidth if ScaleWidth is not None else pythoncom.Missing,
            ScaleHeight if ScaleHeight is not None else pythoncom.Missing,
        ]
        self.sliderange.Export(*params)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.sliderange.Item(*params)

    def MoveTo(self, toPos=None):
        params = [
            toPos if toPos is not None else pythoncom.Missing,
        ]
        self.sliderange.MoveTo(*params)

    def MoveToSectionStart(self, toSection=None):
        params = [
            toSection if toSection is not None else pythoncom.Missing,
        ]
        self.sliderange.MoveToSectionStart(*params)

    def PublishSlides(self, SlideLibraryUrl=None, Overwrite=None):
        params = [
            SlideLibraryUrl if SlideLibraryUrl is not None else pythoncom.Missing,
            Overwrite if Overwrite is not None else pythoncom.Missing,
        ]
        self.sliderange.PublishSlides(*params)

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

    def AddSlide(self, Index=None, pCustomLayout=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
            pCustomLayout if pCustomLayout is not None else pythoncom.Missing,
        ]
        return self.slides.AddSlide(*params)

    def FindBySlideID(self, SlideID=None):
        params = [
            SlideID if SlideID is not None else pythoncom.Missing,
        ]
        return self.slides.FindBySlideID(*params)

    def InsertFromFile(self, FileName=None, Index=None, SlideStart=None, SlideEnd=None):
        params = [
            FileName if FileName is not None else pythoncom.Missing,
            Index if Index is not None else pythoncom.Missing,
            SlideStart if SlideStart is not None else pythoncom.Missing,
            SlideEnd if SlideEnd is not None else pythoncom.Missing,
        ]
        return self.slides.InsertFromFile(*params)

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.slides.Item(*params)

    def Paste(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.slides.Paste(*params)

    def Range(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.slides.Range(*params)


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

    @LoopUntilStopped.setter
    def LoopUntilStopped(self, value):
        self.slideshowsettings.LoopUntilStopped = value

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

    @ShowMediaControls.setter
    def ShowMediaControls(self, value):
        self.slideshowsettings.ShowMediaControls = value

    @property
    def ShowPresenterView(self):
        return SlideShowSettings(self.slideshowsettings.ShowPresenterView)

    @ShowPresenterView.setter
    def ShowPresenterView(self, value):
        self.slideshowsettings.ShowPresenterView = value

    @property
    def ShowScrollbar(self):
        return self.slideshowsettings.ShowScrollbar

    @ShowScrollbar.setter
    def ShowScrollbar(self, value):
        self.slideshowsettings.ShowScrollbar = value

    @property
    def ShowType(self):
        return self.slideshowsettings.ShowType

    @ShowType.setter
    def ShowType(self, value):
        self.slideshowsettings.ShowType = value

    @property
    def ShowWithAnimation(self):
        return self.slideshowsettings.ShowWithAnimation

    @ShowWithAnimation.setter
    def ShowWithAnimation(self, value):
        self.slideshowsettings.ShowWithAnimation = value

    @property
    def ShowWithNarration(self):
        return self.slideshowsettings.ShowWithNarration

    @ShowWithNarration.setter
    def ShowWithNarration(self, value):
        self.slideshowsettings.ShowWithNarration = value

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

    @AdvanceOnClick.setter
    def AdvanceOnClick(self, value):
        self.slideshowtransition.AdvanceOnClick = value

    @property
    def AdvanceOnTime(self):
        return self.slideshowtransition.AdvanceOnTime

    @AdvanceOnTime.setter
    def AdvanceOnTime(self, value):
        self.slideshowtransition.AdvanceOnTime = value

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

    @Hidden.setter
    def Hidden(self, value):
        self.slideshowtransition.Hidden = value

    @property
    def LoopSoundUntilNext(self):
        return self.slideshowtransition.LoopSoundUntilNext

    @LoopSoundUntilNext.setter
    def LoopSoundUntilNext(self, value):
        self.slideshowtransition.LoopSoundUntilNext = value

    @property
    def Parent(self):
        return self.slideshowtransition.Parent

    @property
    def SoundEffect(self):
        return SoundEffect(self.slideshowtransition.SoundEffect)

    @property
    def Speed(self):
        return self.slideshowtransition.Speed

    @Speed.setter
    def Speed(self, value):
        self.slideshowtransition.Speed = value


class SlideShowView:

    def __init__(self, slideshowview=None):
        self.slideshowview = slideshowview

    @property
    def AcceleratorsEnabled(self):
        return self.slideshowview.AcceleratorsEnabled

    @AcceleratorsEnabled.setter
    def AcceleratorsEnabled(self, value):
        self.slideshowview.AcceleratorsEnabled = value

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

    @SlideElapsedTime.setter
    def SlideElapsedTime(self, value):
        self.slideshowview.SlideElapsedTime = value

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

    def DrawLine(self, BeginX=None, BeginY=None, EndX=None, EndY=None):
        params = [
            BeginX if BeginX is not None else pythoncom.Missing,
            BeginY if BeginY is not None else pythoncom.Missing,
            EndX if EndX is not None else pythoncom.Missing,
            EndY if EndY is not None else pythoncom.Missing,
        ]
        self.slideshowview.DrawLine(*params)

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

    def GotoClick(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        self.slideshowview.GotoClick(*params)

    def GotoNamedShow(self, SlideShowName=None):
        params = [
            SlideShowName if SlideShowName is not None else pythoncom.Missing,
        ]
        self.slideshowview.GotoNamedShow(*params)

    def GotoSlide(self, Index=None, ResetSlide=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
            ResetSlide if ResetSlide is not None else pythoncom.Missing,
        ]
        self.slideshowview.GotoSlide(*params)

    def Last(self):
        self.slideshowview.Last()

    def Next(self):
        self.slideshowview.Next()

    def Player(self, ShapeId=None):
        params = [
            ShapeId if ShapeId is not None else pythoncom.Missing,
        ]
        return self.slideshowview.Player(*params)

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

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.slideshowwindows.Item(*params)


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

    @Type.setter
    def Type(self, value):
        self.soundeffect.Type = value

    def ImportFromFile(self, FullName=None):
        params = [
            FullName if FullName is not None else pythoncom.Missing,
        ]
        self.soundeffect.ImportFromFile(*params)

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

    @FirstCol.setter
    def FirstCol(self, value):
        self.table.FirstCol = value

    @property
    def FirstRow(self):
        return self.table.FirstRow

    @FirstRow.setter
    def FirstRow(self, value):
        self.table.FirstRow = value

    @property
    def HorizBanding(self):
        return self.table.HorizBanding

    @HorizBanding.setter
    def HorizBanding(self, value):
        self.table.HorizBanding = value

    @property
    def LastCol(self):
        return self.table.LastCol

    @LastCol.setter
    def LastCol(self, value):
        self.table.LastCol = value

    @property
    def LastRow(self):
        return self.table.LastRow

    @LastRow.setter
    def LastRow(self, value):
        self.table.LastRow = value

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

    @VertBanding.setter
    def VertBanding(self, value):
        self.table.VertBanding = value

    def ApplyStyle(self, StyleID=None, SaveFormatting=None):
        params = [
            StyleID if StyleID is not None else pythoncom.Missing,
            SaveFormatting if SaveFormatting is not None else pythoncom.Missing,
        ]
        self.table.ApplyStyle(*params)

    def Cell(self, Row=None, Column=None):
        params = [
            Row if Row is not None else pythoncom.Missing,
            Column if Column is not None else pythoncom.Missing,
        ]
        return self.table.Cell(*params)

    def ScaleProportionally(self, scale=None):
        params = [
            scale if scale is not None else pythoncom.Missing,
        ]
        self.table.ScaleProportionally(*params)


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

    @Type.setter
    def Type(self, value):
        self.tabstop.Type = value

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

    def Add(self, Type=None, Position=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
            Position if Position is not None else pythoncom.Missing,
        ]
        return TabStop(self.tabstops.Add(*params))

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.tabstops.Item(*params)


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

    def Add(self, Name=None, Value=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
            Value if Value is not None else pythoncom.Missing,
        ]
        self.tags.Add(*params)

    def Delete(self, Name=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
        ]
        self.tags.Delete(*params)

    def Item(self, Name=None):
        params = [
            Name if Name is not None else pythoncom.Missing,
        ]
        return self.tags.Item(*params)

    def Name(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.tags.Name(*params)

    def Value(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.tags.Value(*params)


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

    @WordWrap.setter
    def WordWrap(self, value):
        self.textframe.WordWrap = value

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
    def WordArtFormat(self):
        return self.textframe2.WordArtFormat

    @WordArtFormat.setter
    def WordArtFormat(self, value):
        self.textframe2.WordArtFormat = value

    @property
    def WordWrap(self):
        return self.textframe2.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.textframe2.WordWrap = value

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

    def ChangeCase(self, Type=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
        ]
        self.textrange.ChangeCase(*params)

    def Characters(self, Start=None, Length=None):
        params = [
            Start if Start is not None else pythoncom.Missing,
            Length if Length is not None else pythoncom.Missing,
        ]
        return self.textrange.Characters(*params)

    def Copy(self):
        self.textrange.Copy()

    def Cut(self):
        self.textrange.Cut()

    def Delete(self):
        self.textrange.Delete()

    def Find(self, FindWhat=None, After=None, MatchCase=None, WholeWords=None):
        params = [
            FindWhat if FindWhat is not None else pythoncom.Missing,
            After if After is not None else pythoncom.Missing,
            MatchCase if MatchCase is not None else pythoncom.Missing,
            WholeWords if WholeWords is not None else pythoncom.Missing,
        ]
        return self.textrange.Find(*params)

    def InsertAfter(self, NewText=None):
        params = [
            NewText if NewText is not None else pythoncom.Missing,
        ]
        self.textrange.InsertAfter(*params)

    def InsertBefore(self, NewText=None):
        params = [
            NewText if NewText is not None else pythoncom.Missing,
        ]
        self.textrange.InsertBefore(*params)

    def InsertDateTime(self, DateTimeFormat=None, InsertAsField=None):
        params = [
            DateTimeFormat if DateTimeFormat is not None else pythoncom.Missing,
            InsertAsField if InsertAsField is not None else pythoncom.Missing,
        ]
        return self.textrange.InsertDateTime(*params)

    def InsertSlideNumber(self):
        return self.textrange.InsertSlideNumber()

    def InsertSymbol(self, FontName=None, CharNumber=None, UniCode=None):
        params = [
            FontName if FontName is not None else pythoncom.Missing,
            CharNumber if CharNumber is not None else pythoncom.Missing,
            UniCode if UniCode is not None else pythoncom.Missing,
        ]
        return self.textrange.InsertSymbol(*params)

    def Lines(self, Start=None, Length=None):
        params = [
            Start if Start is not None else pythoncom.Missing,
            Length if Length is not None else pythoncom.Missing,
        ]
        return self.textrange.Lines(*params)

    def LtrRun(self):
        self.textrange.LtrRun()

    def Paragraphs(self, Start=None, Length=None):
        params = [
            Start if Start is not None else pythoncom.Missing,
            Length if Length is not None else pythoncom.Missing,
        ]
        return self.textrange.Paragraphs(*params)

    def Paste(self):
        return self.textrange.Paste()

    def PasteSpecial(self, DataType=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Link=None):
        params = [
            DataType if DataType is not None else pythoncom.Missing,
            DisplayAsIcon if DisplayAsIcon is not None else pythoncom.Missing,
            IconFileName if IconFileName is not None else pythoncom.Missing,
            IconIndex if IconIndex is not None else pythoncom.Missing,
            IconLabel if IconLabel is not None else pythoncom.Missing,
            Link if Link is not None else pythoncom.Missing,
        ]
        return self.textrange.PasteSpecial(*params)

    def RemovePeriods(self):
        self.textrange.RemovePeriods()

    def Replace(self, FindWhat=None, ReplaceWhat=None, After=None, MatchCase=None, WholeWords=None):
        params = [
            FindWhat if FindWhat is not None else pythoncom.Missing,
            ReplaceWhat if ReplaceWhat is not None else pythoncom.Missing,
            After if After is not None else pythoncom.Missing,
            MatchCase if MatchCase is not None else pythoncom.Missing,
            WholeWords if WholeWords is not None else pythoncom.Missing,
        ]
        return self.textrange.Replace(*params)

    def RotatedBounds(self, X1=None, Y1=None, X2=None, Y2=None, X3=None, Y3=None, X4=None, Y4=None):
        params = [
            X1 if X1 is not None else pythoncom.Missing,
            Y1 if Y1 is not None else pythoncom.Missing,
            X2 if X2 is not None else pythoncom.Missing,
            Y2 if Y2 is not None else pythoncom.Missing,
            X3 if X3 is not None else pythoncom.Missing,
            Y3 if Y3 is not None else pythoncom.Missing,
            X4 if X4 is not None else pythoncom.Missing,
            Y4 if Y4 is not None else pythoncom.Missing,
        ]
        self.textrange.RotatedBounds(*params)

    def RtlRun(self):
        self.textrange.RtlRun()

    def Runs(self, Start=None, Length=None):
        params = [
            Start if Start is not None else pythoncom.Missing,
            Length if Length is not None else pythoncom.Missing,
        ]
        return self.textrange.Runs(*params)

    def Select(self):
        self.textrange.Select()

    def Sentences(self, Start=None, Length=None):
        params = [
            Start if Start is not None else pythoncom.Missing,
            Length if Length is not None else pythoncom.Missing,
        ]
        return self.textrange.Sentences(*params)

    def TrimText(self):
        return self.textrange.TrimText()

    def Words(self, Start=None, Length=None):
        params = [
            Start if Start is not None else pythoncom.Missing,
            Length if Length is not None else pythoncom.Missing,
        ]
        return self.textrange.Words(*params)


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

    def Item(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        return self.textstylelevels.Item(*params)


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

    def Item(self, Type=None):
        params = [
            Type if Type is not None else pythoncom.Missing,
        ]
        return self.textstyles.Item(*params)


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

    @AutoReverse.setter
    def AutoReverse(self, value):
        self.timing.AutoReverse = value

    @property
    def BounceEnd(self):
        return self.timing.BounceEnd

    @BounceEnd.setter
    def BounceEnd(self, value):
        self.timing.BounceEnd = value

    @property
    def BounceEndIntensity(self):
        return self.timing.BounceEndIntensity

    @BounceEndIntensity.setter
    def BounceEndIntensity(self, value):
        self.timing.BounceEndIntensity = value

    @property
    def Decelerate(self):
        return self.timing.Decelerate

    @Decelerate.setter
    def Decelerate(self, value):
        self.timing.Decelerate = value

    @property
    def Parent(self):
        return self.timing.Parent

    @property
    def RepeatCount(self):
        return self.timing.RepeatCount

    @RepeatCount.setter
    def RepeatCount(self, value):
        self.timing.RepeatCount = value

    @property
    def RepeatDuration(self):
        return self.timing.RepeatDuration

    @RepeatDuration.setter
    def RepeatDuration(self, value):
        self.timing.RepeatDuration = value

    @property
    def Restart(self):
        return self.timing.Restart

    @Restart.setter
    def Restart(self, value):
        self.timing.Restart = value

    @property
    def RewindAtEnd(self):
        return self.timing.RewindAtEnd

    @RewindAtEnd.setter
    def RewindAtEnd(self, value):
        self.timing.RewindAtEnd = value

    @property
    def SmoothEnd(self):
        return self.timing.SmoothEnd

    @SmoothEnd.setter
    def SmoothEnd(self, value):
        self.timing.SmoothEnd = value

    @property
    def SmoothStart(self):
        return self.timing.SmoothStart

    @SmoothStart.setter
    def SmoothStart(self, value):
        self.timing.SmoothStart = value

    @property
    def Speed(self):
        return self.timing.Speed

    @Speed.setter
    def Speed(self, value):
        self.timing.Speed = value

    @property
    def triggerBookmark(self):
        return self.timing.triggerBookmark

    @triggerBookmark.setter
    def triggerBookmark(self, value):
        self.timing.triggerBookmark = value

    @property
    def TriggerDelayTime(self):
        return self.timing.TriggerDelayTime

    @TriggerDelayTime.setter
    def TriggerDelayTime(self, value):
        self.timing.TriggerDelayTime = value

    @property
    def TriggerShape(self):
        return self.timing.TriggerShape

    @TriggerShape.setter
    def TriggerShape(self, value):
        self.timing.TriggerShape = value

    @property
    def TriggerType(self):
        return self.timing.TriggerType

    @TriggerType.setter
    def TriggerType(self, value):
        self.timing.TriggerType = value


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

    @DisplaySlideMiniature.setter
    def DisplaySlideMiniature(self, value):
        self.view.DisplaySlideMiniature = value

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

    @ZoomToFit.setter
    def ZoomToFit(self, value):
        self.view.ZoomToFit = value

    def GotoSlide(self, Index=None):
        params = [
            Index if Index is not None else pythoncom.Missing,
        ]
        self.view.GotoSlide(*params)

    def Paste(self):
        self.view.Paste()

    def PasteSpecial(self, DataType=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Link=None):
        params = [
            DataType if DataType is not None else pythoncom.Missing,
            DisplayAsIcon if DisplayAsIcon is not None else pythoncom.Missing,
            IconFileName if IconFileName is not None else pythoncom.Missing,
            IconIndex if IconIndex is not None else pythoncom.Missing,
            IconLabel if IconLabel is not None else pythoncom.Missing,
            Link if Link is not None else pythoncom.Missing,
        ]
        self.view.PasteSpecial(*params)

    def Player(self, ShapeId=None):
        params = [
            ShapeId if ShapeId is not None else pythoncom.Missing,
        ]
        return self.view.Player(*params)

    def PrintOut(self, From=None, To=None, PrintToFile=None, Copies=None, Collate=None):
        params = [
            From if From is not None else pythoncom.Missing,
            To if To is not None else pythoncom.Missing,
            PrintToFile if PrintToFile is not None else pythoncom.Missing,
            Copies if Copies is not None else pythoncom.Missing,
            Collate if Collate is not None else pythoncom.Missing,
        ]
        self.view.PrintOut(*params)


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


# XlAxisCrosses enumeration
xlAxisCrossesAutomatic = -4105
xlAxisCrossesCustom = -4114
xlAxisCrossesMaximum = 2
xlAxisCrossesMinimum = 4

# XlAxisGroup enumeration
xlPrimary = 1
xlSecondary = 2

# XlAxisType enumeration
xlCategory = 1
xlSeriesAxis = 3
xlValue = 2

# XlBackground enumeration
xlBackgroundAutomatic = -4105
xlBackgroundOpaque = 3
xlBackgroundTransparent = 2

# XlBarShape enumeration
xlBox = 0
xlConeToMax = 5
xlConeToPoint = 4
xlCylinder = 3
xlPyramidToMax = 2
xlPyramidToPoint = 1

# XlBorderWeight enumeration
xlHairline = 1
xlMedium = -4138
xlThick = 4
xlThin = 2

# XlCategoryType enumeration
xlAutomaticScale = -4105
xlCategoryScale = 2
xlTimeScale = 3

# XlChartElementPosition enumeration
xlChartElementPositionAutomatic = -4105
xlChartElementPositionCustom = -4114

# XlChartGallery enumeration
xlAnyGallery = 23
xlBuiltIn = 21
xlUserDefined = 22

# XlChartItem enumeration
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

# XlColorIndex enumeration
xlColorIndexAutomatic = -4105
xlColorIndexNone = -4142

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

# XlCopyPictureFormat enumeration
xlBitmap = 2
xlPicture = -4147

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

# XlDisplayBlanksAs enumeration
xlInterpolated = 3
xlNotPlotted = 1
xlZero = 2

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

# XlEndStyleCap enumeration
xlCap = 1
xlNoCap = 2

# XlErrorBarDirection enumeration
xlChartX = -4168
xlChartY = 1

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

# XlHAlign enumeration
xlHAlignCenter = -4108
xlHAlignCenterAcrossSelection = 7
xlHAlignDistributed = -4117
xlHAlignFill = 5
xlHAlignGeneral = 1
xlHAlignJustify = -4130
xlHAlignLeft = -4131
xlHAlignRight = -4152

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

# XlOrientation enumeration
xlDownward = -4170
xlHorizontal = -4128
xlUpward = -4171
xlVertical = -4166

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
xlPatternLinearGradient = 4000
xlPatternNone = -4142
xlPatternRectangularGradient = 4001
xlPatternSemiGray75 = 10
xlPatternSolid = 1
xlPatternUp = -4162
xlPatternVertical = -4166

# XlPictureAppearance enumeration
xlPrinter = 2
xlScreen = 1

# XlPivotFieldOrientation enumeration
xlColumnField = 2
xlDataField = 4
xlHidden = 0
xlPageField = 3
xlRowField = 1

# XlReadingOrder enumeration
xlContext = -5002
xlLTR = -5003
xlRTL = -5004

# XlRgbColor enumeration
xlAliceBlue = 16775408
xlAntiqueWhite = 14150650
xlAqua = 16776960
xlAquamarine = 13959039
xlAzure = 16777200
xlBeige = 14480885
xlBisque = 12903679
xlBlack = 0
xlBlanchedAlmond = 13495295
xlBlue = 16711680
xlBlueViolet = 14822282
xlBrown = 2763429
xlBurlyWood = 8894686
xlCadetBlue = 10526303
xlChartreuse = 65407
xlCoral = 5275647
xlCornflowerBlue = 15570276
xlCornsilk = 14481663
xlCrimson = 3937500
xlDarkBlue = 9109504
xlDarkCyan = 9145088
xlDarkGoldenrod = 755384
xlDarkGray = 11119017
xlDarkGreen = 25600
xlDarkGrey = 11119017
xlDarkKhaki = 7059389
xlDarkMagenta = 9109643
xlDarkOliveGreen = 3107669
xlDarkOrange = 36095
xlDarkOrchid = 13382297
xlDarkRed = 139
xlDarkSalmon = 8034025
xlDarkSeaGreen = 9419919
xlDarkSlateBlue = 9125192
xlDarkSlateGray = 5197615
xlDarkSlateGrey = 5197615
xlDarkTurquoise = 13749760
xlDarkViolet = 13828244
xlDeepPink = 9639167
xlDeepSkyBlue = 16760576
xlDimGray = 6908265
xlDimGrey = 6908265
xlDodgerBlue = 16748574
xlFireBrick = 2237106
xlFloralWhite = 15792895
xlForestGreen = 2263842
xlFuchsia = 16711935
xlGainsboro = 14474460
xlGhostWhite = 16775416
xlGold = 55295
xlGoldenrod = 2139610
xlGray = 8421504
xlGreen = 32768
xlGreenYellow = 3145645
xlGrey = 8421504
xlHoneydew = 15794160
xlHotPink = 11823615
xlIndianRed = 6053069
xlIndigo = 8519755
xlIvory = 15794175
xlKhaki = 9234160
xlLavender = 16443110
xlLavenderBlush = 16118015
xlLawnGreen = 64636
xlLemonChiffon = 13499135
xlLightBlue = 15128749
xlLightCoral = 8421616
xlLightCyan = 9145088
xlLightGoldenrodYellow = 13826810
xlLightGray = 13882323
xlLightGreen = 9498256
xlLightGrey = 13882323
xlLightPink = 12695295
xlLightSalmon = 8036607
xlLightSeaGreen = 11186720
xlLightSkyBlue = 16436871
xlLightSlateGray = 10061943
xlLightSlateGrey = 10061943
xlLightSteelBlue = 14599344
xlLightYellow = 14745599
xlLime = 65280
xlLimeGreen = 3329330
xlLinen = 15134970
xlMaroon = 128
xlMediumAquamarine = 11206502
xlMediumBlue = 13434880
xlMediumOrchid = 13850042
xlMediumPurple = 14381203
xlMediumSeaGreen = 7451452
xlMediumSlateBlue = 15624315
xlMediumSpringGreen = 10156544
xlMediumTurquoise = 13422920
xlMediumVioletRed = 8721863
xlMidnightBlue = 7346457
xlMintCream = 16449525
xlMistyRose = 14804223
xlMoccasin = 11920639
xlNavajoWhite = 11394815
xlNavy = 8388608
xlNavyBlue = 8388608
xlOldLace = 15136253
xlOlive = 32896
xlOliveDrab = 2330219
xlOrange = 42495
xlOrangeRed = 17919
xlOrchid = 14053594
xlPaleGoldenrod = 7071982
xlPaleGreen = 10025880
xlPaleTurquoise = 15658671
xlPaleVioletRed = 9662683
xlPapayaWhip = 14020607
xlPeachPuff = 12180223
xlPeru = 4163021
xlPink = 13353215
xlPlum = 14524637
xlPowderBlue = 15130800
xlPurple = 8388736
xlRed = 255
xlRosyBrown = 9408444
xlRoyalBlue = 14772545
xlSalmon = 7504122
xlSandyBrown = 6333684
xlSeaGreen = 5737262
xlSeashell = 15660543
xlSienna = 2970272
xlSilver = 12632256
xlSkyBlue = 15453831
xlSlateBlue = 13458026
xlSlateGray = 9470064
xlSlateGrey = 9470064
xlSnow = 16448255
xlSpringGreen = 8388352
xlSteelBlue = 11829830
xlTan = 9221330
xlTeal = 8421376
xlThistle = 14204888
xlTomato = 4678655
xlTurquoise = 13688896
xlViolet = 15631086
xlWheat = 11788021
xlWhite = 16777215
xlWhiteSmoke = 16119285
xlYellow = 65535
xlYellowGreen = 3329434

# XlRowCol enumeration
xlColumns = 2
xlRows = 1

# XlScaleType enumeration
xlScaleLinear = -4132
xlScaleLogarithmic = -4133

# XlSizeRepresents enumeration
xlSizeIsArea = 1
xlSizeIsWidth = 2

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

# XlTimeUnit enumeration
xlDays = 0
xlMonths = 1
xlYears = 2

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

# XlVAlign enumeration
xlVAlignBottom = -4107
xlVAlignCenter = -4108
xlVAlignDistributed = -4117
xlVAlignJustify = -4130
xlVAlignTop = -4160
