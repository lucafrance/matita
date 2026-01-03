from . import com_arguments

import win32com.client
import pythoncom

class ActionSetting:

    def __init__(self, actionsetting=None):
        self.actionsetting = actionsetting

    @property
    def Action(self):
        return self.actionsetting.Action

    # Lower case alias for Action
    @property
    def action(self):
        return self.Action

    @Action.setter
    def Action(self, value):
        self.actionsetting.Action = value

    # Lower case alias for Action setter
    @action.setter
    def action(self, value):
        self.Action = value

    @property
    def ActionVerb(self):
        return self.actionsetting.ActionVerb

    # Lower case alias for ActionVerb
    @property
    def actionverb(self):
        return self.ActionVerb

    @ActionVerb.setter
    def ActionVerb(self, value):
        self.actionsetting.ActionVerb = value

    # Lower case alias for ActionVerb setter
    @actionverb.setter
    def actionverb(self, value):
        self.ActionVerb = value

    @property
    def AnimateAction(self):
        return self.actionsetting.AnimateAction

    # Lower case alias for AnimateAction
    @property
    def animateaction(self):
        return self.AnimateAction

    @AnimateAction.setter
    def AnimateAction(self, value):
        self.actionsetting.AnimateAction = value

    # Lower case alias for AnimateAction setter
    @animateaction.setter
    def animateaction(self, value):
        self.AnimateAction = value

    @property
    def Application(self):
        return Application(self.actionsetting.Application)

    @property
    def Hyperlink(self):
        return Hyperlink(self.actionsetting.Hyperlink)

    # Lower case alias for Hyperlink
    @property
    def hyperlink(self):
        return self.Hyperlink

    @property
    def Parent(self):
        return self.actionsetting.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Run(self):
        return self.actionsetting.Run

    # Lower case alias for Run
    @property
    def run(self):
        return self.Run

    @Run.setter
    def Run(self, value):
        self.actionsetting.Run = value

    # Lower case alias for Run setter
    @run.setter
    def run(self, value):
        self.Run = value

    @property
    def ShowAndReturn(self):
        return self.actionsetting.ShowAndReturn

    # Lower case alias for ShowAndReturn
    @property
    def showandreturn(self):
        return self.ShowAndReturn

    @ShowAndReturn.setter
    def ShowAndReturn(self, value):
        self.actionsetting.ShowAndReturn = value

    # Lower case alias for ShowAndReturn setter
    @showandreturn.setter
    def showandreturn(self, value):
        self.ShowAndReturn = value

    @property
    def SlideShowName(self):
        return self.actionsetting.SlideShowName

    # Lower case alias for SlideShowName
    @property
    def slideshowname(self):
        return self.SlideShowName

    @SlideShowName.setter
    def SlideShowName(self, value):
        self.actionsetting.SlideShowName = value

    # Lower case alias for SlideShowName setter
    @slideshowname.setter
    def slideshowname(self, value):
        self.SlideShowName = value

    @property
    def SoundEffect(self):
        return SoundEffect(self.actionsetting.SoundEffect)

    # Lower case alias for SoundEffect
    @property
    def soundeffect(self):
        return self.SoundEffect


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.actionsettings.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.actionsettings.Item(*arguments)


class AddIn:

    def __init__(self, addin=None):
        self.addin = addin

    @property
    def Application(self):
        return Application(self.addin.Application)

    @property
    def AutoLoad(self):
        return self.addin.AutoLoad

    # Lower case alias for AutoLoad
    @property
    def autoload(self):
        return self.AutoLoad

    @AutoLoad.setter
    def AutoLoad(self, value):
        self.addin.AutoLoad = value

    # Lower case alias for AutoLoad setter
    @autoload.setter
    def autoload(self, value):
        self.AutoLoad = value

    @property
    def FullName(self):
        return self.addin.FullName

    # Lower case alias for FullName
    @property
    def fullname(self):
        return self.FullName

    @property
    def Loaded(self):
        return self.addin.Loaded

    # Lower case alias for Loaded
    @property
    def loaded(self):
        return self.Loaded

    @Loaded.setter
    def Loaded(self, value):
        self.addin.Loaded = value

    # Lower case alias for Loaded setter
    @loaded.setter
    def loaded(self, value):
        self.Loaded = value

    @property
    def Name(self):
        return self.addin.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.addin.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Path(self):
        return AddIn(self.addin.Path)

    # Lower case alias for Path
    @property
    def path(self):
        return self.Path

    @property
    def Registered(self):
        return self.addin.Registered

    # Lower case alias for Registered
    @property
    def registered(self):
        return self.Registered

    @Registered.setter
    def Registered(self, value):
        self.addin.Registered = value

    # Lower case alias for Registered setter
    @registered.setter
    def registered(self, value):
        self.Registered = value


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.addins.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Add(self, FileName=None):
        arguments = com_arguments([FileName])
        return AddIn(self.addins.Add(*arguments))

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.addins.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.addins.Remove(*arguments)


class Adjustments:

    def __init__(self, adjustments=None):
        self.adjustments = adjustments

    @property
    def Application(self):
        return Application(self.adjustments.Application)

    @property
    def Count(self):
        return self.adjustments.Count

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.adjustments.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Item(self):
        return self.adjustments.Item

    # Lower case alias for Item
    @property
    def item(self):
        return self.Item

    @Item.setter
    def Item(self, value):
        self.adjustments.Item = value

    # Lower case alias for Item setter
    @item.setter
    def item(self, value):
        self.Item = value

    @property
    def Parent(self):
        return self.adjustments.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent


class AnimationBehavior:

    def __init__(self, animationbehavior=None):
        self.animationbehavior = animationbehavior

    @property
    def Accumulate(self):
        return self.animationbehavior.Accumulate

    # Lower case alias for Accumulate
    @property
    def accumulate(self):
        return self.Accumulate

    @Accumulate.setter
    def Accumulate(self, value):
        self.animationbehavior.Accumulate = value

    # Lower case alias for Accumulate setter
    @accumulate.setter
    def accumulate(self, value):
        self.Accumulate = value

    @property
    def Additive(self):
        return self.animationbehavior.Additive

    # Lower case alias for Additive
    @property
    def additive(self):
        return self.Additive

    @Additive.setter
    def Additive(self, value):
        self.animationbehavior.Additive = value

    # Lower case alias for Additive setter
    @additive.setter
    def additive(self, value):
        self.Additive = value

    @property
    def Application(self):
        return Application(self.animationbehavior.Application)

    @property
    def ColorEffect(self):
        return ColorEffect(self.animationbehavior.ColorEffect)

    # Lower case alias for ColorEffect
    @property
    def coloreffect(self):
        return self.ColorEffect

    @property
    def CommandEffect(self):
        return CommandEffect(self.animationbehavior.CommandEffect)

    # Lower case alias for CommandEffect
    @property
    def commandeffect(self):
        return self.CommandEffect

    @property
    def FilterEffect(self):
        return FilterEffect(self.animationbehavior.FilterEffect)

    # Lower case alias for FilterEffect
    @property
    def filtereffect(self):
        return self.FilterEffect

    @property
    def MotionEffect(self):
        return MotionEffect(self.animationbehavior.MotionEffect)

    # Lower case alias for MotionEffect
    @property
    def motioneffect(self):
        return self.MotionEffect

    @property
    def Parent(self):
        return self.animationbehavior.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyEffect(self):
        return PropertyEffect(self.animationbehavior.PropertyEffect)

    # Lower case alias for PropertyEffect
    @property
    def propertyeffect(self):
        return self.PropertyEffect

    @property
    def RotationEffect(self):
        return RotationEffect(self.animationbehavior.RotationEffect)

    # Lower case alias for RotationEffect
    @property
    def rotationeffect(self):
        return self.RotationEffect

    @property
    def ScaleEffect(self):
        return ScaleEffect(self.animationbehavior.ScaleEffect)

    # Lower case alias for ScaleEffect
    @property
    def scaleeffect(self):
        return self.ScaleEffect

    @property
    def SetEffect(self):
        return SetEffect(self.animationbehavior.SetEffect)

    # Lower case alias for SetEffect
    @property
    def seteffect(self):
        return self.SetEffect

    @property
    def Timing(self):
        return Timing(self.animationbehavior.Timing)

    # Lower case alias for Timing
    @property
    def timing(self):
        return self.Timing

    @property
    def Type(self):
        return self.animationbehavior.Type

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @Type.setter
    def Type(self, value):
        self.animationbehavior.Type = value

    # Lower case alias for Type setter
    @type.setter
    def type(self, value):
        self.Type = value

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.animationbehaviors.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Add(self, Type=None, Index=None):
        arguments = com_arguments([Type, Index])
        return self.animationbehaviors.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.animationbehaviors.Item(*arguments)


class AnimationPoint:

    def __init__(self, animationpoint=None):
        self.animationpoint = animationpoint

    @property
    def Application(self):
        return Application(self.animationpoint.Application)

    @property
    def Formula(self):
        return self.animationpoint.Formula

    # Lower case alias for Formula
    @property
    def formula(self):
        return self.Formula

    @Formula.setter
    def Formula(self, value):
        self.animationpoint.Formula = value

    # Lower case alias for Formula setter
    @formula.setter
    def formula(self, value):
        self.Formula = value

    @property
    def Parent(self):
        return self.animationpoint.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Time(self):
        return self.animationpoint.Time

    # Lower case alias for Time
    @property
    def time(self):
        return self.Time

    @Time.setter
    def Time(self, value):
        self.animationpoint.Time = value

    # Lower case alias for Time setter
    @time.setter
    def time(self, value):
        self.Time = value

    @property
    def Value(self):
        return self.animationpoint.Value

    # Lower case alias for Value
    @property
    def value(self):
        return self.Value

    @Value.setter
    def Value(self, value):
        self.animationpoint.Value = value

    # Lower case alias for Value setter
    @value.setter
    def value(self, value):
        self.Value = value

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.animationpoints.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Smooth(self):
        return self.animationpoints.Smooth

    # Lower case alias for Smooth
    @property
    def smooth(self):
        return self.Smooth

    @Smooth.setter
    def Smooth(self, value):
        self.animationpoints.Smooth = value

    # Lower case alias for Smooth setter
    @smooth.setter
    def smooth(self, value):
        self.Smooth = value

    def Add(self, Index=None):
        arguments = com_arguments([Index])
        return self.animationpoints.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.animationpoints.Item(*arguments)


class AnimationSettings:

    def __init__(self, animationsettings=None):
        self.animationsettings = animationsettings

    @property
    def AdvanceMode(self):
        return self.animationsettings.AdvanceMode

    # Lower case alias for AdvanceMode
    @property
    def advancemode(self):
        return self.AdvanceMode

    @AdvanceMode.setter
    def AdvanceMode(self, value):
        self.animationsettings.AdvanceMode = value

    # Lower case alias for AdvanceMode setter
    @advancemode.setter
    def advancemode(self, value):
        self.AdvanceMode = value

    @property
    def AdvanceTime(self):
        return self.animationsettings.AdvanceTime

    # Lower case alias for AdvanceTime
    @property
    def advancetime(self):
        return self.AdvanceTime

    @AdvanceTime.setter
    def AdvanceTime(self, value):
        self.animationsettings.AdvanceTime = value

    # Lower case alias for AdvanceTime setter
    @advancetime.setter
    def advancetime(self, value):
        self.AdvanceTime = value

    @property
    def AfterEffect(self):
        return PpAfterEffect(self.animationsettings.AfterEffect)

    # Lower case alias for AfterEffect
    @property
    def aftereffect(self):
        return self.AfterEffect

    @AfterEffect.setter
    def AfterEffect(self, value):
        self.animationsettings.AfterEffect = value

    # Lower case alias for AfterEffect setter
    @aftereffect.setter
    def aftereffect(self, value):
        self.AfterEffect = value

    @property
    def Animate(self):
        return self.animationsettings.Animate

    # Lower case alias for Animate
    @property
    def animate(self):
        return self.Animate

    @Animate.setter
    def Animate(self, value):
        self.animationsettings.Animate = value

    # Lower case alias for Animate setter
    @animate.setter
    def animate(self, value):
        self.Animate = value

    @property
    def AnimateBackground(self):
        return self.animationsettings.AnimateBackground

    # Lower case alias for AnimateBackground
    @property
    def animatebackground(self):
        return self.AnimateBackground

    @AnimateBackground.setter
    def AnimateBackground(self, value):
        self.animationsettings.AnimateBackground = value

    # Lower case alias for AnimateBackground setter
    @animatebackground.setter
    def animatebackground(self, value):
        self.AnimateBackground = value

    @property
    def AnimateTextInReverse(self):
        return self.animationsettings.AnimateTextInReverse

    # Lower case alias for AnimateTextInReverse
    @property
    def animatetextinreverse(self):
        return self.AnimateTextInReverse

    @AnimateTextInReverse.setter
    def AnimateTextInReverse(self, value):
        self.animationsettings.AnimateTextInReverse = value

    # Lower case alias for AnimateTextInReverse setter
    @animatetextinreverse.setter
    def animatetextinreverse(self, value):
        self.AnimateTextInReverse = value

    @property
    def AnimationOrder(self):
        return self.animationsettings.AnimationOrder

    # Lower case alias for AnimationOrder
    @property
    def animationorder(self):
        return self.AnimationOrder

    @AnimationOrder.setter
    def AnimationOrder(self, value):
        self.animationsettings.AnimationOrder = value

    # Lower case alias for AnimationOrder setter
    @animationorder.setter
    def animationorder(self, value):
        self.AnimationOrder = value

    @property
    def Application(self):
        return Application(self.animationsettings.Application)

    @property
    def ChartUnitEffect(self):
        return self.animationsettings.ChartUnitEffect

    # Lower case alias for ChartUnitEffect
    @property
    def chartuniteffect(self):
        return self.ChartUnitEffect

    @ChartUnitEffect.setter
    def ChartUnitEffect(self, value):
        self.animationsettings.ChartUnitEffect = value

    # Lower case alias for ChartUnitEffect setter
    @chartuniteffect.setter
    def chartuniteffect(self, value):
        self.ChartUnitEffect = value

    @property
    def DimColor(self):
        return ColorFormat(self.animationsettings.DimColor)

    # Lower case alias for DimColor
    @property
    def dimcolor(self):
        return self.DimColor

    @DimColor.setter
    def DimColor(self, value):
        self.animationsettings.DimColor = value

    # Lower case alias for DimColor setter
    @dimcolor.setter
    def dimcolor(self, value):
        self.DimColor = value

    @property
    def EntryEffect(self):
        return self.animationsettings.EntryEffect

    # Lower case alias for EntryEffect
    @property
    def entryeffect(self):
        return self.EntryEffect

    @EntryEffect.setter
    def EntryEffect(self, value):
        self.animationsettings.EntryEffect = value

    # Lower case alias for EntryEffect setter
    @entryeffect.setter
    def entryeffect(self, value):
        self.EntryEffect = value

    @property
    def Parent(self):
        return self.animationsettings.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PlaySettings(self):
        return PlaySettings(self.animationsettings.PlaySettings)

    # Lower case alias for PlaySettings
    @property
    def playsettings(self):
        return self.PlaySettings

    @property
    def SoundEffect(self):
        return SoundEffect(self.animationsettings.SoundEffect)

    # Lower case alias for SoundEffect
    @property
    def soundeffect(self):
        return self.SoundEffect

    @property
    def TextLevelEffect(self):
        return self.animationsettings.TextLevelEffect

    # Lower case alias for TextLevelEffect
    @property
    def textleveleffect(self):
        return self.TextLevelEffect

    @TextLevelEffect.setter
    def TextLevelEffect(self, value):
        self.animationsettings.TextLevelEffect = value

    # Lower case alias for TextLevelEffect setter
    @textleveleffect.setter
    def textleveleffect(self, value):
        self.TextLevelEffect = value

    @property
    def TextUnitEffect(self):
        return self.animationsettings.TextUnitEffect

    # Lower case alias for TextUnitEffect
    @property
    def textuniteffect(self):
        return self.TextUnitEffect

    @TextUnitEffect.setter
    def TextUnitEffect(self, value):
        self.animationsettings.TextUnitEffect = value

    # Lower case alias for TextUnitEffect setter
    @textuniteffect.setter
    def textuniteffect(self, value):
        self.TextUnitEffect = value


class Application:

    def __init__(self, application=None):
        self.application = application

    def new(self):
        self.application = win32com.client.Dispatch("PowerPoint.Application")
        return self

    @property
    def Active(self):
        return self.application.Active

    # Lower case alias for Active
    @property
    def active(self):
        return self.Active

    @property
    def ActiveEncryptionSession(self):
        return self.application.ActiveEncryptionSession

    # Lower case alias for ActiveEncryptionSession
    @property
    def activeencryptionsession(self):
        return self.ActiveEncryptionSession

    @property
    def ActivePresentation(self):
        return Presentation(self.application.ActivePresentation)

    # Lower case alias for ActivePresentation
    @property
    def activepresentation(self):
        return self.ActivePresentation

    @property
    def ActivePrinter(self):
        return self.application.ActivePrinter

    # Lower case alias for ActivePrinter
    @property
    def activeprinter(self):
        return self.ActivePrinter

    @property
    def ActiveProtectedViewWindow(self):
        return ProtectedViewWindow(self.application.ActiveProtectedViewWindow)

    # Lower case alias for ActiveProtectedViewWindow
    @property
    def activeprotectedviewwindow(self):
        return self.ActiveProtectedViewWindow

    @property
    def ActiveWindow(self):
        return DocumentWindow(self.application.ActiveWindow)

    # Lower case alias for ActiveWindow
    @property
    def activewindow(self):
        return self.ActiveWindow

    @property
    def AddIns(self):
        return AddIns(self.application.AddIns)

    # Lower case alias for AddIns
    @property
    def addins(self):
        return self.AddIns

    @property
    def Assistance(self):
        return self.application.Assistance

    # Lower case alias for Assistance
    @property
    def assistance(self):
        return self.Assistance

    @property
    def AutoCorrect(self):
        return AutoCorrect(self.application.AutoCorrect)

    # Lower case alias for AutoCorrect
    @property
    def autocorrect(self):
        return self.AutoCorrect

    @property
    def AutomationSecurity(self):
        return self.application.AutomationSecurity

    # Lower case alias for AutomationSecurity
    @property
    def automationsecurity(self):
        return self.AutomationSecurity

    @AutomationSecurity.setter
    def AutomationSecurity(self, value):
        self.application.AutomationSecurity = value

    # Lower case alias for AutomationSecurity setter
    @automationsecurity.setter
    def automationsecurity(self, value):
        self.AutomationSecurity = value

    @property
    def Build(self):
        return self.application.Build

    # Lower case alias for Build
    @property
    def build(self):
        return self.Build

    @property
    def Caption(self):
        return self.application.Caption

    # Lower case alias for Caption
    @property
    def caption(self):
        return self.Caption

    @Caption.setter
    def Caption(self, value):
        self.application.Caption = value

    # Lower case alias for Caption setter
    @caption.setter
    def caption(self, value):
        self.Caption = value

    @property
    def COMAddIns(self):
        return self.application.COMAddIns

    # Lower case alias for COMAddIns
    @property
    def comaddins(self):
        return self.COMAddIns

    @property
    def CommandBars(self):
        return self.application.CommandBars

    # Lower case alias for CommandBars
    @property
    def commandbars(self):
        return self.CommandBars

    @property
    def Creator(self):
        return self.application.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def DisplayAlerts(self):
        return self.application.DisplayAlerts

    # Lower case alias for DisplayAlerts
    @property
    def displayalerts(self):
        return self.DisplayAlerts

    @DisplayAlerts.setter
    def DisplayAlerts(self, value):
        self.application.DisplayAlerts = value

    # Lower case alias for DisplayAlerts setter
    @displayalerts.setter
    def displayalerts(self, value):
        self.DisplayAlerts = value

    @property
    def DisplayDocumentInformationPanel(self):
        return self.application.DisplayDocumentInformationPanel

    # Lower case alias for DisplayDocumentInformationPanel
    @property
    def displaydocumentinformationpanel(self):
        return self.DisplayDocumentInformationPanel

    @DisplayDocumentInformationPanel.setter
    def DisplayDocumentInformationPanel(self, value):
        self.application.DisplayDocumentInformationPanel = value

    # Lower case alias for DisplayDocumentInformationPanel setter
    @displaydocumentinformationpanel.setter
    def displaydocumentinformationpanel(self, value):
        self.DisplayDocumentInformationPanel = value

    @property
    def DisplayGridLines(self):
        return self.application.DisplayGridLines

    # Lower case alias for DisplayGridLines
    @property
    def displaygridlines(self):
        return self.DisplayGridLines

    @DisplayGridLines.setter
    def DisplayGridLines(self, value):
        self.application.DisplayGridLines = value

    # Lower case alias for DisplayGridLines setter
    @displaygridlines.setter
    def displaygridlines(self, value):
        self.DisplayGridLines = value

    @property
    def FeatureInstall(self):
        return self.application.FeatureInstall

    # Lower case alias for FeatureInstall
    @property
    def featureinstall(self):
        return self.FeatureInstall

    @FeatureInstall.setter
    def FeatureInstall(self, value):
        self.application.FeatureInstall = value

    # Lower case alias for FeatureInstall setter
    @featureinstall.setter
    def featureinstall(self, value):
        self.FeatureInstall = value

    def FileConverters(self, Index1=None, Index2=None):
        arguments = com_arguments([Index1, Index2])
        if callable(self.application.FileConverters):
            return self.application.FileConverters(*arguments)
        else:
            return self.application.GetFileConverters(*arguments)

    # Lower case alias for FileConverters
    def fileconverters(self, Index1=None, Index2=None):
        arguments = [Index1, Index2]
        return self.FileConverters(*arguments)

    def FileDialog(self, Type=None):
        arguments = com_arguments([Type])
        if callable(self.application.FileDialog):
            return self.application.FileDialog(*arguments)
        else:
            return self.application.GetFileDialog(*arguments)

    # Lower case alias for FileDialog
    def filedialog(self, Type=None):
        arguments = [Type]
        return self.FileDialog(*arguments)

    @property
    def FileValidation(self):
        return self.application.FileValidation

    # Lower case alias for FileValidation
    @property
    def filevalidation(self):
        return self.FileValidation

    @FileValidation.setter
    def FileValidation(self, value):
        self.application.FileValidation = value

    # Lower case alias for FileValidation setter
    @filevalidation.setter
    def filevalidation(self, value):
        self.FileValidation = value

    @property
    def Height(self):
        return self.application.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @Height.setter
    def Height(self, value):
        self.application.Height = value

    # Lower case alias for Height setter
    @height.setter
    def height(self, value):
        self.Height = value

    @property
    def IsSandboxed(self):
        return self.application.IsSandboxed

    # Lower case alias for IsSandboxed
    @property
    def issandboxed(self):
        return self.IsSandboxed

    @property
    def LanguageSettings(self):
        return self.application.LanguageSettings

    # Lower case alias for LanguageSettings
    @property
    def languagesettings(self):
        return self.LanguageSettings

    @property
    def Left(self):
        return self.application.Left

    # Lower case alias for Left
    @property
    def left(self):
        return self.Left

    @Left.setter
    def Left(self, value):
        self.application.Left = value

    # Lower case alias for Left setter
    @left.setter
    def left(self, value):
        self.Left = value

    @property
    def Name(self):
        return self.application.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def NewPresentation(self):
        return self.application.NewPresentation

    # Lower case alias for NewPresentation
    @property
    def newpresentation(self):
        return self.NewPresentation

    @property
    def OperatingSystem(self):
        return self.application.OperatingSystem

    # Lower case alias for OperatingSystem
    @property
    def operatingsystem(self):
        return self.OperatingSystem

    @property
    def Options(self):
        return Options(self.application.Options)

    # Lower case alias for Options
    @property
    def options(self):
        return self.Options

    @property
    def Path(self):
        return Application(self.application.Path)

    # Lower case alias for Path
    @property
    def path(self):
        return self.Path

    @property
    def Presentations(self):
        return Presentations(self.application.Presentations)

    # Lower case alias for Presentations
    @property
    def presentations(self):
        return self.Presentations

    @property
    def ProductCode(self):
        return self.application.ProductCode

    # Lower case alias for ProductCode
    @property
    def productcode(self):
        return self.ProductCode

    @property
    def ProtectedViewWindows(self):
        return ProtectedViewWindows(self.application.ProtectedViewWindows)

    # Lower case alias for ProtectedViewWindows
    @property
    def protectedviewwindows(self):
        return self.ProtectedViewWindows

    @property
    def SensitivityLabelPolicy(self):
        return self.application.SensitivityLabelPolicy

    # Lower case alias for SensitivityLabelPolicy
    @property
    def sensitivitylabelpolicy(self):
        return self.SensitivityLabelPolicy

    @property
    def ShowStartupDialog(self):
        return self.application.ShowStartupDialog

    # Lower case alias for ShowStartupDialog
    @property
    def showstartupdialog(self):
        return self.ShowStartupDialog

    @ShowStartupDialog.setter
    def ShowStartupDialog(self, value):
        self.application.ShowStartupDialog = value

    # Lower case alias for ShowStartupDialog setter
    @showstartupdialog.setter
    def showstartupdialog(self, value):
        self.ShowStartupDialog = value

    @property
    def ShowWindowsInTaskbar(self):
        return self.application.ShowWindowsInTaskbar

    # Lower case alias for ShowWindowsInTaskbar
    @property
    def showwindowsintaskbar(self):
        return self.ShowWindowsInTaskbar

    @ShowWindowsInTaskbar.setter
    def ShowWindowsInTaskbar(self, value):
        self.application.ShowWindowsInTaskbar = value

    # Lower case alias for ShowWindowsInTaskbar setter
    @showwindowsintaskbar.setter
    def showwindowsintaskbar(self, value):
        self.ShowWindowsInTaskbar = value

    @property
    def SlideShowWindows(self):
        return SlideShowWindows(self.application.SlideShowWindows)

    # Lower case alias for SlideShowWindows
    @property
    def slideshowwindows(self):
        return self.SlideShowWindows

    @property
    def SmartArtColors(self):
        return Application(self.application.SmartArtColors)

    # Lower case alias for SmartArtColors
    @property
    def smartartcolors(self):
        return self.SmartArtColors

    @property
    def SmartArtLayouts(self):
        return Application(self.application.SmartArtLayouts)

    # Lower case alias for SmartArtLayouts
    @property
    def smartartlayouts(self):
        return self.SmartArtLayouts

    @property
    def SmartArtQuickStyles(self):
        return Application(self.application.SmartArtQuickStyles)

    # Lower case alias for SmartArtQuickStyles
    @property
    def smartartquickstyles(self):
        return self.SmartArtQuickStyles

    @property
    def Top(self):
        return self.application.Top

    # Lower case alias for Top
    @property
    def top(self):
        return self.Top

    @Top.setter
    def Top(self, value):
        self.application.Top = value

    # Lower case alias for Top setter
    @top.setter
    def top(self, value):
        self.Top = value

    @property
    def VBE(self):
        return self.application.VBE

    # Lower case alias for VBE
    @property
    def vbe(self):
        return self.VBE

    @property
    def Version(self):
        return self.application.Version

    # Lower case alias for Version
    @property
    def version(self):
        return self.Version

    @property
    def Visible(self):
        return self.application.Visible

    # Lower case alias for Visible
    @property
    def visible(self):
        return self.Visible

    @Visible.setter
    def Visible(self, value):
        self.application.Visible = value

    # Lower case alias for Visible setter
    @visible.setter
    def visible(self, value):
        self.Visible = value

    @property
    def Width(self):
        return self.application.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

    @Width.setter
    def Width(self, value):
        self.application.Width = value

    # Lower case alias for Width setter
    @width.setter
    def width(self, value):
        self.Width = value

    @property
    def Windows(self):
        return DocumentWindows(self.application.Windows)

    # Lower case alias for Windows
    @property
    def windows(self):
        return self.Windows

    @property
    def WindowState(self):
        return self.application.WindowState

    # Lower case alias for WindowState
    @property
    def windowstate(self):
        return self.WindowState

    @WindowState.setter
    def WindowState(self, value):
        self.application.WindowState = value

    # Lower case alias for WindowState setter
    @windowstate.setter
    def windowstate(self, value):
        self.WindowState = value

    def Activate(self):
        self.application.Activate()

    def Help(self, HelpFile=None, ContextID=None):
        arguments = com_arguments([HelpFile, ContextID])
        self.application.Help(*arguments)

    def Quit(self):
        self.application.Quit()

    def Run(self, MacroName=None, safeArrayOfParams=None):
        arguments = com_arguments([MacroName, safeArrayOfParams])
        return self.application.Run(*arguments)

    def StartNewUndoEntry(self):
        self.application.StartNewUndoEntry()


class AutoCorrect:

    def __init__(self, autocorrect=None):
        self.autocorrect = autocorrect

    @property
    def DisplayAutoCorrectOptions(self):
        return self.autocorrect.DisplayAutoCorrectOptions

    # Lower case alias for DisplayAutoCorrectOptions
    @property
    def displayautocorrectoptions(self):
        return self.DisplayAutoCorrectOptions

    @DisplayAutoCorrectOptions.setter
    def DisplayAutoCorrectOptions(self, value):
        self.autocorrect.DisplayAutoCorrectOptions = value

    # Lower case alias for DisplayAutoCorrectOptions setter
    @displayautocorrectoptions.setter
    def displayautocorrectoptions(self, value):
        self.DisplayAutoCorrectOptions = value

    @property
    def DisplayAutoLayoutOptions(self):
        return self.autocorrect.DisplayAutoLayoutOptions

    # Lower case alias for DisplayAutoLayoutOptions
    @property
    def displayautolayoutoptions(self):
        return self.DisplayAutoLayoutOptions

    @DisplayAutoLayoutOptions.setter
    def DisplayAutoLayoutOptions(self, value):
        self.autocorrect.DisplayAutoLayoutOptions = value

    # Lower case alias for DisplayAutoLayoutOptions setter
    @displayautolayoutoptions.setter
    def displayautolayoutoptions(self, value):
        self.DisplayAutoLayoutOptions = value


class Axes:

    def __init__(self, axes=None):
        self.axes = axes

    @property
    def Application(self):
        return self.axes.Application

    @property
    def Count(self):
        return self.axes.Count

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.axes.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Parent(self):
        return self.axes.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

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

    # Lower case alias for AxisBetweenCategories
    @property
    def axisbetweencategories(self):
        return self.AxisBetweenCategories

    @AxisBetweenCategories.setter
    def AxisBetweenCategories(self, value):
        self.axis.AxisBetweenCategories = value

    # Lower case alias for AxisBetweenCategories setter
    @axisbetweencategories.setter
    def axisbetweencategories(self, value):
        self.AxisBetweenCategories = value

    @property
    def AxisGroup(self):
        return XlAxisGroup(self.axis.AxisGroup)

    # Lower case alias for AxisGroup
    @property
    def axisgroup(self):
        return self.AxisGroup

    @property
    def AxisTitle(self):
        return AxisTitle(self.axis.AxisTitle)

    # Lower case alias for AxisTitle
    @property
    def axistitle(self):
        return self.AxisTitle

    @property
    def BaseUnit(self):
        return XlTimeUnit(self.axis.BaseUnit)

    # Lower case alias for BaseUnit
    @property
    def baseunit(self):
        return self.BaseUnit

    @BaseUnit.setter
    def BaseUnit(self, value):
        self.axis.BaseUnit = value

    # Lower case alias for BaseUnit setter
    @baseunit.setter
    def baseunit(self, value):
        self.BaseUnit = value

    @property
    def BaseUnitIsAuto(self):
        return self.axis.BaseUnitIsAuto

    # Lower case alias for BaseUnitIsAuto
    @property
    def baseunitisauto(self):
        return self.BaseUnitIsAuto

    @BaseUnitIsAuto.setter
    def BaseUnitIsAuto(self, value):
        self.axis.BaseUnitIsAuto = value

    # Lower case alias for BaseUnitIsAuto setter
    @baseunitisauto.setter
    def baseunitisauto(self, value):
        self.BaseUnitIsAuto = value

    @property
    def Border(self):
        return ChartBorder(self.axis.Border)

    # Lower case alias for Border
    @property
    def border(self):
        return self.Border

    @property
    def CategoryNames(self):
        return self.axis.CategoryNames

    # Lower case alias for CategoryNames
    @property
    def categorynames(self):
        return self.CategoryNames

    @CategoryNames.setter
    def CategoryNames(self, value):
        self.axis.CategoryNames = value

    # Lower case alias for CategoryNames setter
    @categorynames.setter
    def categorynames(self, value):
        self.CategoryNames = value

    @property
    def CategoryType(self):
        return XlCategoryType(self.axis.CategoryType)

    # Lower case alias for CategoryType
    @property
    def categorytype(self):
        return self.CategoryType

    @CategoryType.setter
    def CategoryType(self, value):
        self.axis.CategoryType = value

    # Lower case alias for CategoryType setter
    @categorytype.setter
    def categorytype(self, value):
        self.CategoryType = value

    @property
    def Creator(self):
        return self.axis.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Crosses(self):
        return self.axis.Crosses

    # Lower case alias for Crosses
    @property
    def crosses(self):
        return self.Crosses

    @Crosses.setter
    def Crosses(self, value):
        self.axis.Crosses = value

    # Lower case alias for Crosses setter
    @crosses.setter
    def crosses(self, value):
        self.Crosses = value

    @property
    def CrossesAt(self):
        return self.axis.CrossesAt

    # Lower case alias for CrossesAt
    @property
    def crossesat(self):
        return self.CrossesAt

    @CrossesAt.setter
    def CrossesAt(self, value):
        self.axis.CrossesAt = value

    # Lower case alias for CrossesAt setter
    @crossesat.setter
    def crossesat(self, value):
        self.CrossesAt = value

    @property
    def DisplayUnit(self):
        return XlDisplayUnit(self.axis.DisplayUnit)

    # Lower case alias for DisplayUnit
    @property
    def displayunit(self):
        return self.DisplayUnit

    @DisplayUnit.setter
    def DisplayUnit(self, value):
        self.axis.DisplayUnit = value

    # Lower case alias for DisplayUnit setter
    @displayunit.setter
    def displayunit(self, value):
        self.DisplayUnit = value

    @property
    def DisplayUnitCustom(self):
        return self.axis.DisplayUnitCustom

    # Lower case alias for DisplayUnitCustom
    @property
    def displayunitcustom(self):
        return self.DisplayUnitCustom

    @DisplayUnitCustom.setter
    def DisplayUnitCustom(self, value):
        self.axis.DisplayUnitCustom = value

    # Lower case alias for DisplayUnitCustom setter
    @displayunitcustom.setter
    def displayunitcustom(self, value):
        self.DisplayUnitCustom = value

    @property
    def DisplayUnitLabel(self):
        return DisplayUnitLabel(self.axis.DisplayUnitLabel)

    # Lower case alias for DisplayUnitLabel
    @property
    def displayunitlabel(self):
        return self.DisplayUnitLabel

    @property
    def Format(self):
        return ChartFormat(self.axis.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def HasDisplayUnitLabel(self):
        return self.axis.HasDisplayUnitLabel

    # Lower case alias for HasDisplayUnitLabel
    @property
    def hasdisplayunitlabel(self):
        return self.HasDisplayUnitLabel

    @HasDisplayUnitLabel.setter
    def HasDisplayUnitLabel(self, value):
        self.axis.HasDisplayUnitLabel = value

    # Lower case alias for HasDisplayUnitLabel setter
    @hasdisplayunitlabel.setter
    def hasdisplayunitlabel(self, value):
        self.HasDisplayUnitLabel = value

    @property
    def HasMajorGridlines(self):
        return self.axis.HasMajorGridlines

    # Lower case alias for HasMajorGridlines
    @property
    def hasmajorgridlines(self):
        return self.HasMajorGridlines

    @HasMajorGridlines.setter
    def HasMajorGridlines(self, value):
        self.axis.HasMajorGridlines = value

    # Lower case alias for HasMajorGridlines setter
    @hasmajorgridlines.setter
    def hasmajorgridlines(self, value):
        self.HasMajorGridlines = value

    @property
    def HasMinorGridlines(self):
        return self.axis.HasMinorGridlines

    # Lower case alias for HasMinorGridlines
    @property
    def hasminorgridlines(self):
        return self.HasMinorGridlines

    @HasMinorGridlines.setter
    def HasMinorGridlines(self, value):
        self.axis.HasMinorGridlines = value

    # Lower case alias for HasMinorGridlines setter
    @hasminorgridlines.setter
    def hasminorgridlines(self, value):
        self.HasMinorGridlines = value

    @property
    def HasTitle(self):
        return self.axis.HasTitle

    # Lower case alias for HasTitle
    @property
    def hastitle(self):
        return self.HasTitle

    @HasTitle.setter
    def HasTitle(self, value):
        self.axis.HasTitle = value

    # Lower case alias for HasTitle setter
    @hastitle.setter
    def hastitle(self, value):
        self.HasTitle = value

    @property
    def Height(self):
        return self.axis.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @property
    def Left(self):
        return self.axis.Left

    # Lower case alias for Left
    @property
    def left(self):
        return self.Left

    @property
    def LogBase(self):
        return self.axis.LogBase

    # Lower case alias for LogBase
    @property
    def logbase(self):
        return self.LogBase

    @LogBase.setter
    def LogBase(self, value):
        self.axis.LogBase = value

    # Lower case alias for LogBase setter
    @logbase.setter
    def logbase(self, value):
        self.LogBase = value

    @property
    def MajorGridlines(self):
        return Gridlines(self.axis.MajorGridlines)

    # Lower case alias for MajorGridlines
    @property
    def majorgridlines(self):
        return self.MajorGridlines

    @property
    def MajorTickMark(self):
        return XlTickMark(self.axis.MajorTickMark)

    # Lower case alias for MajorTickMark
    @property
    def majortickmark(self):
        return self.MajorTickMark

    @MajorTickMark.setter
    def MajorTickMark(self, value):
        self.axis.MajorTickMark = value

    # Lower case alias for MajorTickMark setter
    @majortickmark.setter
    def majortickmark(self, value):
        self.MajorTickMark = value

    @property
    def MajorUnit(self):
        return self.axis.MajorUnit

    # Lower case alias for MajorUnit
    @property
    def majorunit(self):
        return self.MajorUnit

    @MajorUnit.setter
    def MajorUnit(self, value):
        self.axis.MajorUnit = value

    # Lower case alias for MajorUnit setter
    @majorunit.setter
    def majorunit(self, value):
        self.MajorUnit = value

    @property
    def MajorUnitIsAuto(self):
        return self.axis.MajorUnitIsAuto

    # Lower case alias for MajorUnitIsAuto
    @property
    def majorunitisauto(self):
        return self.MajorUnitIsAuto

    @MajorUnitIsAuto.setter
    def MajorUnitIsAuto(self, value):
        self.axis.MajorUnitIsAuto = value

    # Lower case alias for MajorUnitIsAuto setter
    @majorunitisauto.setter
    def majorunitisauto(self, value):
        self.MajorUnitIsAuto = value

    @property
    def MajorUnitScale(self):
        return self.axis.MajorUnitScale

    # Lower case alias for MajorUnitScale
    @property
    def majorunitscale(self):
        return self.MajorUnitScale

    @MajorUnitScale.setter
    def MajorUnitScale(self, value):
        self.axis.MajorUnitScale = value

    # Lower case alias for MajorUnitScale setter
    @majorunitscale.setter
    def majorunitscale(self, value):
        self.MajorUnitScale = value

    @property
    def MaximumScale(self):
        return self.axis.MaximumScale

    # Lower case alias for MaximumScale
    @property
    def maximumscale(self):
        return self.MaximumScale

    @MaximumScale.setter
    def MaximumScale(self, value):
        self.axis.MaximumScale = value

    # Lower case alias for MaximumScale setter
    @maximumscale.setter
    def maximumscale(self, value):
        self.MaximumScale = value

    @property
    def MaximumScaleIsAuto(self):
        return self.axis.MaximumScaleIsAuto

    # Lower case alias for MaximumScaleIsAuto
    @property
    def maximumscaleisauto(self):
        return self.MaximumScaleIsAuto

    @MaximumScaleIsAuto.setter
    def MaximumScaleIsAuto(self, value):
        self.axis.MaximumScaleIsAuto = value

    # Lower case alias for MaximumScaleIsAuto setter
    @maximumscaleisauto.setter
    def maximumscaleisauto(self, value):
        self.MaximumScaleIsAuto = value

    @property
    def MinimumScale(self):
        return self.axis.MinimumScale

    # Lower case alias for MinimumScale
    @property
    def minimumscale(self):
        return self.MinimumScale

    @MinimumScale.setter
    def MinimumScale(self, value):
        self.axis.MinimumScale = value

    # Lower case alias for MinimumScale setter
    @minimumscale.setter
    def minimumscale(self, value):
        self.MinimumScale = value

    @property
    def MinimumScaleIsAuto(self):
        return self.axis.MinimumScaleIsAuto

    # Lower case alias for MinimumScaleIsAuto
    @property
    def minimumscaleisauto(self):
        return self.MinimumScaleIsAuto

    @MinimumScaleIsAuto.setter
    def MinimumScaleIsAuto(self, value):
        self.axis.MinimumScaleIsAuto = value

    # Lower case alias for MinimumScaleIsAuto setter
    @minimumscaleisauto.setter
    def minimumscaleisauto(self, value):
        self.MinimumScaleIsAuto = value

    @property
    def MinorGridlines(self):
        return Gridlines(self.axis.MinorGridlines)

    # Lower case alias for MinorGridlines
    @property
    def minorgridlines(self):
        return self.MinorGridlines

    @property
    def MinorTickMark(self):
        return XlTickMark(self.axis.MinorTickMark)

    # Lower case alias for MinorTickMark
    @property
    def minortickmark(self):
        return self.MinorTickMark

    @MinorTickMark.setter
    def MinorTickMark(self, value):
        self.axis.MinorTickMark = value

    # Lower case alias for MinorTickMark setter
    @minortickmark.setter
    def minortickmark(self, value):
        self.MinorTickMark = value

    @property
    def MinorUnit(self):
        return self.axis.MinorUnit

    # Lower case alias for MinorUnit
    @property
    def minorunit(self):
        return self.MinorUnit

    @MinorUnit.setter
    def MinorUnit(self, value):
        self.axis.MinorUnit = value

    # Lower case alias for MinorUnit setter
    @minorunit.setter
    def minorunit(self, value):
        self.MinorUnit = value

    @property
    def MinorUnitIsAuto(self):
        return self.axis.MinorUnitIsAuto

    # Lower case alias for MinorUnitIsAuto
    @property
    def minorunitisauto(self):
        return self.MinorUnitIsAuto

    @MinorUnitIsAuto.setter
    def MinorUnitIsAuto(self, value):
        self.axis.MinorUnitIsAuto = value

    # Lower case alias for MinorUnitIsAuto setter
    @minorunitisauto.setter
    def minorunitisauto(self, value):
        self.MinorUnitIsAuto = value

    @property
    def MinorUnitScale(self):
        return self.axis.MinorUnitScale

    # Lower case alias for MinorUnitScale
    @property
    def minorunitscale(self):
        return self.MinorUnitScale

    @MinorUnitScale.setter
    def MinorUnitScale(self, value):
        self.axis.MinorUnitScale = value

    # Lower case alias for MinorUnitScale setter
    @minorunitscale.setter
    def minorunitscale(self, value):
        self.MinorUnitScale = value

    @property
    def Parent(self):
        return self.axis.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ReversePlotOrder(self):
        return self.axis.ReversePlotOrder

    # Lower case alias for ReversePlotOrder
    @property
    def reverseplotorder(self):
        return self.ReversePlotOrder

    @ReversePlotOrder.setter
    def ReversePlotOrder(self, value):
        self.axis.ReversePlotOrder = value

    # Lower case alias for ReversePlotOrder setter
    @reverseplotorder.setter
    def reverseplotorder(self, value):
        self.ReversePlotOrder = value

    @property
    def ScaleType(self):
        return XlScaleType(self.axis.ScaleType)

    # Lower case alias for ScaleType
    @property
    def scaletype(self):
        return self.ScaleType

    @ScaleType.setter
    def ScaleType(self, value):
        self.axis.ScaleType = value

    # Lower case alias for ScaleType setter
    @scaletype.setter
    def scaletype(self, value):
        self.ScaleType = value

    @property
    def TickLabelPosition(self):
        return self.axis.TickLabelPosition

    # Lower case alias for TickLabelPosition
    @property
    def ticklabelposition(self):
        return self.TickLabelPosition

    @TickLabelPosition.setter
    def TickLabelPosition(self, value):
        self.axis.TickLabelPosition = value

    # Lower case alias for TickLabelPosition setter
    @ticklabelposition.setter
    def ticklabelposition(self, value):
        self.TickLabelPosition = value

    @property
    def TickLabels(self):
        return TickLabels(self.axis.TickLabels)

    # Lower case alias for TickLabels
    @property
    def ticklabels(self):
        return self.TickLabels

    @property
    def TickLabelSpacing(self):
        return self.axis.TickLabelSpacing

    # Lower case alias for TickLabelSpacing
    @property
    def ticklabelspacing(self):
        return self.TickLabelSpacing

    @TickLabelSpacing.setter
    def TickLabelSpacing(self, value):
        self.axis.TickLabelSpacing = value

    # Lower case alias for TickLabelSpacing setter
    @ticklabelspacing.setter
    def ticklabelspacing(self, value):
        self.TickLabelSpacing = value

    @property
    def TickLabelSpacingIsAuto(self):
        return self.axis.TickLabelSpacingIsAuto

    # Lower case alias for TickLabelSpacingIsAuto
    @property
    def ticklabelspacingisauto(self):
        return self.TickLabelSpacingIsAuto

    @TickLabelSpacingIsAuto.setter
    def TickLabelSpacingIsAuto(self, value):
        self.axis.TickLabelSpacingIsAuto = value

    # Lower case alias for TickLabelSpacingIsAuto setter
    @ticklabelspacingisauto.setter
    def ticklabelspacingisauto(self, value):
        self.TickLabelSpacingIsAuto = value

    @property
    def TickMarkSpacing(self):
        return self.axis.TickMarkSpacing

    # Lower case alias for TickMarkSpacing
    @property
    def tickmarkspacing(self):
        return self.TickMarkSpacing

    @TickMarkSpacing.setter
    def TickMarkSpacing(self, value):
        self.axis.TickMarkSpacing = value

    # Lower case alias for TickMarkSpacing setter
    @tickmarkspacing.setter
    def tickmarkspacing(self, value):
        self.TickMarkSpacing = value

    @property
    def Top(self):
        return self.axis.Top

    # Lower case alias for Top
    @property
    def top(self):
        return self.Top

    @property
    def Type(self):
        return XlAxisType(self.axis.Type)

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @property
    def Width(self):
        return self.axis.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

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

    # Lower case alias for Caption
    @property
    def caption(self):
        return self.Caption

    @Caption.setter
    def Caption(self, value):
        self.axistitle.Caption = value

    # Lower case alias for Caption setter
    @caption.setter
    def caption(self, value):
        self.Caption = value

    def Characters(self, Start=None, Length=None):
        arguments = com_arguments([Start, Length])
        if callable(self.axistitle.Characters):
            return ChartCharacters(self.axistitle.Characters(*arguments))
        else:
            return ChartCharacters(self.axistitle.GetCharacters(*arguments))

    # Lower case alias for Characters
    def characters(self, Start=None, Length=None):
        arguments = [Start, Length]
        return self.Characters(*arguments)

    @property
    def Creator(self):
        return self.axistitle.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.axistitle.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def Formula(self):
        return self.axistitle.Formula

    # Lower case alias for Formula
    @property
    def formula(self):
        return self.Formula

    @Formula.setter
    def Formula(self, value):
        self.axistitle.Formula = value

    # Lower case alias for Formula setter
    @formula.setter
    def formula(self, value):
        self.Formula = value

    @property
    def FormulaLocal(self):
        return self.axistitle.FormulaLocal

    # Lower case alias for FormulaLocal
    @property
    def formulalocal(self):
        return self.FormulaLocal

    @FormulaLocal.setter
    def FormulaLocal(self, value):
        self.axistitle.FormulaLocal = value

    # Lower case alias for FormulaLocal setter
    @formulalocal.setter
    def formulalocal(self, value):
        self.FormulaLocal = value

    @property
    def FormulaR1C1(self):
        return self.axistitle.FormulaR1C1

    # Lower case alias for FormulaR1C1
    @property
    def formular1c1(self):
        return self.FormulaR1C1

    @FormulaR1C1.setter
    def FormulaR1C1(self, value):
        self.axistitle.FormulaR1C1 = value

    # Lower case alias for FormulaR1C1 setter
    @formular1c1.setter
    def formular1c1(self, value):
        self.FormulaR1C1 = value

    @property
    def FormulaR1C1Local(self):
        return self.axistitle.FormulaR1C1Local

    # Lower case alias for FormulaR1C1Local
    @property
    def formular1c1local(self):
        return self.FormulaR1C1Local

    @FormulaR1C1Local.setter
    def FormulaR1C1Local(self, value):
        self.axistitle.FormulaR1C1Local = value

    # Lower case alias for FormulaR1C1Local setter
    @formular1c1local.setter
    def formular1c1local(self, value):
        self.FormulaR1C1Local = value

    @property
    def Height(self):
        return self.axistitle.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @property
    def HorizontalAlignment(self):
        return self.axistitle.HorizontalAlignment

    # Lower case alias for HorizontalAlignment
    @property
    def horizontalalignment(self):
        return self.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.axistitle.HorizontalAlignment = value

    # Lower case alias for HorizontalAlignment setter
    @horizontalalignment.setter
    def horizontalalignment(self, value):
        self.HorizontalAlignment = value

    @property
    def IncludeInLayout(self):
        return self.axistitle.IncludeInLayout

    # Lower case alias for IncludeInLayout
    @property
    def includeinlayout(self):
        return self.IncludeInLayout

    @IncludeInLayout.setter
    def IncludeInLayout(self, value):
        self.axistitle.IncludeInLayout = value

    # Lower case alias for IncludeInLayout setter
    @includeinlayout.setter
    def includeinlayout(self, value):
        self.IncludeInLayout = value

    @property
    def Left(self):
        return self.axistitle.Left

    # Lower case alias for Left
    @property
    def left(self):
        return self.Left

    @Left.setter
    def Left(self, value):
        self.axistitle.Left = value

    # Lower case alias for Left setter
    @left.setter
    def left(self, value):
        self.Left = value

    @property
    def Name(self):
        return self.axistitle.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Orientation(self):
        return self.axistitle.Orientation

    # Lower case alias for Orientation
    @property
    def orientation(self):
        return self.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.axistitle.Orientation = value

    # Lower case alias for Orientation setter
    @orientation.setter
    def orientation(self, value):
        self.Orientation = value

    @property
    def Parent(self):
        return self.axistitle.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return XlChartElementPosition(self.axistitle.Position)

    # Lower case alias for Position
    @property
    def position(self):
        return self.Position

    @Position.setter
    def Position(self, value):
        self.axistitle.Position = value

    # Lower case alias for Position setter
    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def ReadingOrder(self):
        return XlReadingOrder(self.axistitle.ReadingOrder)

    # Lower case alias for ReadingOrder
    @property
    def readingorder(self):
        return self.ReadingOrder

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.axistitle.ReadingOrder = value

    # Lower case alias for ReadingOrder setter
    @readingorder.setter
    def readingorder(self, value):
        self.ReadingOrder = value

    @property
    def Shadow(self):
        return self.axistitle.Shadow

    # Lower case alias for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.axistitle.Shadow = value

    # Lower case alias for Shadow setter
    @shadow.setter
    def shadow(self, value):
        self.Shadow = value

    @property
    def Text(self):
        return self.axistitle.Text

    # Lower case alias for Text
    @property
    def text(self):
        return self.Text

    @Text.setter
    def Text(self, value):
        self.axistitle.Text = value

    # Lower case alias for Text setter
    @text.setter
    def text(self, value):
        self.Text = value

    @property
    def Top(self):
        return self.axistitle.Top

    # Lower case alias for Top
    @property
    def top(self):
        return self.Top

    @Top.setter
    def Top(self, value):
        self.axistitle.Top = value

    # Lower case alias for Top setter
    @top.setter
    def top(self, value):
        self.Top = value

    @property
    def VerticalAlignment(self):
        return self.axistitle.VerticalAlignment

    # Lower case alias for VerticalAlignment
    @property
    def verticalalignment(self):
        return self.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.axistitle.VerticalAlignment = value

    # Lower case alias for VerticalAlignment setter
    @verticalalignment.setter
    def verticalalignment(self, value):
        self.VerticalAlignment = value

    @property
    def Width(self):
        return self.axistitle.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.borders.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Item(self, BorderType=None):
        arguments = com_arguments([BorderType])
        return self.borders.Item(*arguments)


class Broadcast:

    def __init__(self, broadcast=None):
        self.broadcast = broadcast

    @property
    def Application(self):
        return Application(self.broadcast.Application)

    @property
    def AttendeeUrl(self):
        return self.broadcast.AttendeeUrl

    # Lower case alias for AttendeeUrl
    @property
    def attendeeurl(self):
        return self.AttendeeUrl

    @property
    def IsBroadcasting(self):
        return self.broadcast.IsBroadcasting

    # Lower case alias for IsBroadcasting
    @property
    def isbroadcasting(self):
        return self.IsBroadcasting

    @property
    def Parent(self):
        return self.broadcast.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def End(self):
        return self.broadcast.End()

    def Start(self, serverUrl=None):
        arguments = com_arguments([serverUrl])
        self.broadcast.Start(*arguments)


class BulletFormat:

    def __init__(self, bulletformat=None):
        self.bulletformat = bulletformat

    @property
    def Application(self):
        return Application(self.bulletformat.Application)

    @property
    def Character(self):
        return self.bulletformat.Character

    # Lower case alias for Character
    @property
    def character(self):
        return self.Character

    @Character.setter
    def Character(self, value):
        self.bulletformat.Character = value

    # Lower case alias for Character setter
    @character.setter
    def character(self, value):
        self.Character = value

    @property
    def Font(self):
        return Font(self.bulletformat.Font)

    # Lower case alias for Font
    @property
    def font(self):
        return self.Font

    @property
    def Number(self):
        return self.bulletformat.Number

    # Lower case alias for Number
    @property
    def number(self):
        return self.Number

    @property
    def Parent(self):
        return self.bulletformat.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RelativeSize(self):
        return self.bulletformat.RelativeSize

    # Lower case alias for RelativeSize
    @property
    def relativesize(self):
        return self.RelativeSize

    @RelativeSize.setter
    def RelativeSize(self, value):
        self.bulletformat.RelativeSize = value

    # Lower case alias for RelativeSize setter
    @relativesize.setter
    def relativesize(self, value):
        self.RelativeSize = value

    @property
    def StartValue(self):
        return self.bulletformat.StartValue

    # Lower case alias for StartValue
    @property
    def startvalue(self):
        return self.StartValue

    @StartValue.setter
    def StartValue(self, value):
        self.bulletformat.StartValue = value

    # Lower case alias for StartValue setter
    @startvalue.setter
    def startvalue(self, value):
        self.StartValue = value

    @property
    def Style(self):
        return self.bulletformat.Style

    # Lower case alias for Style
    @property
    def style(self):
        return self.Style

    @Style.setter
    def Style(self, value):
        self.bulletformat.Style = value

    # Lower case alias for Style setter
    @style.setter
    def style(self, value):
        self.Style = value

    @property
    def Type(self):
        return self.bulletformat.Type

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @Type.setter
    def Type(self, value):
        self.bulletformat.Type = value

    # Lower case alias for Type setter
    @type.setter
    def type(self, value):
        self.Type = value

    @property
    def UseTextColor(self):
        return self.bulletformat.UseTextColor

    # Lower case alias for UseTextColor
    @property
    def usetextcolor(self):
        return self.UseTextColor

    @UseTextColor.setter
    def UseTextColor(self, value):
        self.bulletformat.UseTextColor = value

    # Lower case alias for UseTextColor setter
    @usetextcolor.setter
    def usetextcolor(self, value):
        self.UseTextColor = value

    @property
    def UseTextFont(self):
        return self.bulletformat.UseTextFont

    # Lower case alias for UseTextFont
    @property
    def usetextfont(self):
        return self.UseTextFont

    @UseTextFont.setter
    def UseTextFont(self, value):
        self.bulletformat.UseTextFont = value

    # Lower case alias for UseTextFont setter
    @usetextfont.setter
    def usetextfont(self, value):
        self.UseTextFont = value

    def Picture(self):
        self.bulletformat.Picture()


class CalloutFormat:

    def __init__(self, calloutformat=None):
        self.calloutformat = calloutformat

    @property
    def Accent(self):
        return self.calloutformat.Accent

    # Lower case alias for Accent
    @property
    def accent(self):
        return self.Accent

    @Accent.setter
    def Accent(self, value):
        self.calloutformat.Accent = value

    # Lower case alias for Accent setter
    @accent.setter
    def accent(self, value):
        self.Accent = value

    @property
    def Angle(self):
        return self.calloutformat.Angle

    # Lower case alias for Angle
    @property
    def angle(self):
        return self.Angle

    @Angle.setter
    def Angle(self, value):
        self.calloutformat.Angle = value

    # Lower case alias for Angle setter
    @angle.setter
    def angle(self, value):
        self.Angle = value

    @property
    def Application(self):
        return Application(self.calloutformat.Application)

    @property
    def AutoAttach(self):
        return self.calloutformat.AutoAttach

    # Lower case alias for AutoAttach
    @property
    def autoattach(self):
        return self.AutoAttach

    @AutoAttach.setter
    def AutoAttach(self, value):
        self.calloutformat.AutoAttach = value

    # Lower case alias for AutoAttach setter
    @autoattach.setter
    def autoattach(self, value):
        self.AutoAttach = value

    @property
    def AutoLength(self):
        return self.calloutformat.AutoLength

    # Lower case alias for AutoLength
    @property
    def autolength(self):
        return self.AutoLength

    @property
    def Border(self):
        return self.calloutformat.Border

    # Lower case alias for Border
    @property
    def border(self):
        return self.Border

    @Border.setter
    def Border(self, value):
        self.calloutformat.Border = value

    # Lower case alias for Border setter
    @border.setter
    def border(self, value):
        self.Border = value

    @property
    def Creator(self):
        return self.calloutformat.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Drop(self):
        return self.calloutformat.Drop

    # Lower case alias for Drop
    @property
    def drop(self):
        return self.Drop

    @property
    def DropType(self):
        return self.calloutformat.DropType

    # Lower case alias for DropType
    @property
    def droptype(self):
        return self.DropType

    @property
    def Gap(self):
        return self.calloutformat.Gap

    # Lower case alias for Gap
    @property
    def gap(self):
        return self.Gap

    @Gap.setter
    def Gap(self, value):
        self.calloutformat.Gap = value

    # Lower case alias for Gap setter
    @gap.setter
    def gap(self, value):
        self.Gap = value

    @property
    def Length(self):
        return self.calloutformat.Length

    # Lower case alias for Length
    @property
    def length(self):
        return self.Length

    @property
    def Parent(self):
        return self.calloutformat.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Type(self):
        return self.calloutformat.Type

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @Type.setter
    def Type(self, value):
        self.calloutformat.Type = value

    # Lower case alias for Type setter
    @type.setter
    def type(self, value):
        self.Type = value

    def AutomaticLength(self):
        self.calloutformat.AutomaticLength()

    def CustomDrop(self, Drop=None):
        arguments = com_arguments([Drop])
        return self.calloutformat.CustomDrop(*arguments)

    def CustomLength(self, Length=None):
        arguments = com_arguments([Length])
        return self.calloutformat.CustomLength(*arguments)

    def PresetDrop(self, DropType=None):
        arguments = com_arguments([DropType])
        self.calloutformat.PresetDrop(*arguments)


class Cell:

    def __init__(self, cell=None):
        self.cell = cell

    @property
    def Application(self):
        return Application(self.cell.Application)

    @property
    def Borders(self):
        return Borders(self.cell.Borders)

    # Lower case alias for Borders
    @property
    def borders(self):
        return self.Borders

    @property
    def Parent(self):
        return self.cell.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Selected(self):
        return self.cell.Selected

    # Lower case alias for Selected
    @property
    def selected(self):
        return self.Selected

    @property
    def Shape(self):
        return Shape(self.cell.Shape)

    # Lower case alias for Shape
    @property
    def shape(self):
        return self.Shape

    def Merge(self, MergeTo=None):
        arguments = com_arguments([MergeTo])
        self.cell.Merge(*arguments)

    def Select(self):
        self.cell.Select()

    def Split(self, NumRows=None, NumColumns=None):
        arguments = com_arguments([NumRows, NumColumns])
        self.cell.Split(*arguments)


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

    # Lower case alias for Borders
    @property
    def borders(self):
        return self.Borders

    @property
    def Count(self):
        return self.cellrange.Count

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.cellrange.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.cellrange.Item(*arguments)


class Chart:

    def __init__(self, chart=None):
        self.chart = chart

    @property
    def AlternativeText(self):
        return self.chart.AlternativeText

    # Lower case alias for AlternativeText
    @property
    def alternativetext(self):
        return self.AlternativeText

    @AlternativeText.setter
    def AlternativeText(self, value):
        self.chart.AlternativeText = value

    # Lower case alias for AlternativeText setter
    @alternativetext.setter
    def alternativetext(self, value):
        self.AlternativeText = value

    @property
    def Application(self):
        return self.chart.Application

    @property
    def AutoScaling(self):
        return self.chart.AutoScaling

    # Lower case alias for AutoScaling
    @property
    def autoscaling(self):
        return self.AutoScaling

    @AutoScaling.setter
    def AutoScaling(self, value):
        self.chart.AutoScaling = value

    # Lower case alias for AutoScaling setter
    @autoscaling.setter
    def autoscaling(self, value):
        self.AutoScaling = value

    @property
    def BackWall(self):
        return Walls(self.chart.BackWall)

    # Lower case alias for BackWall
    @property
    def backwall(self):
        return self.BackWall

    @property
    def BarShape(self):
        return XlBarShape(self.chart.BarShape)

    # Lower case alias for BarShape
    @property
    def barshape(self):
        return self.BarShape

    @BarShape.setter
    def BarShape(self, value):
        self.chart.BarShape = value

    # Lower case alias for BarShape setter
    @barshape.setter
    def barshape(self, value):
        self.BarShape = value

    @property
    def ChartArea(self):
        return ChartArea(self.chart.ChartArea)

    # Lower case alias for ChartArea
    @property
    def chartarea(self):
        return self.ChartArea

    @property
    def ChartData(self):
        return ChartData(self.chart.ChartData)

    # Lower case alias for ChartData
    @property
    def chartdata(self):
        return self.ChartData

    @property
    def ChartStyle(self):
        return self.chart.ChartStyle

    # Lower case alias for ChartStyle
    @property
    def chartstyle(self):
        return self.ChartStyle

    @ChartStyle.setter
    def ChartStyle(self, value):
        self.chart.ChartStyle = value

    # Lower case alias for ChartStyle setter
    @chartstyle.setter
    def chartstyle(self, value):
        self.ChartStyle = value

    @property
    def ChartTitle(self):
        return ChartTitle(self.chart.ChartTitle)

    # Lower case alias for ChartTitle
    @property
    def charttitle(self):
        return self.ChartTitle

    @property
    def ChartType(self):
        return self.chart.ChartType

    # Lower case alias for ChartType
    @property
    def charttype(self):
        return self.ChartType

    @ChartType.setter
    def ChartType(self, value):
        self.chart.ChartType = value

    # Lower case alias for ChartType setter
    @charttype.setter
    def charttype(self, value):
        self.ChartType = value

    @property
    def Creator(self):
        return self.chart.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def DataTable(self):
        return DataTable(self.chart.DataTable)

    # Lower case alias for DataTable
    @property
    def datatable(self):
        return self.DataTable

    @property
    def DepthPercent(self):
        return self.chart.DepthPercent

    # Lower case alias for DepthPercent
    @property
    def depthpercent(self):
        return self.DepthPercent

    @DepthPercent.setter
    def DepthPercent(self, value):
        self.chart.DepthPercent = value

    # Lower case alias for DepthPercent setter
    @depthpercent.setter
    def depthpercent(self, value):
        self.DepthPercent = value

    @property
    def DisplayBlanksAs(self):
        return XlDisplayBlanksAs(self.chart.DisplayBlanksAs)

    # Lower case alias for DisplayBlanksAs
    @property
    def displayblanksas(self):
        return self.DisplayBlanksAs

    @DisplayBlanksAs.setter
    def DisplayBlanksAs(self, value):
        self.chart.DisplayBlanksAs = value

    # Lower case alias for DisplayBlanksAs setter
    @displayblanksas.setter
    def displayblanksas(self, value):
        self.DisplayBlanksAs = value

    @property
    def Elevation(self):
        return self.chart.Elevation

    # Lower case alias for Elevation
    @property
    def elevation(self):
        return self.Elevation

    @Elevation.setter
    def Elevation(self, value):
        self.chart.Elevation = value

    # Lower case alias for Elevation setter
    @elevation.setter
    def elevation(self, value):
        self.Elevation = value

    @property
    def Floor(self):
        return Floor(self.chart.Floor)

    # Lower case alias for Floor
    @property
    def floor(self):
        return self.Floor

    @property
    def Format(self):
        return self.chart.Format

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def GapDepth(self):
        return self.chart.GapDepth

    # Lower case alias for GapDepth
    @property
    def gapdepth(self):
        return self.GapDepth

    @GapDepth.setter
    def GapDepth(self, value):
        self.chart.GapDepth = value

    # Lower case alias for GapDepth setter
    @gapdepth.setter
    def gapdepth(self, value):
        self.GapDepth = value

    @property
    def HasAxis(self):
        return self.chart.HasAxis

    # Lower case alias for HasAxis
    @property
    def hasaxis(self):
        return self.HasAxis

    @HasAxis.setter
    def HasAxis(self, value):
        self.chart.HasAxis = value

    # Lower case alias for HasAxis setter
    @hasaxis.setter
    def hasaxis(self, value):
        self.HasAxis = value

    @property
    def HasDataTable(self):
        return self.chart.HasDataTable

    # Lower case alias for HasDataTable
    @property
    def hasdatatable(self):
        return self.HasDataTable

    @HasDataTable.setter
    def HasDataTable(self, value):
        self.chart.HasDataTable = value

    # Lower case alias for HasDataTable setter
    @hasdatatable.setter
    def hasdatatable(self, value):
        self.HasDataTable = value

    @property
    def HasLegend(self):
        return self.chart.HasLegend

    # Lower case alias for HasLegend
    @property
    def haslegend(self):
        return self.HasLegend

    @HasLegend.setter
    def HasLegend(self, value):
        self.chart.HasLegend = value

    # Lower case alias for HasLegend setter
    @haslegend.setter
    def haslegend(self, value):
        self.HasLegend = value

    @property
    def HasTitle(self):
        return self.chart.HasTitle

    # Lower case alias for HasTitle
    @property
    def hastitle(self):
        return self.HasTitle

    @HasTitle.setter
    def HasTitle(self, value):
        self.chart.HasTitle = value

    # Lower case alias for HasTitle setter
    @hastitle.setter
    def hastitle(self, value):
        self.HasTitle = value

    @property
    def HeightPercent(self):
        return self.chart.HeightPercent

    # Lower case alias for HeightPercent
    @property
    def heightpercent(self):
        return self.HeightPercent

    @HeightPercent.setter
    def HeightPercent(self, value):
        self.chart.HeightPercent = value

    # Lower case alias for HeightPercent setter
    @heightpercent.setter
    def heightpercent(self, value):
        self.HeightPercent = value

    @property
    def Legend(self):
        return Legend(self.chart.Legend)

    # Lower case alias for Legend
    @property
    def legend(self):
        return self.Legend

    @property
    def Name(self):
        return self.chart.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.chart.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.chart.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Perspective(self):
        return self.chart.Perspective

    # Lower case alias for Perspective
    @property
    def perspective(self):
        return self.Perspective

    @Perspective.setter
    def Perspective(self, value):
        self.chart.Perspective = value

    # Lower case alias for Perspective setter
    @perspective.setter
    def perspective(self, value):
        self.Perspective = value

    @property
    def PlotArea(self):
        return PlotArea(self.chart.PlotArea)

    # Lower case alias for PlotArea
    @property
    def plotarea(self):
        return self.PlotArea

    @property
    def PlotBy(self):
        return self.chart.PlotBy

    # Lower case alias for PlotBy
    @property
    def plotby(self):
        return self.PlotBy

    @PlotBy.setter
    def PlotBy(self, value):
        self.chart.PlotBy = value

    # Lower case alias for PlotBy setter
    @plotby.setter
    def plotby(self, value):
        self.PlotBy = value

    @property
    def PlotVisibleOnly(self):
        return self.chart.PlotVisibleOnly

    # Lower case alias for PlotVisibleOnly
    @property
    def plotvisibleonly(self):
        return self.PlotVisibleOnly

    @PlotVisibleOnly.setter
    def PlotVisibleOnly(self, value):
        self.chart.PlotVisibleOnly = value

    # Lower case alias for PlotVisibleOnly setter
    @plotvisibleonly.setter
    def plotvisibleonly(self, value):
        self.PlotVisibleOnly = value

    @property
    def RightAngleAxes(self):
        return self.chart.RightAngleAxes

    # Lower case alias for RightAngleAxes
    @property
    def rightangleaxes(self):
        return self.RightAngleAxes

    @RightAngleAxes.setter
    def RightAngleAxes(self, value):
        self.chart.RightAngleAxes = value

    # Lower case alias for RightAngleAxes setter
    @rightangleaxes.setter
    def rightangleaxes(self, value):
        self.RightAngleAxes = value

    @property
    def Rotation(self):
        return self.chart.Rotation

    # Lower case alias for Rotation
    @property
    def rotation(self):
        return self.Rotation

    @Rotation.setter
    def Rotation(self, value):
        self.chart.Rotation = value

    # Lower case alias for Rotation setter
    @rotation.setter
    def rotation(self, value):
        self.Rotation = value

    @property
    def Shapes(self):
        return Shapes(self.chart.Shapes)

    # Lower case alias for Shapes
    @property
    def shapes(self):
        return self.Shapes

    @property
    def ShowAllFieldButtons(self):
        return self.chart.ShowAllFieldButtons

    # Lower case alias for ShowAllFieldButtons
    @property
    def showallfieldbuttons(self):
        return self.ShowAllFieldButtons

    @ShowAllFieldButtons.setter
    def ShowAllFieldButtons(self, value):
        self.chart.ShowAllFieldButtons = value

    # Lower case alias for ShowAllFieldButtons setter
    @showallfieldbuttons.setter
    def showallfieldbuttons(self, value):
        self.ShowAllFieldButtons = value

    @property
    def ShowAxisFieldButtons(self):
        return self.chart.ShowAxisFieldButtons

    # Lower case alias for ShowAxisFieldButtons
    @property
    def showaxisfieldbuttons(self):
        return self.ShowAxisFieldButtons

    @ShowAxisFieldButtons.setter
    def ShowAxisFieldButtons(self, value):
        self.chart.ShowAxisFieldButtons = value

    # Lower case alias for ShowAxisFieldButtons setter
    @showaxisfieldbuttons.setter
    def showaxisfieldbuttons(self, value):
        self.ShowAxisFieldButtons = value

    @property
    def ShowDataLabelsOverMaximum(self):
        return self.chart.ShowDataLabelsOverMaximum

    # Lower case alias for ShowDataLabelsOverMaximum
    @property
    def showdatalabelsovermaximum(self):
        return self.ShowDataLabelsOverMaximum

    @ShowDataLabelsOverMaximum.setter
    def ShowDataLabelsOverMaximum(self, value):
        self.chart.ShowDataLabelsOverMaximum = value

    # Lower case alias for ShowDataLabelsOverMaximum setter
    @showdatalabelsovermaximum.setter
    def showdatalabelsovermaximum(self, value):
        self.ShowDataLabelsOverMaximum = value

    @property
    def ShowLegendFieldButtons(self):
        return self.chart.ShowLegendFieldButtons

    # Lower case alias for ShowLegendFieldButtons
    @property
    def showlegendfieldbuttons(self):
        return self.ShowLegendFieldButtons

    @ShowLegendFieldButtons.setter
    def ShowLegendFieldButtons(self, value):
        self.chart.ShowLegendFieldButtons = value

    # Lower case alias for ShowLegendFieldButtons setter
    @showlegendfieldbuttons.setter
    def showlegendfieldbuttons(self, value):
        self.ShowLegendFieldButtons = value

    @property
    def ShowReportFilterFieldButtons(self):
        return self.chart.ShowReportFilterFieldButtons

    # Lower case alias for ShowReportFilterFieldButtons
    @property
    def showreportfilterfieldbuttons(self):
        return self.ShowReportFilterFieldButtons

    @ShowReportFilterFieldButtons.setter
    def ShowReportFilterFieldButtons(self, value):
        self.chart.ShowReportFilterFieldButtons = value

    # Lower case alias for ShowReportFilterFieldButtons setter
    @showreportfilterfieldbuttons.setter
    def showreportfilterfieldbuttons(self, value):
        self.ShowReportFilterFieldButtons = value

    @property
    def ShowValueFieldButtons(self):
        return self.chart.ShowValueFieldButtons

    # Lower case alias for ShowValueFieldButtons
    @property
    def showvaluefieldbuttons(self):
        return self.ShowValueFieldButtons

    @ShowValueFieldButtons.setter
    def ShowValueFieldButtons(self, value):
        self.chart.ShowValueFieldButtons = value

    # Lower case alias for ShowValueFieldButtons setter
    @showvaluefieldbuttons.setter
    def showvaluefieldbuttons(self, value):
        self.ShowValueFieldButtons = value

    @property
    def SideWall(self):
        return Walls(self.chart.SideWall)

    # Lower case alias for SideWall
    @property
    def sidewall(self):
        return self.SideWall

    @property
    def Title(self):
        return self.chart.Title

    # Lower case alias for Title
    @property
    def title(self):
        return self.Title

    @Title.setter
    def Title(self, value):
        self.chart.Title = value

    # Lower case alias for Title setter
    @title.setter
    def title(self, value):
        self.Title = value

    @property
    def Walls(self):
        return Walls(self.chart.Walls)

    # Lower case alias for Walls
    @property
    def walls(self):
        return self.Walls

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

    def ChartGroups(self, Index=None):
        arguments = com_arguments([Index])
        self.chart.ChartGroups(*arguments)

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
        self.chart.Export(*arguments)

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

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.chartarea.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def Height(self):
        return self.chartarea.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @Height.setter
    def Height(self, value):
        self.chartarea.Height = value

    # Lower case alias for Height setter
    @height.setter
    def height(self, value):
        self.Height = value

    @property
    def Left(self):
        return self.chartarea.Left

    # Lower case alias for Left
    @property
    def left(self):
        return self.Left

    @Left.setter
    def Left(self, value):
        self.chartarea.Left = value

    # Lower case alias for Left setter
    @left.setter
    def left(self, value):
        self.Left = value

    @property
    def Name(self):
        return self.chartarea.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.chartarea.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Shadow(self):
        return self.chartarea.Shadow

    # Lower case alias for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.chartarea.Shadow = value

    # Lower case alias for Shadow setter
    @shadow.setter
    def shadow(self, value):
        self.Shadow = value

    @property
    def Top(self):
        return self.chartarea.Top

    # Lower case alias for Top
    @property
    def top(self):
        return self.Top

    @Top.setter
    def Top(self, value):
        self.chartarea.Top = value

    # Lower case alias for Top setter
    @top.setter
    def top(self, value):
        self.Top = value

    @property
    def Width(self):
        return self.chartarea.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

    @Width.setter
    def Width(self, value):
        self.chartarea.Width = value

    # Lower case alias for Width setter
    @width.setter
    def width(self, value):
        self.Width = value

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

    # Lower case alias for Color
    @property
    def color(self):
        return self.Color

    @Color.setter
    def Color(self, value):
        self.chartborder.Color = value

    # Lower case alias for Color setter
    @color.setter
    def color(self, value):
        self.Color = value

    @property
    def ColorIndex(self):
        return self.chartborder.ColorIndex

    # Lower case alias for ColorIndex
    @property
    def colorindex(self):
        return self.ColorIndex

    @ColorIndex.setter
    def ColorIndex(self, value):
        self.chartborder.ColorIndex = value

    # Lower case alias for ColorIndex setter
    @colorindex.setter
    def colorindex(self, value):
        self.ColorIndex = value

    @property
    def Creator(self):
        return self.chartborder.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def LineStyle(self):
        return XlLineStyle(self.chartborder.LineStyle)

    # Lower case alias for LineStyle
    @property
    def linestyle(self):
        return self.LineStyle

    @LineStyle.setter
    def LineStyle(self, value):
        self.chartborder.LineStyle = value

    # Lower case alias for LineStyle setter
    @linestyle.setter
    def linestyle(self, value):
        self.LineStyle = value

    @property
    def Parent(self):
        return self.chartborder.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Weight(self):
        return XlBorderWeight(self.chartborder.Weight)

    # Lower case alias for Weight
    @property
    def weight(self):
        return self.Weight

    @Weight.setter
    def Weight(self, value):
        self.chartborder.Weight = value

    # Lower case alias for Weight setter
    @weight.setter
    def weight(self, value):
        self.Weight = value


class ChartCharacters:

    def __init__(self, chartcharacters=None):
        self.chartcharacters = chartcharacters

    @property
    def Application(self):
        return self.chartcharacters.Application

    @property
    def Caption(self):
        return self.chartcharacters.Caption

    # Lower case alias for Caption
    @property
    def caption(self):
        return self.Caption

    @property
    def Count(self):
        return self.chartcharacters.Count

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.chartcharacters.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Font(self):
        return ChartFont(self.chartcharacters.Font)

    # Lower case alias for Font
    @property
    def font(self):
        return self.Font

    @property
    def Parent(self):
        return self.chartcharacters.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PhoneticCharacters(self):
        return self.chartcharacters.PhoneticCharacters

    # Lower case alias for PhoneticCharacters
    @property
    def phoneticcharacters(self):
        return self.PhoneticCharacters

    @PhoneticCharacters.setter
    def PhoneticCharacters(self, value):
        self.chartcharacters.PhoneticCharacters = value

    # Lower case alias for PhoneticCharacters setter
    @phoneticcharacters.setter
    def phoneticcharacters(self, value):
        self.PhoneticCharacters = value

    @property
    def Text(self):
        return self.chartcharacters.Text

    # Lower case alias for Text
    @property
    def text(self):
        return self.Text

    @Text.setter
    def Text(self, value):
        self.chartcharacters.Text = value

    # Lower case alias for Text setter
    @text.setter
    def text(self, value):
        self.Text = value

    def Delete(self):
        self.chartcharacters.Delete()

    def Insert(self, String=None):
        arguments = com_arguments([String])
        self.chartcharacters.Insert(*arguments)


class ChartData:

    def __init__(self, chartdata=None):
        self.chartdata = chartdata

    @property
    def IsLinked(self):
        return self.chartdata.IsLinked

    # Lower case alias for IsLinked
    @property
    def islinked(self):
        return self.IsLinked

    @property
    def Workbook(self):
        return self.chartdata.Workbook

    # Lower case alias for Workbook
    @property
    def workbook(self):
        return self.Workbook

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

    # Lower case alias for Background
    @property
    def background(self):
        return self.Background

    @Background.setter
    def Background(self, value):
        self.chartfont.Background = value

    # Lower case alias for Background setter
    @background.setter
    def background(self, value):
        self.Background = value

    @property
    def Bold(self):
        return self.chartfont.Bold

    # Lower case alias for Bold
    @property
    def bold(self):
        return self.Bold

    @Bold.setter
    def Bold(self, value):
        self.chartfont.Bold = value

    # Lower case alias for Bold setter
    @bold.setter
    def bold(self, value):
        self.Bold = value

    @property
    def Color(self):
        return self.chartfont.Color

    # Lower case alias for Color
    @property
    def color(self):
        return self.Color

    @Color.setter
    def Color(self, value):
        self.chartfont.Color = value

    # Lower case alias for Color setter
    @color.setter
    def color(self, value):
        self.Color = value

    @property
    def ColorIndex(self):
        return self.chartfont.ColorIndex

    # Lower case alias for ColorIndex
    @property
    def colorindex(self):
        return self.ColorIndex

    @ColorIndex.setter
    def ColorIndex(self, value):
        self.chartfont.ColorIndex = value

    # Lower case alias for ColorIndex setter
    @colorindex.setter
    def colorindex(self, value):
        self.ColorIndex = value

    @property
    def Creator(self):
        return self.chartfont.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def FontStyle(self):
        return self.chartfont.FontStyle

    # Lower case alias for FontStyle
    @property
    def fontstyle(self):
        return self.FontStyle

    @FontStyle.setter
    def FontStyle(self, value):
        self.chartfont.FontStyle = value

    # Lower case alias for FontStyle setter
    @fontstyle.setter
    def fontstyle(self, value):
        self.FontStyle = value

    @property
    def Italic(self):
        return self.chartfont.Italic

    # Lower case alias for Italic
    @property
    def italic(self):
        return self.Italic

    @Italic.setter
    def Italic(self, value):
        self.chartfont.Italic = value

    # Lower case alias for Italic setter
    @italic.setter
    def italic(self, value):
        self.Italic = value

    @property
    def Name(self):
        return self.chartfont.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.chartfont.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.chartfont.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Size(self):
        return self.chartfont.Size

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @Size.setter
    def Size(self, value):
        self.chartfont.Size = value

    # Lower case alias for Size setter
    @size.setter
    def size(self, value):
        self.Size = value

    @property
    def StrikeThrough(self):
        return self.chartfont.StrikeThrough

    # Lower case alias for StrikeThrough
    @property
    def strikethrough(self):
        return self.StrikeThrough

    @StrikeThrough.setter
    def StrikeThrough(self, value):
        self.chartfont.StrikeThrough = value

    # Lower case alias for StrikeThrough setter
    @strikethrough.setter
    def strikethrough(self, value):
        self.StrikeThrough = value

    @property
    def Subscript(self):
        return self.chartfont.Subscript

    # Lower case alias for Subscript
    @property
    def subscript(self):
        return self.Subscript

    @Subscript.setter
    def Subscript(self, value):
        self.chartfont.Subscript = value

    # Lower case alias for Subscript setter
    @subscript.setter
    def subscript(self, value):
        self.Subscript = value

    @property
    def Underline(self):
        return XlUnderlineStyle(self.chartfont.Underline)

    # Lower case alias for Underline
    @property
    def underline(self):
        return self.Underline

    @Underline.setter
    def Underline(self, value):
        self.chartfont.Underline = value

    # Lower case alias for Underline setter
    @underline.setter
    def underline(self, value):
        self.Underline = value


class ChartFormat:

    def __init__(self, chartformat=None):
        self.chartformat = chartformat

    @property
    def Application(self):
        return self.chartformat.Application

    @property
    def Creator(self):
        return self.chartformat.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Fill(self):
        return FillFormat(self.chartformat.Fill)

    # Lower case alias for Fill
    @property
    def fill(self):
        return self.Fill

    @property
    def Glow(self):
        return self.chartformat.Glow

    # Lower case alias for Glow
    @property
    def glow(self):
        return self.Glow

    @property
    def Line(self):
        return LineFormat(self.chartformat.Line)

    # Lower case alias for Line
    @property
    def line(self):
        return self.Line

    @property
    def Parent(self):
        return self.chartformat.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PictureFormat(self):
        return PictureFormat(self.chartformat.PictureFormat)

    # Lower case alias for PictureFormat
    @property
    def pictureformat(self):
        return self.PictureFormat

    @property
    def Shadow(self):
        return ShadowFormat(self.chartformat.Shadow)

    # Lower case alias for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @property
    def SoftEdge(self):
        return self.chartformat.SoftEdge

    # Lower case alias for SoftEdge
    @property
    def softedge(self):
        return self.SoftEdge

    @property
    def TextFrame2(self):
        return TextFrame2(self.chartformat.TextFrame2)

    # Lower case alias for TextFrame2
    @property
    def textframe2(self):
        return self.TextFrame2

    @property
    def ThreeD(self):
        return ThreeDFormat(self.chartformat.ThreeD)

    # Lower case alias for ThreeD
    @property
    def threed(self):
        return self.ThreeD


class ChartGroup:

    def __init__(self, chartgroup=None):
        self.chartgroup = chartgroup

    @property
    def Application(self):
        return self.chartgroup.Application

    @property
    def AxisGroup(self):
        return XlAxisGroup(self.chartgroup.AxisGroup)

    # Lower case alias for AxisGroup
    @property
    def axisgroup(self):
        return self.AxisGroup

    @AxisGroup.setter
    def AxisGroup(self, value):
        self.chartgroup.AxisGroup = value

    # Lower case alias for AxisGroup setter
    @axisgroup.setter
    def axisgroup(self, value):
        self.AxisGroup = value

    @property
    def BubbleScale(self):
        return self.chartgroup.BubbleScale

    # Lower case alias for BubbleScale
    @property
    def bubblescale(self):
        return self.BubbleScale

    @BubbleScale.setter
    def BubbleScale(self, value):
        self.chartgroup.BubbleScale = value

    # Lower case alias for BubbleScale setter
    @bubblescale.setter
    def bubblescale(self, value):
        self.BubbleScale = value

    @property
    def Creator(self):
        return self.chartgroup.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def DoughnutHoleSize(self):
        return self.chartgroup.DoughnutHoleSize

    # Lower case alias for DoughnutHoleSize
    @property
    def doughnutholesize(self):
        return self.DoughnutHoleSize

    @DoughnutHoleSize.setter
    def DoughnutHoleSize(self, value):
        self.chartgroup.DoughnutHoleSize = value

    # Lower case alias for DoughnutHoleSize setter
    @doughnutholesize.setter
    def doughnutholesize(self, value):
        self.DoughnutHoleSize = value

    @property
    def DownBars(self):
        return DownBars(self.chartgroup.DownBars)

    # Lower case alias for DownBars
    @property
    def downbars(self):
        return self.DownBars

    @property
    def DropLines(self):
        return DropLines(self.chartgroup.DropLines)

    # Lower case alias for DropLines
    @property
    def droplines(self):
        return self.DropLines

    @property
    def FirstSliceAngle(self):
        return self.chartgroup.FirstSliceAngle

    # Lower case alias for FirstSliceAngle
    @property
    def firstsliceangle(self):
        return self.FirstSliceAngle

    @FirstSliceAngle.setter
    def FirstSliceAngle(self, value):
        self.chartgroup.FirstSliceAngle = value

    # Lower case alias for FirstSliceAngle setter
    @firstsliceangle.setter
    def firstsliceangle(self, value):
        self.FirstSliceAngle = value

    @property
    def GapWidth(self):
        return self.chartgroup.GapWidth

    # Lower case alias for GapWidth
    @property
    def gapwidth(self):
        return self.GapWidth

    @GapWidth.setter
    def GapWidth(self, value):
        self.chartgroup.GapWidth = value

    # Lower case alias for GapWidth setter
    @gapwidth.setter
    def gapwidth(self, value):
        self.GapWidth = value

    @property
    def Has3DShading(self):
        return self.chartgroup.Has3DShading

    # Lower case alias for Has3DShading
    @property
    def has3dshading(self):
        return self.Has3DShading

    @Has3DShading.setter
    def Has3DShading(self, value):
        self.chartgroup.Has3DShading = value

    # Lower case alias for Has3DShading setter
    @has3dshading.setter
    def has3dshading(self, value):
        self.Has3DShading = value

    @property
    def HasDropLines(self):
        return self.chartgroup.HasDropLines

    # Lower case alias for HasDropLines
    @property
    def hasdroplines(self):
        return self.HasDropLines

    @HasDropLines.setter
    def HasDropLines(self, value):
        self.chartgroup.HasDropLines = value

    # Lower case alias for HasDropLines setter
    @hasdroplines.setter
    def hasdroplines(self, value):
        self.HasDropLines = value

    @property
    def HasHiLoLines(self):
        return self.chartgroup.HasHiLoLines

    # Lower case alias for HasHiLoLines
    @property
    def hashilolines(self):
        return self.HasHiLoLines

    @HasHiLoLines.setter
    def HasHiLoLines(self, value):
        self.chartgroup.HasHiLoLines = value

    # Lower case alias for HasHiLoLines setter
    @hashilolines.setter
    def hashilolines(self, value):
        self.HasHiLoLines = value

    @property
    def HasRadarAxisLabels(self):
        return self.chartgroup.HasRadarAxisLabels

    # Lower case alias for HasRadarAxisLabels
    @property
    def hasradaraxislabels(self):
        return self.HasRadarAxisLabels

    @HasRadarAxisLabels.setter
    def HasRadarAxisLabels(self, value):
        self.chartgroup.HasRadarAxisLabels = value

    # Lower case alias for HasRadarAxisLabels setter
    @hasradaraxislabels.setter
    def hasradaraxislabels(self, value):
        self.HasRadarAxisLabels = value

    @property
    def HasSeriesLines(self):
        return self.chartgroup.HasSeriesLines

    # Lower case alias for HasSeriesLines
    @property
    def hasserieslines(self):
        return self.HasSeriesLines

    @HasSeriesLines.setter
    def HasSeriesLines(self, value):
        self.chartgroup.HasSeriesLines = value

    # Lower case alias for HasSeriesLines setter
    @hasserieslines.setter
    def hasserieslines(self, value):
        self.HasSeriesLines = value

    @property
    def HasUpDownBars(self):
        return self.chartgroup.HasUpDownBars

    # Lower case alias for HasUpDownBars
    @property
    def hasupdownbars(self):
        return self.HasUpDownBars

    @HasUpDownBars.setter
    def HasUpDownBars(self, value):
        self.chartgroup.HasUpDownBars = value

    # Lower case alias for HasUpDownBars setter
    @hasupdownbars.setter
    def hasupdownbars(self, value):
        self.HasUpDownBars = value

    @property
    def HiLoLines(self):
        return HiLoLines(self.chartgroup.HiLoLines)

    # Lower case alias for HiLoLines
    @property
    def hilolines(self):
        return self.HiLoLines

    @property
    def Index(self):
        return self.chartgroup.Index

    # Lower case alias for Index
    @property
    def index(self):
        return self.Index

    @property
    def Overlap(self):
        return self.chartgroup.Overlap

    # Lower case alias for Overlap
    @property
    def overlap(self):
        return self.Overlap

    @Overlap.setter
    def Overlap(self, value):
        self.chartgroup.Overlap = value

    # Lower case alias for Overlap setter
    @overlap.setter
    def overlap(self, value):
        self.Overlap = value

    @property
    def Parent(self):
        return self.chartgroup.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RadarAxisLabels(self):
        return TickLabels(self.chartgroup.RadarAxisLabels)

    # Lower case alias for RadarAxisLabels
    @property
    def radaraxislabels(self):
        return self.RadarAxisLabels

    @property
    def SecondPlotSize(self):
        return self.chartgroup.SecondPlotSize

    # Lower case alias for SecondPlotSize
    @property
    def secondplotsize(self):
        return self.SecondPlotSize

    @SecondPlotSize.setter
    def SecondPlotSize(self, value):
        self.chartgroup.SecondPlotSize = value

    # Lower case alias for SecondPlotSize setter
    @secondplotsize.setter
    def secondplotsize(self, value):
        self.SecondPlotSize = value

    @property
    def SeriesLines(self):
        return SeriesLines(self.chartgroup.SeriesLines)

    # Lower case alias for SeriesLines
    @property
    def serieslines(self):
        return self.SeriesLines

    @property
    def ShowNegativeBubbles(self):
        return self.chartgroup.ShowNegativeBubbles

    # Lower case alias for ShowNegativeBubbles
    @property
    def shownegativebubbles(self):
        return self.ShowNegativeBubbles

    @ShowNegativeBubbles.setter
    def ShowNegativeBubbles(self, value):
        self.chartgroup.ShowNegativeBubbles = value

    # Lower case alias for ShowNegativeBubbles setter
    @shownegativebubbles.setter
    def shownegativebubbles(self, value):
        self.ShowNegativeBubbles = value

    @property
    def SizeRepresents(self):
        return self.chartgroup.SizeRepresents

    # Lower case alias for SizeRepresents
    @property
    def sizerepresents(self):
        return self.SizeRepresents

    @SizeRepresents.setter
    def SizeRepresents(self, value):
        self.chartgroup.SizeRepresents = value

    # Lower case alias for SizeRepresents setter
    @sizerepresents.setter
    def sizerepresents(self, value):
        self.SizeRepresents = value

    @property
    def SplitType(self):
        return XlChartSplitType(self.chartgroup.SplitType)

    # Lower case alias for SplitType
    @property
    def splittype(self):
        return self.SplitType

    @SplitType.setter
    def SplitType(self, value):
        self.chartgroup.SplitType = value

    # Lower case alias for SplitType setter
    @splittype.setter
    def splittype(self, value):
        self.SplitType = value

    @property
    def SplitValue(self):
        return self.chartgroup.SplitValue

    # Lower case alias for SplitValue
    @property
    def splitvalue(self):
        return self.SplitValue

    @SplitValue.setter
    def SplitValue(self, value):
        self.chartgroup.SplitValue = value

    # Lower case alias for SplitValue setter
    @splitvalue.setter
    def splitvalue(self, value):
        self.SplitValue = value

    @property
    def UpBars(self):
        return UpBars(self.chartgroup.UpBars)

    # Lower case alias for UpBars
    @property
    def upbars(self):
        return self.UpBars

    @property
    def VaryByCategories(self):
        return self.chartgroup.VaryByCategories

    # Lower case alias for VaryByCategories
    @property
    def varybycategories(self):
        return self.VaryByCategories

    @VaryByCategories.setter
    def VaryByCategories(self, value):
        self.chartgroup.VaryByCategories = value

    # Lower case alias for VaryByCategories setter
    @varybycategories.setter
    def varybycategories(self, value):
        self.VaryByCategories = value

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.chartgroups.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Parent(self):
        return self.chartgroups.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

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

    # Lower case alias for Caption
    @property
    def caption(self):
        return self.Caption

    @Caption.setter
    def Caption(self, value):
        self.charttitle.Caption = value

    # Lower case alias for Caption setter
    @caption.setter
    def caption(self, value):
        self.Caption = value

    def Characters(self, Start=None, Length=None):
        arguments = com_arguments([Start, Length])
        if callable(self.charttitle.Characters):
            return ChartCharacters(self.charttitle.Characters(*arguments))
        else:
            return ChartCharacters(self.charttitle.GetCharacters(*arguments))

    # Lower case alias for Characters
    def characters(self, Start=None, Length=None):
        arguments = [Start, Length]
        return self.Characters(*arguments)

    @property
    def Creator(self):
        return self.charttitle.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.charttitle.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def Formula(self):
        return self.charttitle.Formula

    # Lower case alias for Formula
    @property
    def formula(self):
        return self.Formula

    @Formula.setter
    def Formula(self, value):
        self.charttitle.Formula = value

    # Lower case alias for Formula setter
    @formula.setter
    def formula(self, value):
        self.Formula = value

    @property
    def FormulaLocal(self):
        return self.charttitle.FormulaLocal

    # Lower case alias for FormulaLocal
    @property
    def formulalocal(self):
        return self.FormulaLocal

    @FormulaLocal.setter
    def FormulaLocal(self, value):
        self.charttitle.FormulaLocal = value

    # Lower case alias for FormulaLocal setter
    @formulalocal.setter
    def formulalocal(self, value):
        self.FormulaLocal = value

    @property
    def FormulaR1C1(self):
        return self.charttitle.FormulaR1C1

    # Lower case alias for FormulaR1C1
    @property
    def formular1c1(self):
        return self.FormulaR1C1

    @FormulaR1C1.setter
    def FormulaR1C1(self, value):
        self.charttitle.FormulaR1C1 = value

    # Lower case alias for FormulaR1C1 setter
    @formular1c1.setter
    def formular1c1(self, value):
        self.FormulaR1C1 = value

    @property
    def FormulaR1C1Local(self):
        return self.charttitle.FormulaR1C1Local

    # Lower case alias for FormulaR1C1Local
    @property
    def formular1c1local(self):
        return self.FormulaR1C1Local

    @FormulaR1C1Local.setter
    def FormulaR1C1Local(self, value):
        self.charttitle.FormulaR1C1Local = value

    # Lower case alias for FormulaR1C1Local setter
    @formular1c1local.setter
    def formular1c1local(self, value):
        self.FormulaR1C1Local = value

    @property
    def Height(self):
        return self.charttitle.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @Height.setter
    def Height(self, value):
        self.charttitle.Height = value

    # Lower case alias for Height setter
    @height.setter
    def height(self, value):
        self.Height = value

    @property
    def HorizontalAlignment(self):
        return self.charttitle.HorizontalAlignment

    # Lower case alias for HorizontalAlignment
    @property
    def horizontalalignment(self):
        return self.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.charttitle.HorizontalAlignment = value

    # Lower case alias for HorizontalAlignment setter
    @horizontalalignment.setter
    def horizontalalignment(self, value):
        self.HorizontalAlignment = value

    @property
    def IncludeInLayout(self):
        return self.charttitle.IncludeInLayout

    # Lower case alias for IncludeInLayout
    @property
    def includeinlayout(self):
        return self.IncludeInLayout

    @IncludeInLayout.setter
    def IncludeInLayout(self, value):
        self.charttitle.IncludeInLayout = value

    # Lower case alias for IncludeInLayout setter
    @includeinlayout.setter
    def includeinlayout(self, value):
        self.IncludeInLayout = value

    @property
    def Left(self):
        return self.charttitle.Left

    # Lower case alias for Left
    @property
    def left(self):
        return self.Left

    @Left.setter
    def Left(self, value):
        self.charttitle.Left = value

    # Lower case alias for Left setter
    @left.setter
    def left(self, value):
        self.Left = value

    @property
    def Name(self):
        return self.charttitle.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Orientation(self):
        return self.charttitle.Orientation

    # Lower case alias for Orientation
    @property
    def orientation(self):
        return self.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.charttitle.Orientation = value

    # Lower case alias for Orientation setter
    @orientation.setter
    def orientation(self, value):
        self.Orientation = value

    @property
    def Parent(self):
        return self.charttitle.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return XlChartElementPosition(self.charttitle.Position)

    # Lower case alias for Position
    @property
    def position(self):
        return self.Position

    @Position.setter
    def Position(self, value):
        self.charttitle.Position = value

    # Lower case alias for Position setter
    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def ReadingOrder(self):
        return XlReadingOrder(self.charttitle.ReadingOrder)

    # Lower case alias for ReadingOrder
    @property
    def readingorder(self):
        return self.ReadingOrder

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.charttitle.ReadingOrder = value

    # Lower case alias for ReadingOrder setter
    @readingorder.setter
    def readingorder(self, value):
        self.ReadingOrder = value

    @property
    def Shadow(self):
        return self.charttitle.Shadow

    # Lower case alias for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.charttitle.Shadow = value

    # Lower case alias for Shadow setter
    @shadow.setter
    def shadow(self, value):
        self.Shadow = value

    @property
    def Text(self):
        return self.charttitle.Text

    # Lower case alias for Text
    @property
    def text(self):
        return self.Text

    @Text.setter
    def Text(self, value):
        self.charttitle.Text = value

    # Lower case alias for Text setter
    @text.setter
    def text(self, value):
        self.Text = value

    @property
    def Top(self):
        return self.charttitle.Top

    # Lower case alias for Top
    @property
    def top(self):
        return self.Top

    @Top.setter
    def Top(self, value):
        self.charttitle.Top = value

    # Lower case alias for Top setter
    @top.setter
    def top(self, value):
        self.Top = value

    @property
    def VerticalAlignment(self):
        return self.charttitle.VerticalAlignment

    # Lower case alias for VerticalAlignment
    @property
    def verticalalignment(self):
        return self.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.charttitle.VerticalAlignment = value

    # Lower case alias for VerticalAlignment setter
    @verticalalignment.setter
    def verticalalignment(self, value):
        self.VerticalAlignment = value

    @property
    def Width(self):
        return self.charttitle.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

    @Width.setter
    def Width(self, value):
        self.charttitle.Width = value

    # Lower case alias for Width setter
    @width.setter
    def width(self, value):
        self.Width = value

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

    # Lower case alias for FavorServerEditsDuringMerge
    @property
    def favorservereditsduringmerge(self):
        return self.FavorServerEditsDuringMerge

    @FavorServerEditsDuringMerge.setter
    def FavorServerEditsDuringMerge(self, value):
        self.coauthoring.FavorServerEditsDuringMerge = value

    # Lower case alias for FavorServerEditsDuringMerge setter
    @favorservereditsduringmerge.setter
    def favorservereditsduringmerge(self, value):
        self.FavorServerEditsDuringMerge = value

    @property
    def MergeMode(self):
        return self.coauthoring.MergeMode

    # Lower case alias for MergeMode
    @property
    def mergemode(self):
        return self.MergeMode

    @property
    def Parent(self):
        return self.coauthoring.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PendingUpdates(self):
        return self.coauthoring.PendingUpdates

    # Lower case alias for PendingUpdates
    @property
    def pendingupdates(self):
        return self.PendingUpdates

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

    # Lower case alias for By
    @property
    def by(self):
        return self.By

    @property
    def From(self):
        return self.coloreffect.From

    @property
    def Parent(self):
        return self.coloreffect.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def To(self):
        return self.coloreffect.To

    # Lower case alias for To
    @property
    def to(self):
        return self.To

    @To.setter
    def To(self, value):
        self.coloreffect.To = value

    # Lower case alias for To setter
    @to.setter
    def to(self, value):
        self.To = value


class ColorFormat:

    def __init__(self, colorformat=None):
        self.colorformat = colorformat

    @property
    def Application(self):
        return Application(self.colorformat.Application)

    @property
    def Brightness(self):
        return self.colorformat.Brightness

    # Lower case alias for Brightness
    @property
    def brightness(self):
        return self.Brightness

    @Brightness.setter
    def Brightness(self, value):
        self.colorformat.Brightness = value

    # Lower case alias for Brightness setter
    @brightness.setter
    def brightness(self, value):
        self.Brightness = value

    @property
    def Creator(self):
        return self.colorformat.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def ObjectThemeColor(self):
        return ColorFormat(self.colorformat.ObjectThemeColor)

    # Lower case alias for ObjectThemeColor
    @property
    def objectthemecolor(self):
        return self.ObjectThemeColor

    @ObjectThemeColor.setter
    def ObjectThemeColor(self, value):
        self.colorformat.ObjectThemeColor = value

    # Lower case alias for ObjectThemeColor setter
    @objectthemecolor.setter
    def objectthemecolor(self, value):
        self.ObjectThemeColor = value

    @property
    def Parent(self):
        return self.colorformat.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RGB(self):
        return self.colorformat.RGB

    # Lower case alias for RGB
    @property
    def rgb(self):
        return self.RGB

    @RGB.setter
    def RGB(self, value):
        self.colorformat.RGB = value

    # Lower case alias for RGB setter
    @rgb.setter
    def rgb(self, value):
        self.RGB = value

    @property
    def SchemeColor(self):
        return self.colorformat.SchemeColor

    # Lower case alias for SchemeColor
    @property
    def schemecolor(self):
        return self.SchemeColor

    @SchemeColor.setter
    def SchemeColor(self, value):
        self.colorformat.SchemeColor = value

    # Lower case alias for SchemeColor setter
    @schemecolor.setter
    def schemecolor(self, value):
        self.SchemeColor = value

    @property
    def TintAndShade(self):
        return self.colorformat.TintAndShade

    # Lower case alias for TintAndShade
    @property
    def tintandshade(self):
        return self.TintAndShade

    @TintAndShade.setter
    def TintAndShade(self, value):
        self.colorformat.TintAndShade = value

    # Lower case alias for TintAndShade setter
    @tintandshade.setter
    def tintandshade(self, value):
        self.TintAndShade = value

    @property
    def Type(self):
        return self.colorformat.Type

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type


class ColorScheme:

    def __init__(self, colorscheme=None):
        self.colorscheme = colorscheme

    @property
    def Application(self):
        return Application(self.colorscheme.Application)

    @property
    def Count(self):
        return self.colorscheme.Count

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.colorscheme.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Colors(self, SchemeColor=None):
        arguments = com_arguments([SchemeColor])
        return self.colorscheme.Colors(*arguments)

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.colorschemes.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Add(self, Scheme=None):
        arguments = com_arguments([Scheme])
        return ColorScheme(self.colorschemes.Add(*arguments))

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.colorschemes.Item(*arguments)


class Column:

    def __init__(self, column=None):
        self.column = column

    @property
    def Application(self):
        return Application(self.column.Application)

    def Cells(self, RowIndex=None, ColumnIndex=None):
        arguments = com_arguments([RowIndex, ColumnIndex])
        if callable(self.column.Cells):
            return CellRange(self.column.Cells(*arguments))
        else:
            return CellRange(self.column.GetCells(*arguments))

    # Lower case alias for Cells
    def cells(self, RowIndex=None, ColumnIndex=None):
        arguments = [RowIndex, ColumnIndex]
        return self.Cells(*arguments)

    @property
    def Parent(self):
        return self.column.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Width(self):
        return self.column.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

    @Width.setter
    def Width(self, value):
        self.column.Width = value

    # Lower case alias for Width setter
    @width.setter
    def width(self, value):
        self.Width = value

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.columns.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Add(self, BeforeColumn=None):
        arguments = com_arguments([BeforeColumn])
        return Column(self.columns.Add(*arguments))

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.columns.Item(*arguments)


class CommandEffect:

    def __init__(self, commandeffect=None):
        self.commandeffect = commandeffect

    @property
    def Application(self):
        return Application(self.commandeffect.Application)

    @property
    def Bookmark(self):
        return self.commandeffect.Bookmark

    # Lower case alias for Bookmark
    @property
    def bookmark(self):
        return self.Bookmark

    @Bookmark.setter
    def Bookmark(self, value):
        self.commandeffect.Bookmark = value

    # Lower case alias for Bookmark setter
    @bookmark.setter
    def bookmark(self, value):
        self.Bookmark = value

    @property
    def Command(self):
        return self.commandeffect.Command

    # Lower case alias for Command
    @property
    def command(self):
        return self.Command

    @Command.setter
    def Command(self, value):
        self.commandeffect.Command = value

    # Lower case alias for Command setter
    @command.setter
    def command(self, value):
        self.Command = value

    @property
    def Parent(self):
        return self.commandeffect.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Type(self):
        return self.commandeffect.Type

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @Type.setter
    def Type(self, value):
        self.commandeffect.Type = value

    # Lower case alias for Type setter
    @type.setter
    def type(self, value):
        self.Type = value


class Comment:

    def __init__(self, comment=None):
        self.comment = comment

    @property
    def Application(self):
        return Application(self.comment.Application)

    @property
    def Author(self):
        return Comment(self.comment.Author)

    # Lower case alias for Author
    @property
    def author(self):
        return self.Author

    @property
    def AuthorIndex(self):
        return self.comment.AuthorIndex

    # Lower case alias for AuthorIndex
    @property
    def authorindex(self):
        return self.AuthorIndex

    @property
    def AuthorInitials(self):
        return Comment(self.comment.AuthorInitials)

    # Lower case alias for AuthorInitials
    @property
    def authorinitials(self):
        return self.AuthorInitials

    @property
    def DateTime(self):
        return self.comment.DateTime

    # Lower case alias for DateTime
    @property
    def datetime(self):
        return self.DateTime

    @property
    def Left(self):
        return self.comment.Left

    # Lower case alias for Left
    @property
    def left(self):
        return self.Left

    @property
    def Parent(self):
        return self.comment.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Text(self):
        return self.comment.Text

    # Lower case alias for Text
    @property
    def text(self):
        return self.Text

    @property
    def Top(self):
        return self.comment.Top

    # Lower case alias for Top
    @property
    def top(self):
        return self.Top

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.comments.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Add(self, Left=None, Top=None, Author=None, AuthorInitials=None, Text=None):
        arguments = com_arguments([Left, Top, Author, AuthorInitials, Text])
        return self.comments.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.comments.Item(*arguments)


class ConnectorFormat:

    def __init__(self, connectorformat=None):
        self.connectorformat = connectorformat

    @property
    def Application(self):
        return Application(self.connectorformat.Application)

    @property
    def BeginConnected(self):
        return self.connectorformat.BeginConnected

    # Lower case alias for BeginConnected
    @property
    def beginconnected(self):
        return self.BeginConnected

    @BeginConnected.setter
    def BeginConnected(self, value):
        self.connectorformat.BeginConnected = value

    # Lower case alias for BeginConnected setter
    @beginconnected.setter
    def beginconnected(self, value):
        self.BeginConnected = value

    @property
    def BeginConnectedShape(self):
        return Shape(self.connectorformat.BeginConnectedShape)

    # Lower case alias for BeginConnectedShape
    @property
    def beginconnectedshape(self):
        return self.BeginConnectedShape

    @property
    def BeginConnectionSite(self):
        return self.connectorformat.BeginConnectionSite

    # Lower case alias for BeginConnectionSite
    @property
    def beginconnectionsite(self):
        return self.BeginConnectionSite

    @property
    def Creator(self):
        return self.connectorformat.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def EndConnected(self):
        return self.connectorformat.EndConnected

    # Lower case alias for EndConnected
    @property
    def endconnected(self):
        return self.EndConnected

    @property
    def EndConnectedShape(self):
        return Shape(self.connectorformat.EndConnectedShape)

    # Lower case alias for EndConnectedShape
    @property
    def endconnectedshape(self):
        return self.EndConnectedShape

    @property
    def EndConnectionSite(self):
        return self.connectorformat.EndConnectionSite

    # Lower case alias for EndConnectionSite
    @property
    def endconnectionsite(self):
        return self.EndConnectionSite

    @property
    def Parent(self):
        return self.connectorformat.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Type(self):
        return self.connectorformat.Type

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @Type.setter
    def Type(self, value):
        self.connectorformat.Type = value

    # Lower case alias for Type setter
    @type.setter
    def type(self, value):
        self.Type = value

    def BeginConnect(self, ConnectedShape=None, ConnectionSite=None):
        arguments = com_arguments([ConnectedShape, ConnectionSite])
        self.connectorformat.BeginConnect(*arguments)

    def BeginDisconnect(self):
        self.connectorformat.BeginDisconnect()

    def EndConnect(self, ConnectedShape=None, ConnectionSite=None):
        arguments = com_arguments([ConnectedShape, ConnectionSite])
        self.connectorformat.EndConnect(*arguments)

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return CustomerData(self.customerdata.Parent)

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Add(self):
        return self.customerdata.Add()

    def Delete(self, Id=None):
        arguments = com_arguments([Id])
        self.customerdata.Delete(*arguments)

    def Item(self, Id=None):
        arguments = com_arguments([Id])
        return self.customerdata.Item(*arguments)


class CustomLayout:

    def __init__(self, customlayout=None):
        self.customlayout = customlayout

    @property
    def Application(self):
        return Application(self.customlayout.Application)

    @property
    def Background(self):
        return ShapeRange(self.customlayout.Background)

    # Lower case alias for Background
    @property
    def background(self):
        return self.Background

    @property
    def CustomerData(self):
        return CustomerData(self.customlayout.CustomerData)

    # Lower case alias for CustomerData
    @property
    def customerdata(self):
        return self.CustomerData

    @property
    def Design(self):
        return Design(self.customlayout.Design)

    # Lower case alias for Design
    @property
    def design(self):
        return self.Design

    @property
    def DisplayMasterShapes(self):
        return self.customlayout.DisplayMasterShapes

    # Lower case alias for DisplayMasterShapes
    @property
    def displaymastershapes(self):
        return self.DisplayMasterShapes

    @DisplayMasterShapes.setter
    def DisplayMasterShapes(self, value):
        self.customlayout.DisplayMasterShapes = value

    # Lower case alias for DisplayMasterShapes setter
    @displaymastershapes.setter
    def displaymastershapes(self, value):
        self.DisplayMasterShapes = value

    @property
    def FollowMasterBackground(self):
        return self.customlayout.FollowMasterBackground

    # Lower case alias for FollowMasterBackground
    @property
    def followmasterbackground(self):
        return self.FollowMasterBackground

    @FollowMasterBackground.setter
    def FollowMasterBackground(self, value):
        self.customlayout.FollowMasterBackground = value

    # Lower case alias for FollowMasterBackground setter
    @followmasterbackground.setter
    def followmasterbackground(self, value):
        self.FollowMasterBackground = value

    @property
    def HeadersFooters(self):
        return HeadersFooters(self.customlayout.HeadersFooters)

    # Lower case alias for HeadersFooters
    @property
    def headersfooters(self):
        return self.HeadersFooters

    @property
    def Height(self):
        return self.customlayout.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @property
    def Hyperlinks(self):
        return Hyperlinks(self.customlayout.Hyperlinks)

    # Lower case alias for Hyperlinks
    @property
    def hyperlinks(self):
        return self.Hyperlinks

    @property
    def Index(self):
        return CustomLayouts(self.customlayout.Index)

    # Lower case alias for Index
    @property
    def index(self):
        return self.Index

    @property
    def MatchingName(self):
        return self.customlayout.MatchingName

    # Lower case alias for MatchingName
    @property
    def matchingname(self):
        return self.MatchingName

    @MatchingName.setter
    def MatchingName(self, value):
        self.customlayout.MatchingName = value

    # Lower case alias for MatchingName setter
    @matchingname.setter
    def matchingname(self, value):
        self.MatchingName = value

    @property
    def Name(self):
        return self.customlayout.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.customlayout.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return CustomLayout(self.customlayout.Parent)

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Preserved(self):
        return self.customlayout.Preserved

    # Lower case alias for Preserved
    @property
    def preserved(self):
        return self.Preserved

    @Preserved.setter
    def Preserved(self, value):
        self.customlayout.Preserved = value

    # Lower case alias for Preserved setter
    @preserved.setter
    def preserved(self, value):
        self.Preserved = value

    @property
    def Shapes(self):
        return Shapes(self.customlayout.Shapes)

    # Lower case alias for Shapes
    @property
    def shapes(self):
        return self.Shapes

    @property
    def SlideShowTransition(self):
        return SlideShowTransition(self.customlayout.SlideShowTransition)

    # Lower case alias for SlideShowTransition
    @property
    def slideshowtransition(self):
        return self.SlideShowTransition

    @property
    def ThemeColorScheme(self):
        return self.customlayout.ThemeColorScheme

    # Lower case alias for ThemeColorScheme
    @property
    def themecolorscheme(self):
        return self.ThemeColorScheme

    @property
    def TimeLine(self):
        return TimeLine(self.customlayout.TimeLine)

    # Lower case alias for TimeLine
    @property
    def timeline(self):
        return self.TimeLine

    @property
    def Width(self):
        return self.customlayout.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

    def Copy(self):
        self.customlayout.Copy()

    def Cut(self):
        self.customlayout.Cut()

    def Delete(self):
        self.customlayout.Delete()

    def Duplicate(self):
        return self.customlayout.Duplicate()

    def MoveTo(self, toPos=None):
        arguments = com_arguments([toPos])
        self.customlayout.MoveTo(*arguments)

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.customlayouts.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Add(self, Index=None):
        arguments = com_arguments([Index])
        return self.customlayouts.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.customlayouts.Item(*arguments)

    def Paste(self, Index=None):
        arguments = com_arguments([Index])
        return self.customlayouts.Paste(*arguments)


class DataLabel:

    def __init__(self, datalabel=None):
        self.datalabel = datalabel

    @property
    def Application(self):
        return self.datalabel.Application

    @property
    def AutoText(self):
        return self.datalabel.AutoText

    # Lower case alias for AutoText
    @property
    def autotext(self):
        return self.AutoText

    @AutoText.setter
    def AutoText(self, value):
        self.datalabel.AutoText = value

    # Lower case alias for AutoText setter
    @autotext.setter
    def autotext(self, value):
        self.AutoText = value

    @property
    def Caption(self):
        return self.datalabel.Caption

    # Lower case alias for Caption
    @property
    def caption(self):
        return self.Caption

    @Caption.setter
    def Caption(self, value):
        self.datalabel.Caption = value

    # Lower case alias for Caption setter
    @caption.setter
    def caption(self, value):
        self.Caption = value

    def Characters(self, Start=None, Length=None):
        arguments = com_arguments([Start, Length])
        if callable(self.datalabel.Characters):
            return ChartCharacters(self.datalabel.Characters(*arguments))
        else:
            return ChartCharacters(self.datalabel.GetCharacters(*arguments))

    # Lower case alias for Characters
    def characters(self, Start=None, Length=None):
        arguments = [Start, Length]
        return self.Characters(*arguments)

    @property
    def Creator(self):
        return self.datalabel.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.datalabel.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def Formula(self):
        return self.datalabel.Formula

    # Lower case alias for Formula
    @property
    def formula(self):
        return self.Formula

    @Formula.setter
    def Formula(self, value):
        self.datalabel.Formula = value

    # Lower case alias for Formula setter
    @formula.setter
    def formula(self, value):
        self.Formula = value

    @property
    def FormulaLocal(self):
        return self.datalabel.FormulaLocal

    # Lower case alias for FormulaLocal
    @property
    def formulalocal(self):
        return self.FormulaLocal

    @FormulaLocal.setter
    def FormulaLocal(self, value):
        self.datalabel.FormulaLocal = value

    # Lower case alias for FormulaLocal setter
    @formulalocal.setter
    def formulalocal(self, value):
        self.FormulaLocal = value

    @property
    def FormulaR1C1(self):
        return self.datalabel.FormulaR1C1

    # Lower case alias for FormulaR1C1
    @property
    def formular1c1(self):
        return self.FormulaR1C1

    @FormulaR1C1.setter
    def FormulaR1C1(self, value):
        self.datalabel.FormulaR1C1 = value

    # Lower case alias for FormulaR1C1 setter
    @formular1c1.setter
    def formular1c1(self, value):
        self.FormulaR1C1 = value

    @property
    def FormulaR1C1Local(self):
        return self.datalabel.FormulaR1C1Local

    # Lower case alias for FormulaR1C1Local
    @property
    def formular1c1local(self):
        return self.FormulaR1C1Local

    @FormulaR1C1Local.setter
    def FormulaR1C1Local(self, value):
        self.datalabel.FormulaR1C1Local = value

    # Lower case alias for FormulaR1C1Local setter
    @formular1c1local.setter
    def formular1c1local(self, value):
        self.FormulaR1C1Local = value

    @property
    def Height(self):
        return self.datalabel.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @property
    def HorizontalAlignment(self):
        return self.datalabel.HorizontalAlignment

    # Lower case alias for HorizontalAlignment
    @property
    def horizontalalignment(self):
        return self.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.datalabel.HorizontalAlignment = value

    # Lower case alias for HorizontalAlignment setter
    @horizontalalignment.setter
    def horizontalalignment(self, value):
        self.HorizontalAlignment = value

    @property
    def Left(self):
        return self.datalabel.Left

    # Lower case alias for Left
    @property
    def left(self):
        return self.Left

    @Left.setter
    def Left(self, value):
        self.datalabel.Left = value

    # Lower case alias for Left setter
    @left.setter
    def left(self, value):
        self.Left = value

    @property
    def Name(self):
        return self.datalabel.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def NumberFormat(self):
        return self.datalabel.NumberFormat

    # Lower case alias for NumberFormat
    @property
    def numberformat(self):
        return self.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value):
        self.datalabel.NumberFormat = value

    # Lower case alias for NumberFormat setter
    @numberformat.setter
    def numberformat(self, value):
        self.NumberFormat = value

    @property
    def NumberFormatLinked(self):
        return self.datalabel.NumberFormatLinked

    # Lower case alias for NumberFormatLinked
    @property
    def numberformatlinked(self):
        return self.NumberFormatLinked

    @NumberFormatLinked.setter
    def NumberFormatLinked(self, value):
        self.datalabel.NumberFormatLinked = value

    # Lower case alias for NumberFormatLinked setter
    @numberformatlinked.setter
    def numberformatlinked(self, value):
        self.NumberFormatLinked = value

    @property
    def NumberFormatLocal(self):
        return self.datalabel.NumberFormatLocal

    # Lower case alias for NumberFormatLocal
    @property
    def numberformatlocal(self):
        return self.NumberFormatLocal

    @NumberFormatLocal.setter
    def NumberFormatLocal(self, value):
        self.datalabel.NumberFormatLocal = value

    # Lower case alias for NumberFormatLocal setter
    @numberformatlocal.setter
    def numberformatlocal(self, value):
        self.NumberFormatLocal = value

    @property
    def Orientation(self):
        return self.datalabel.Orientation

    # Lower case alias for Orientation
    @property
    def orientation(self):
        return self.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.datalabel.Orientation = value

    # Lower case alias for Orientation setter
    @orientation.setter
    def orientation(self, value):
        self.Orientation = value

    @property
    def Parent(self):
        return self.datalabel.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return XlDataLabelPosition(self.datalabel.Position)

    # Lower case alias for Position
    @property
    def position(self):
        return self.Position

    @Position.setter
    def Position(self, value):
        self.datalabel.Position = value

    # Lower case alias for Position setter
    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def ReadingOrder(self):
        return XlReadingOrder(self.datalabel.ReadingOrder)

    # Lower case alias for ReadingOrder
    @property
    def readingorder(self):
        return self.ReadingOrder

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.datalabel.ReadingOrder = value

    # Lower case alias for ReadingOrder setter
    @readingorder.setter
    def readingorder(self, value):
        self.ReadingOrder = value

    @property
    def Separator(self):
        return self.datalabel.Separator

    # Lower case alias for Separator
    @property
    def separator(self):
        return self.Separator

    @Separator.setter
    def Separator(self, value):
        self.datalabel.Separator = value

    # Lower case alias for Separator setter
    @separator.setter
    def separator(self, value):
        self.Separator = value

    @property
    def Shadow(self):
        return self.datalabel.Shadow

    # Lower case alias for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.datalabel.Shadow = value

    # Lower case alias for Shadow setter
    @shadow.setter
    def shadow(self, value):
        self.Shadow = value

    @property
    def ShowBubbleSize(self):
        return self.datalabel.ShowBubbleSize

    # Lower case alias for ShowBubbleSize
    @property
    def showbubblesize(self):
        return self.ShowBubbleSize

    @ShowBubbleSize.setter
    def ShowBubbleSize(self, value):
        self.datalabel.ShowBubbleSize = value

    # Lower case alias for ShowBubbleSize setter
    @showbubblesize.setter
    def showbubblesize(self, value):
        self.ShowBubbleSize = value

    @property
    def ShowCategoryName(self):
        return self.datalabel.ShowCategoryName

    # Lower case alias for ShowCategoryName
    @property
    def showcategoryname(self):
        return self.ShowCategoryName

    @ShowCategoryName.setter
    def ShowCategoryName(self, value):
        self.datalabel.ShowCategoryName = value

    # Lower case alias for ShowCategoryName setter
    @showcategoryname.setter
    def showcategoryname(self, value):
        self.ShowCategoryName = value

    @property
    def ShowLegendKey(self):
        return self.datalabel.ShowLegendKey

    # Lower case alias for ShowLegendKey
    @property
    def showlegendkey(self):
        return self.ShowLegendKey

    @ShowLegendKey.setter
    def ShowLegendKey(self, value):
        self.datalabel.ShowLegendKey = value

    # Lower case alias for ShowLegendKey setter
    @showlegendkey.setter
    def showlegendkey(self, value):
        self.ShowLegendKey = value

    @property
    def ShowPercentage(self):
        return self.datalabel.ShowPercentage

    # Lower case alias for ShowPercentage
    @property
    def showpercentage(self):
        return self.ShowPercentage

    @ShowPercentage.setter
    def ShowPercentage(self, value):
        self.datalabel.ShowPercentage = value

    # Lower case alias for ShowPercentage setter
    @showpercentage.setter
    def showpercentage(self, value):
        self.ShowPercentage = value

    @property
    def ShowSeriesName(self):
        return self.datalabel.ShowSeriesName

    # Lower case alias for ShowSeriesName
    @property
    def showseriesname(self):
        return self.ShowSeriesName

    @ShowSeriesName.setter
    def ShowSeriesName(self, value):
        self.datalabel.ShowSeriesName = value

    # Lower case alias for ShowSeriesName setter
    @showseriesname.setter
    def showseriesname(self, value):
        self.ShowSeriesName = value

    @property
    def ShowValue(self):
        return self.datalabel.ShowValue

    # Lower case alias for ShowValue
    @property
    def showvalue(self):
        return self.ShowValue

    @ShowValue.setter
    def ShowValue(self, value):
        self.datalabel.ShowValue = value

    # Lower case alias for ShowValue setter
    @showvalue.setter
    def showvalue(self, value):
        self.ShowValue = value

    @property
    def Text(self):
        return self.datalabel.Text

    # Lower case alias for Text
    @property
    def text(self):
        return self.Text

    @Text.setter
    def Text(self, value):
        self.datalabel.Text = value

    # Lower case alias for Text setter
    @text.setter
    def text(self, value):
        self.Text = value

    @property
    def Top(self):
        return self.datalabel.Top

    # Lower case alias for Top
    @property
    def top(self):
        return self.Top

    @Top.setter
    def Top(self, value):
        self.datalabel.Top = value

    # Lower case alias for Top setter
    @top.setter
    def top(self, value):
        self.Top = value

    @property
    def VerticalAlignment(self):
        return self.datalabel.VerticalAlignment

    # Lower case alias for VerticalAlignment
    @property
    def verticalalignment(self):
        return self.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.datalabel.VerticalAlignment = value

    # Lower case alias for VerticalAlignment setter
    @verticalalignment.setter
    def verticalalignment(self, value):
        self.VerticalAlignment = value

    @property
    def Width(self):
        return self.datalabel.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

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

    # Lower case alias for AutoText
    @property
    def autotext(self):
        return self.AutoText

    @AutoText.setter
    def AutoText(self, value):
        self.datalabels.AutoText = value

    # Lower case alias for AutoText setter
    @autotext.setter
    def autotext(self, value):
        self.AutoText = value

    @property
    def Count(self):
        return self.datalabels.Count

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.datalabels.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.datalabels.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def HorizontalAlignment(self):
        return self.datalabels.HorizontalAlignment

    # Lower case alias for HorizontalAlignment
    @property
    def horizontalalignment(self):
        return self.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.datalabels.HorizontalAlignment = value

    # Lower case alias for HorizontalAlignment setter
    @horizontalalignment.setter
    def horizontalalignment(self, value):
        self.HorizontalAlignment = value

    @property
    def Name(self):
        return self.datalabels.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def NumberFormat(self):
        return self.datalabels.NumberFormat

    # Lower case alias for NumberFormat
    @property
    def numberformat(self):
        return self.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value):
        self.datalabels.NumberFormat = value

    # Lower case alias for NumberFormat setter
    @numberformat.setter
    def numberformat(self, value):
        self.NumberFormat = value

    @property
    def NumberFormatLinked(self):
        return self.datalabels.NumberFormatLinked

    # Lower case alias for NumberFormatLinked
    @property
    def numberformatlinked(self):
        return self.NumberFormatLinked

    @NumberFormatLinked.setter
    def NumberFormatLinked(self, value):
        self.datalabels.NumberFormatLinked = value

    # Lower case alias for NumberFormatLinked setter
    @numberformatlinked.setter
    def numberformatlinked(self, value):
        self.NumberFormatLinked = value

    @property
    def NumberFormatLocal(self):
        return self.datalabels.NumberFormatLocal

    # Lower case alias for NumberFormatLocal
    @property
    def numberformatlocal(self):
        return self.NumberFormatLocal

    @NumberFormatLocal.setter
    def NumberFormatLocal(self, value):
        self.datalabels.NumberFormatLocal = value

    # Lower case alias for NumberFormatLocal setter
    @numberformatlocal.setter
    def numberformatlocal(self, value):
        self.NumberFormatLocal = value

    @property
    def Orientation(self):
        return self.datalabels.Orientation

    # Lower case alias for Orientation
    @property
    def orientation(self):
        return self.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.datalabels.Orientation = value

    # Lower case alias for Orientation setter
    @orientation.setter
    def orientation(self, value):
        self.Orientation = value

    @property
    def Parent(self):
        return self.datalabels.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ReadingOrder(self):
        return XlReadingOrder(self.datalabels.ReadingOrder)

    # Lower case alias for ReadingOrder
    @property
    def readingorder(self):
        return self.ReadingOrder

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.datalabels.ReadingOrder = value

    # Lower case alias for ReadingOrder setter
    @readingorder.setter
    def readingorder(self, value):
        self.ReadingOrder = value

    @property
    def Separator(self):
        return self.datalabels.Separator

    # Lower case alias for Separator
    @property
    def separator(self):
        return self.Separator

    @Separator.setter
    def Separator(self, value):
        self.datalabels.Separator = value

    # Lower case alias for Separator setter
    @separator.setter
    def separator(self, value):
        self.Separator = value

    @property
    def Shadow(self):
        return self.datalabels.Shadow

    # Lower case alias for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.datalabels.Shadow = value

    # Lower case alias for Shadow setter
    @shadow.setter
    def shadow(self, value):
        self.Shadow = value

    @property
    def ShowBubbleSize(self):
        return self.datalabels.ShowBubbleSize

    # Lower case alias for ShowBubbleSize
    @property
    def showbubblesize(self):
        return self.ShowBubbleSize

    @ShowBubbleSize.setter
    def ShowBubbleSize(self, value):
        self.datalabels.ShowBubbleSize = value

    # Lower case alias for ShowBubbleSize setter
    @showbubblesize.setter
    def showbubblesize(self, value):
        self.ShowBubbleSize = value

    @property
    def ShowCategoryName(self):
        return self.datalabels.ShowCategoryName

    # Lower case alias for ShowCategoryName
    @property
    def showcategoryname(self):
        return self.ShowCategoryName

    @ShowCategoryName.setter
    def ShowCategoryName(self, value):
        self.datalabels.ShowCategoryName = value

    # Lower case alias for ShowCategoryName setter
    @showcategoryname.setter
    def showcategoryname(self, value):
        self.ShowCategoryName = value

    @property
    def ShowLegendKey(self):
        return self.datalabels.ShowLegendKey

    # Lower case alias for ShowLegendKey
    @property
    def showlegendkey(self):
        return self.ShowLegendKey

    @ShowLegendKey.setter
    def ShowLegendKey(self, value):
        self.datalabels.ShowLegendKey = value

    # Lower case alias for ShowLegendKey setter
    @showlegendkey.setter
    def showlegendkey(self, value):
        self.ShowLegendKey = value

    @property
    def ShowPercentage(self):
        return self.datalabels.ShowPercentage

    # Lower case alias for ShowPercentage
    @property
    def showpercentage(self):
        return self.ShowPercentage

    @ShowPercentage.setter
    def ShowPercentage(self, value):
        self.datalabels.ShowPercentage = value

    # Lower case alias for ShowPercentage setter
    @showpercentage.setter
    def showpercentage(self, value):
        self.ShowPercentage = value

    @property
    def ShowSeriesName(self):
        return self.datalabels.ShowSeriesName

    # Lower case alias for ShowSeriesName
    @property
    def showseriesname(self):
        return self.ShowSeriesName

    @ShowSeriesName.setter
    def ShowSeriesName(self, value):
        self.datalabels.ShowSeriesName = value

    # Lower case alias for ShowSeriesName setter
    @showseriesname.setter
    def showseriesname(self, value):
        self.ShowSeriesName = value

    @property
    def ShowValue(self):
        return self.datalabels.ShowValue

    # Lower case alias for ShowValue
    @property
    def showvalue(self):
        return self.ShowValue

    @ShowValue.setter
    def ShowValue(self, value):
        self.datalabels.ShowValue = value

    # Lower case alias for ShowValue setter
    @showvalue.setter
    def showvalue(self, value):
        self.ShowValue = value

    @property
    def VerticalAlignment(self):
        return self.datalabels.VerticalAlignment

    # Lower case alias for VerticalAlignment
    @property
    def verticalalignment(self):
        return self.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.datalabels.VerticalAlignment = value

    # Lower case alias for VerticalAlignment setter
    @verticalalignment.setter
    def verticalalignment(self, value):
        self.VerticalAlignment = value

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

    # Lower case alias for Border
    @property
    def border(self):
        return self.Border

    @property
    def Creator(self):
        return self.datatable.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Font(self):
        return ChartFont(self.datatable.Font)

    # Lower case alias for Font
    @property
    def font(self):
        return self.Font

    @property
    def Format(self):
        return ChartFormat(self.datatable.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def HasBorderHorizontal(self):
        return self.datatable.HasBorderHorizontal

    # Lower case alias for HasBorderHorizontal
    @property
    def hasborderhorizontal(self):
        return self.HasBorderHorizontal

    @HasBorderHorizontal.setter
    def HasBorderHorizontal(self, value):
        self.datatable.HasBorderHorizontal = value

    # Lower case alias for HasBorderHorizontal setter
    @hasborderhorizontal.setter
    def hasborderhorizontal(self, value):
        self.HasBorderHorizontal = value

    @property
    def HasBorderOutline(self):
        return self.datatable.HasBorderOutline

    # Lower case alias for HasBorderOutline
    @property
    def hasborderoutline(self):
        return self.HasBorderOutline

    @HasBorderOutline.setter
    def HasBorderOutline(self, value):
        self.datatable.HasBorderOutline = value

    # Lower case alias for HasBorderOutline setter
    @hasborderoutline.setter
    def hasborderoutline(self, value):
        self.HasBorderOutline = value

    @property
    def HasBorderVertical(self):
        return self.datatable.HasBorderVertical

    # Lower case alias for HasBorderVertical
    @property
    def hasbordervertical(self):
        return self.HasBorderVertical

    @HasBorderVertical.setter
    def HasBorderVertical(self, value):
        self.datatable.HasBorderVertical = value

    # Lower case alias for HasBorderVertical setter
    @hasbordervertical.setter
    def hasbordervertical(self, value):
        self.HasBorderVertical = value

    @property
    def Parent(self):
        return self.datatable.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ShowLegendKey(self):
        return self.datatable.ShowLegendKey

    # Lower case alias for ShowLegendKey
    @property
    def showlegendkey(self):
        return self.ShowLegendKey

    @ShowLegendKey.setter
    def ShowLegendKey(self, value):
        self.datatable.ShowLegendKey = value

    # Lower case alias for ShowLegendKey setter
    @showlegendkey.setter
    def showlegendkey(self, value):
        self.ShowLegendKey = value

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

    # Lower case alias for Index
    @property
    def index(self):
        return self.Index

    @property
    def Name(self):
        return self.design.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.design.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.design.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Preserved(self):
        return self.design.Preserved

    # Lower case alias for Preserved
    @property
    def preserved(self):
        return self.Preserved

    @Preserved.setter
    def Preserved(self, value):
        self.design.Preserved = value

    # Lower case alias for Preserved setter
    @preserved.setter
    def preserved(self, value):
        self.Preserved = value

    @property
    def SlideMaster(self):
        return Master(self.design.SlideMaster)

    # Lower case alias for SlideMaster
    @property
    def slidemaster(self):
        return self.SlideMaster

    def Delete(self):
        self.design.Delete()

    def MoveTo(self, toPos=None):
        arguments = com_arguments([toPos])
        self.design.MoveTo(*arguments)


class Designs:

    def __init__(self, designs=None):
        self.designs = designs

    @property
    def Application(self):
        return Application(self.designs.Application)

    @property
    def Count(self):
        return self.designs.Count

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.designs.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Add(self, designName=None, Index=None):
        arguments = com_arguments([designName, Index])
        return self.designs.Add(*arguments)

    def Clone(self, pOriginal=None, Index=None):
        arguments = com_arguments([pOriginal, Index])
        return self.designs.Clone(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.designs.Item(*arguments)

    def Load(self, TemplateName=None, Index=None):
        arguments = com_arguments([TemplateName, Index])
        return self.designs.Load(*arguments)


class DisplayUnitLabel:

    def __init__(self, displayunitlabel=None):
        self.displayunitlabel = displayunitlabel

    @property
    def Application(self):
        return self.displayunitlabel.Application

    @property
    def Caption(self):
        return self.displayunitlabel.Caption

    # Lower case alias for Caption
    @property
    def caption(self):
        return self.Caption

    @Caption.setter
    def Caption(self, value):
        self.displayunitlabel.Caption = value

    # Lower case alias for Caption setter
    @caption.setter
    def caption(self, value):
        self.Caption = value

    def Characters(self, Start=None, Length=None):
        arguments = com_arguments([Start, Length])
        if callable(self.displayunitlabel.Characters):
            return ChartCharacters(self.displayunitlabel.Characters(*arguments))
        else:
            return ChartCharacters(self.displayunitlabel.GetCharacters(*arguments))

    # Lower case alias for Characters
    def characters(self, Start=None, Length=None):
        arguments = [Start, Length]
        return self.Characters(*arguments)

    @property
    def Creator(self):
        return self.displayunitlabel.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.displayunitlabel.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def Formula(self):
        return self.displayunitlabel.Formula

    # Lower case alias for Formula
    @property
    def formula(self):
        return self.Formula

    @Formula.setter
    def Formula(self, value):
        self.displayunitlabel.Formula = value

    # Lower case alias for Formula setter
    @formula.setter
    def formula(self, value):
        self.Formula = value

    @property
    def FormulaLocal(self):
        return self.displayunitlabel.FormulaLocal

    # Lower case alias for FormulaLocal
    @property
    def formulalocal(self):
        return self.FormulaLocal

    @FormulaLocal.setter
    def FormulaLocal(self, value):
        self.displayunitlabel.FormulaLocal = value

    # Lower case alias for FormulaLocal setter
    @formulalocal.setter
    def formulalocal(self, value):
        self.FormulaLocal = value

    @property
    def FormulaR1C1(self):
        return self.displayunitlabel.FormulaR1C1

    # Lower case alias for FormulaR1C1
    @property
    def formular1c1(self):
        return self.FormulaR1C1

    @FormulaR1C1.setter
    def FormulaR1C1(self, value):
        self.displayunitlabel.FormulaR1C1 = value

    # Lower case alias for FormulaR1C1 setter
    @formular1c1.setter
    def formular1c1(self, value):
        self.FormulaR1C1 = value

    @property
    def FormulaR1C1Local(self):
        return self.displayunitlabel.FormulaR1C1Local

    # Lower case alias for FormulaR1C1Local
    @property
    def formular1c1local(self):
        return self.FormulaR1C1Local

    @FormulaR1C1Local.setter
    def FormulaR1C1Local(self, value):
        self.displayunitlabel.FormulaR1C1Local = value

    # Lower case alias for FormulaR1C1Local setter
    @formular1c1local.setter
    def formular1c1local(self, value):
        self.FormulaR1C1Local = value

    @property
    def Height(self):
        return self.displayunitlabel.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @property
    def HorizontalAlignment(self):
        return self.displayunitlabel.HorizontalAlignment

    # Lower case alias for HorizontalAlignment
    @property
    def horizontalalignment(self):
        return self.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.displayunitlabel.HorizontalAlignment = value

    # Lower case alias for HorizontalAlignment setter
    @horizontalalignment.setter
    def horizontalalignment(self, value):
        self.HorizontalAlignment = value

    @property
    def Left(self):
        return self.displayunitlabel.Left

    # Lower case alias for Left
    @property
    def left(self):
        return self.Left

    @Left.setter
    def Left(self, value):
        self.displayunitlabel.Left = value

    # Lower case alias for Left setter
    @left.setter
    def left(self, value):
        self.Left = value

    @property
    def Name(self):
        return self.displayunitlabel.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Orientation(self):
        return self.displayunitlabel.Orientation

    # Lower case alias for Orientation
    @property
    def orientation(self):
        return self.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.displayunitlabel.Orientation = value

    # Lower case alias for Orientation setter
    @orientation.setter
    def orientation(self, value):
        self.Orientation = value

    @property
    def Parent(self):
        return self.displayunitlabel.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return XlChartElementPosition(self.displayunitlabel.Position)

    # Lower case alias for Position
    @property
    def position(self):
        return self.Position

    @Position.setter
    def Position(self, value):
        self.displayunitlabel.Position = value

    # Lower case alias for Position setter
    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def ReadingOrder(self):
        return XlReadingOrder(self.displayunitlabel.ReadingOrder)

    # Lower case alias for ReadingOrder
    @property
    def readingorder(self):
        return self.ReadingOrder

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.displayunitlabel.ReadingOrder = value

    # Lower case alias for ReadingOrder setter
    @readingorder.setter
    def readingorder(self, value):
        self.ReadingOrder = value

    @property
    def Shadow(self):
        return self.displayunitlabel.Shadow

    # Lower case alias for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.displayunitlabel.Shadow = value

    # Lower case alias for Shadow setter
    @shadow.setter
    def shadow(self, value):
        self.Shadow = value

    @property
    def Text(self):
        return self.displayunitlabel.Text

    # Lower case alias for Text
    @property
    def text(self):
        return self.Text

    @Text.setter
    def Text(self, value):
        self.displayunitlabel.Text = value

    # Lower case alias for Text setter
    @text.setter
    def text(self, value):
        self.Text = value

    @property
    def Top(self):
        return self.displayunitlabel.Top

    # Lower case alias for Top
    @property
    def top(self):
        return self.Top

    @Top.setter
    def Top(self, value):
        self.displayunitlabel.Top = value

    # Lower case alias for Top setter
    @top.setter
    def top(self, value):
        self.Top = value

    @property
    def VerticalAlignment(self):
        return self.displayunitlabel.VerticalAlignment

    # Lower case alias for VerticalAlignment
    @property
    def verticalalignment(self):
        return self.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.displayunitlabel.VerticalAlignment = value

    # Lower case alias for VerticalAlignment setter
    @verticalalignment.setter
    def verticalalignment(self, value):
        self.VerticalAlignment = value

    @property
    def Width(self):
        return self.displayunitlabel.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

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

    # Lower case alias for Active
    @property
    def active(self):
        return self.Active

    @property
    def ActivePane(self):
        return Pane(self.documentwindow.ActivePane)

    # Lower case alias for ActivePane
    @property
    def activepane(self):
        return self.ActivePane

    @property
    def Application(self):
        return Application(self.documentwindow.Application)

    @property
    def BlackAndWhite(self):
        return self.documentwindow.BlackAndWhite

    # Lower case alias for BlackAndWhite
    @property
    def blackandwhite(self):
        return self.BlackAndWhite

    @BlackAndWhite.setter
    def BlackAndWhite(self, value):
        self.documentwindow.BlackAndWhite = value

    # Lower case alias for BlackAndWhite setter
    @blackandwhite.setter
    def blackandwhite(self, value):
        self.BlackAndWhite = value

    @property
    def Caption(self):
        return self.documentwindow.Caption

    # Lower case alias for Caption
    @property
    def caption(self):
        return self.Caption

    @property
    def Height(self):
        return self.documentwindow.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @Height.setter
    def Height(self, value):
        self.documentwindow.Height = value

    # Lower case alias for Height setter
    @height.setter
    def height(self, value):
        self.Height = value

    @property
    def Left(self):
        return self.documentwindow.Left

    # Lower case alias for Left
    @property
    def left(self):
        return self.Left

    @Left.setter
    def Left(self, value):
        self.documentwindow.Left = value

    # Lower case alias for Left setter
    @left.setter
    def left(self, value):
        self.Left = value

    @property
    def Panes(self):
        return Panes(self.documentwindow.Panes)

    # Lower case alias for Panes
    @property
    def panes(self):
        return self.Panes

    @property
    def Parent(self):
        return self.documentwindow.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Presentation(self):
        return Presentation(self.documentwindow.Presentation)

    # Lower case alias for Presentation
    @property
    def presentation(self):
        return self.Presentation

    @property
    def Selection(self):
        return Selection(self.documentwindow.Selection)

    # Lower case alias for Selection
    @property
    def selection(self):
        return self.Selection

    @property
    def SplitHorizontal(self):
        return self.documentwindow.SplitHorizontal

    # Lower case alias for SplitHorizontal
    @property
    def splithorizontal(self):
        return self.SplitHorizontal

    @SplitHorizontal.setter
    def SplitHorizontal(self, value):
        self.documentwindow.SplitHorizontal = value

    # Lower case alias for SplitHorizontal setter
    @splithorizontal.setter
    def splithorizontal(self, value):
        self.SplitHorizontal = value

    @property
    def SplitVertical(self):
        return self.documentwindow.SplitVertical

    # Lower case alias for SplitVertical
    @property
    def splitvertical(self):
        return self.SplitVertical

    @SplitVertical.setter
    def SplitVertical(self, value):
        self.documentwindow.SplitVertical = value

    # Lower case alias for SplitVertical setter
    @splitvertical.setter
    def splitvertical(self, value):
        self.SplitVertical = value

    @property
    def Top(self):
        return self.documentwindow.Top

    # Lower case alias for Top
    @property
    def top(self):
        return self.Top

    @Top.setter
    def Top(self, value):
        self.documentwindow.Top = value

    # Lower case alias for Top setter
    @top.setter
    def top(self, value):
        self.Top = value

    @property
    def View(self):
        return View(self.documentwindow.View)

    # Lower case alias for View
    @property
    def view(self):
        return self.View

    @property
    def ViewType(self):
        return self.documentwindow.ViewType

    # Lower case alias for ViewType
    @property
    def viewtype(self):
        return self.ViewType

    @ViewType.setter
    def ViewType(self, value):
        self.documentwindow.ViewType = value

    # Lower case alias for ViewType setter
    @viewtype.setter
    def viewtype(self, value):
        self.ViewType = value

    @property
    def Width(self):
        return self.documentwindow.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

    @Width.setter
    def Width(self, value):
        self.documentwindow.Width = value

    # Lower case alias for Width setter
    @width.setter
    def width(self, value):
        self.Width = value

    @property
    def WindowState(self):
        return self.documentwindow.WindowState

    # Lower case alias for WindowState
    @property
    def windowstate(self):
        return self.WindowState

    @WindowState.setter
    def WindowState(self, value):
        self.documentwindow.WindowState = value

    # Lower case alias for WindowState setter
    @windowstate.setter
    def windowstate(self, value):
        self.WindowState = value

    def Activate(self):
        self.documentwindow.Activate()

    def Close(self):
        self.documentwindow.Close()

    def ExpandSection(self, sectionIndex=None, Expand=None):
        arguments = com_arguments([sectionIndex, Expand])
        self.documentwindow.ExpandSection(*arguments)

    def FitToPage(self):
        self.documentwindow.FitToPage()

    def IsSectionExpanded(self, sectionIndex=None):
        arguments = com_arguments([sectionIndex])
        return self.documentwindow.IsSectionExpanded(*arguments)

    def LargeScroll(self, Down=None, Up=None, ToRight=None, ToLeft=None):
        arguments = com_arguments([Down, Up, ToRight, ToLeft])
        self.documentwindow.LargeScroll(*arguments)

    def NewWindow(self):
        return self.documentwindow.NewWindow()

    def PointsToScreenPixelsX(self, Points=None):
        arguments = com_arguments([Points])
        return self.documentwindow.PointsToScreenPixelsX(*arguments)

    def PointsToScreenPixelsY(self, Points=None):
        arguments = com_arguments([Points])
        return self.documentwindow.PointsToScreenPixelsY(*arguments)

    def RangeFromPoint(self, x=None, y=None):
        arguments = com_arguments([x, y])
        self.documentwindow.RangeFromPoint(*arguments)

    def ScrollIntoView(self, Left=None, Top=None, Width=None, Height=None, Start=None):
        arguments = com_arguments([Left, Top, Width, Height, Start])
        self.documentwindow.ScrollIntoView(*arguments)

    def SmallScroll(self, Down=None, Up=None, ToRight=None, ToLeft=None):
        arguments = com_arguments([Down, Up, ToRight, ToLeft])
        self.documentwindow.SmallScroll(*arguments)


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.documentwindows.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Arrange(self, arrangeStyle=None):
        arguments = com_arguments([arrangeStyle])
        return self.documentwindows.Arrange(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.documentwindows.Item(*arguments)


class DownBars:

    def __init__(self, downbars=None):
        self.downbars = downbars

    @property
    def Application(self):
        return self.downbars.Application

    @property
    def Creator(self):
        return self.downbars.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.downbars.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def Name(self):
        return self.downbars.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.downbars.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

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

    # Lower case alias for Border
    @property
    def border(self):
        return self.Border

    @property
    def Creator(self):
        return self.droplines.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.droplines.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def Name(self):
        return self.droplines.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.droplines.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

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

    # Lower case alias for Behaviors
    @property
    def behaviors(self):
        return self.Behaviors

    @property
    def DisplayName(self):
        return self.effect.DisplayName

    # Lower case alias for DisplayName
    @property
    def displayname(self):
        return self.DisplayName

    @property
    def EffectInformation(self):
        return EffectInformation(self.effect.EffectInformation)

    # Lower case alias for EffectInformation
    @property
    def effectinformation(self):
        return self.EffectInformation

    @property
    def EffectParameters(self):
        return EffectParameters(self.effect.EffectParameters)

    # Lower case alias for EffectParameters
    @property
    def effectparameters(self):
        return self.EffectParameters

    @property
    def EffectType(self):
        return self.effect.EffectType

    # Lower case alias for EffectType
    @property
    def effecttype(self):
        return self.EffectType

    @EffectType.setter
    def EffectType(self, value):
        self.effect.EffectType = value

    # Lower case alias for EffectType setter
    @effecttype.setter
    def effecttype(self, value):
        self.EffectType = value

    @property
    def Exit(self):
        return self.effect.Exit

    # Lower case alias for Exit
    @property
    def exit(self):
        return self.Exit

    @Exit.setter
    def Exit(self, value):
        self.effect.Exit = value

    # Lower case alias for Exit setter
    @exit.setter
    def exit(self, value):
        self.Exit = value

    @property
    def Index(self):
        return self.effect.Index

    # Lower case alias for Index
    @property
    def index(self):
        return self.Index

    @property
    def Paragraph(self):
        return self.effect.Paragraph

    # Lower case alias for Paragraph
    @property
    def paragraph(self):
        return self.Paragraph

    @Paragraph.setter
    def Paragraph(self, value):
        self.effect.Paragraph = value

    # Lower case alias for Paragraph setter
    @paragraph.setter
    def paragraph(self, value):
        self.Paragraph = value

    @property
    def Parent(self):
        return self.effect.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Shape(self):
        return Shape(self.effect.Shape)

    # Lower case alias for Shape
    @property
    def shape(self):
        return self.Shape

    @property
    def TextRangeLength(self):
        return self.effect.TextRangeLength

    # Lower case alias for TextRangeLength
    @property
    def textrangelength(self):
        return self.TextRangeLength

    @TextRangeLength.setter
    def TextRangeLength(self, value):
        self.effect.TextRangeLength = value

    # Lower case alias for TextRangeLength setter
    @textrangelength.setter
    def textrangelength(self, value):
        self.TextRangeLength = value

    @property
    def TextRangeStart(self):
        return self.effect.TextRangeStart

    # Lower case alias for TextRangeStart
    @property
    def textrangestart(self):
        return self.TextRangeStart

    @TextRangeStart.setter
    def TextRangeStart(self, value):
        self.effect.TextRangeStart = value

    # Lower case alias for TextRangeStart setter
    @textrangestart.setter
    def textrangestart(self, value):
        self.TextRangeStart = value

    @property
    def Timing(self):
        return Timing(self.effect.Timing)

    # Lower case alias for Timing
    @property
    def timing(self):
        return self.Timing

    def Delete(self):
        self.effect.Delete()

    def MoveAfter(self, Effect=None):
        arguments = com_arguments([Effect])
        self.effect.MoveAfter(*arguments)

    def MoveBefore(self, Effect=None):
        arguments = com_arguments([Effect])
        self.effect.MoveBefore(*arguments)

    def MoveTo(self, toPos=None):
        arguments = com_arguments([toPos])
        self.effect.MoveTo(*arguments)


class EffectInformation:

    def __init__(self, effectinformation=None):
        self.effectinformation = effectinformation

    @property
    def AfterEffect(self):
        return PpAfterEffect(self.effectinformation.AfterEffect)

    # Lower case alias for AfterEffect
    @property
    def aftereffect(self):
        return self.AfterEffect

    @property
    def AnimateBackground(self):
        return self.effectinformation.AnimateBackground

    # Lower case alias for AnimateBackground
    @property
    def animatebackground(self):
        return self.AnimateBackground

    @property
    def AnimateTextInReverse(self):
        return self.effectinformation.AnimateTextInReverse

    # Lower case alias for AnimateTextInReverse
    @property
    def animatetextinreverse(self):
        return self.AnimateTextInReverse

    @AnimateTextInReverse.setter
    def AnimateTextInReverse(self, value):
        self.effectinformation.AnimateTextInReverse = value

    # Lower case alias for AnimateTextInReverse setter
    @animatetextinreverse.setter
    def animatetextinreverse(self, value):
        self.AnimateTextInReverse = value

    @property
    def Application(self):
        return Application(self.effectinformation.Application)

    @property
    def BuildByLevelEffect(self):
        return self.effectinformation.BuildByLevelEffect

    # Lower case alias for BuildByLevelEffect
    @property
    def buildbyleveleffect(self):
        return self.BuildByLevelEffect

    @property
    def Dim(self):
        return ColorFormat(self.effectinformation.Dim)

    # Lower case alias for Dim
    @property
    def dim(self):
        return self.Dim

    @property
    def Parent(self):
        return self.effectinformation.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PlaySettings(self):
        return PlaySettings(self.effectinformation.PlaySettings)

    # Lower case alias for PlaySettings
    @property
    def playsettings(self):
        return self.PlaySettings

    @property
    def SoundEffect(self):
        return SoundEffect(self.effectinformation.SoundEffect)

    # Lower case alias for SoundEffect
    @property
    def soundeffect(self):
        return self.SoundEffect

    @property
    def TextUnitEffect(self):
        return self.effectinformation.TextUnitEffect

    # Lower case alias for TextUnitEffect
    @property
    def textuniteffect(self):
        return self.TextUnitEffect


class EffectParameters:

    def __init__(self, effectparameters=None):
        self.effectparameters = effectparameters

    @property
    def Amount(self):
        return self.effectparameters.Amount

    # Lower case alias for Amount
    @property
    def amount(self):
        return self.Amount

    @Amount.setter
    def Amount(self, value):
        self.effectparameters.Amount = value

    # Lower case alias for Amount setter
    @amount.setter
    def amount(self, value):
        self.Amount = value

    @property
    def Application(self):
        return Application(self.effectparameters.Application)

    @property
    def Color2(self):
        return ColorFormat(self.effectparameters.Color2)

    # Lower case alias for Color2
    @property
    def color2(self):
        return self.Color2

    @property
    def Direction(self):
        return self.effectparameters.Direction

    # Lower case alias for Direction
    @property
    def direction(self):
        return self.Direction

    @Direction.setter
    def Direction(self, value):
        self.effectparameters.Direction = value

    # Lower case alias for Direction setter
    @direction.setter
    def direction(self, value):
        self.Direction = value

    @property
    def FontName(self):
        return self.effectparameters.FontName

    # Lower case alias for FontName
    @property
    def fontname(self):
        return self.FontName

    @FontName.setter
    def FontName(self, value):
        self.effectparameters.FontName = value

    # Lower case alias for FontName setter
    @fontname.setter
    def fontname(self, value):
        self.FontName = value

    @property
    def Parent(self):
        return self.effectparameters.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Relative(self):
        return self.effectparameters.Relative

    # Lower case alias for Relative
    @property
    def relative(self):
        return self.Relative

    @Relative.setter
    def Relative(self, value):
        self.effectparameters.Relative = value

    # Lower case alias for Relative setter
    @relative.setter
    def relative(self, value):
        self.Relative = value

    @property
    def Size(self):
        return self.effectparameters.Size

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @Size.setter
    def Size(self, value):
        self.effectparameters.Size = value

    # Lower case alias for Size setter
    @size.setter
    def size(self, value):
        self.Size = value


class ErrorBars:

    def __init__(self, errorbars=None):
        self.errorbars = errorbars

    @property
    def Application(self):
        return self.errorbars.Application

    @property
    def Border(self):
        return ChartBorder(self.errorbars.Border)

    # Lower case alias for Border
    @property
    def border(self):
        return self.Border

    @property
    def Creator(self):
        return self.errorbars.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def EndStyle(self):
        return self.errorbars.EndStyle

    # Lower case alias for EndStyle
    @property
    def endstyle(self):
        return self.EndStyle

    @EndStyle.setter
    def EndStyle(self, value):
        self.errorbars.EndStyle = value

    # Lower case alias for EndStyle setter
    @endstyle.setter
    def endstyle(self, value):
        self.EndStyle = value

    @property
    def Format(self):
        return ChartFormat(self.errorbars.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def Name(self):
        return self.errorbars.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.errorbars.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.extracolors.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Add(self, Type=None):
        arguments = com_arguments([Type])
        self.extracolors.Add(*arguments)

    def Clear(self):
        self.extracolors.Clear()

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return MsoThemeColorSchemeIndex(self.extracolors.Item(*arguments))


class FileConverter:

    def __init__(self, fileconverter=None):
        self.fileconverter = fileconverter

    @property
    def Application(self):
        return Application(self.fileconverter.Application)

    @property
    def CanOpen(self):
        return self.fileconverter.CanOpen

    # Lower case alias for CanOpen
    @property
    def canopen(self):
        return self.CanOpen

    @property
    def CanSave(self):
        return self.fileconverter.CanSave

    # Lower case alias for CanSave
    @property
    def cansave(self):
        return self.CanSave

    @property
    def ClassName(self):
        return self.fileconverter.ClassName

    # Lower case alias for ClassName
    @property
    def classname(self):
        return self.ClassName

    @property
    def Creator(self):
        return self.fileconverter.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Extensions(self):
        return FileConverter(self.fileconverter.Extensions)

    # Lower case alias for Extensions
    @property
    def extensions(self):
        return self.Extensions

    @property
    def FormatName(self):
        return self.fileconverter.FormatName

    # Lower case alias for FormatName
    @property
    def formatname(self):
        return self.FormatName

    @property
    def Name(self):
        return self.fileconverter.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def OpenFormat(self):
        return self.fileconverter.OpenFormat

    # Lower case alias for OpenFormat
    @property
    def openformat(self):
        return self.OpenFormat

    @property
    def Parent(self):
        return FileConverter(self.fileconverter.Parent)

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Path(self):
        return self.fileconverter.Path

    # Lower case alias for Path
    @property
    def path(self):
        return self.Path

    @property
    def SaveFormat(self):
        return self.fileconverter.SaveFormat

    # Lower case alias for SaveFormat
    @property
    def saveformat(self):
        return self.SaveFormat


class FileConverters:

    def __init__(self, fileconverters=None):
        self.fileconverters = fileconverters

    def __call__(self, item):
        return FileConverter(self.fileconverters(item))

    @property
    def Count(self):
        return self.fileconverters.Count

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.fileconverters.Item(*arguments)


class FillFormat:

    def __init__(self, fillformat=None):
        self.fillformat = fillformat

    @property
    def Application(self):
        return Application(self.fillformat.Application)

    @property
    def BackColor(self):
        return ColorFormat(self.fillformat.BackColor)

    # Lower case alias for BackColor
    @property
    def backcolor(self):
        return self.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.fillformat.BackColor = value

    # Lower case alias for BackColor setter
    @backcolor.setter
    def backcolor(self, value):
        self.BackColor = value

    @property
    def Creator(self):
        return self.fillformat.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def ForeColor(self):
        return ColorFormat(self.fillformat.ForeColor)

    # Lower case alias for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.fillformat.ForeColor = value

    # Lower case alias for ForeColor setter
    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def GradientAngle(self):
        return self.fillformat.GradientAngle

    # Lower case alias for GradientAngle
    @property
    def gradientangle(self):
        return self.GradientAngle

    @GradientAngle.setter
    def GradientAngle(self, value):
        self.fillformat.GradientAngle = value

    # Lower case alias for GradientAngle setter
    @gradientangle.setter
    def gradientangle(self, value):
        self.GradientAngle = value

    @property
    def GradientColorType(self):
        return self.fillformat.GradientColorType

    # Lower case alias for GradientColorType
    @property
    def gradientcolortype(self):
        return self.GradientColorType

    @property
    def GradientDegree(self):
        return self.fillformat.GradientDegree

    # Lower case alias for GradientDegree
    @property
    def gradientdegree(self):
        return self.GradientDegree

    @property
    def GradientStops(self):
        return self.fillformat.GradientStops

    # Lower case alias for GradientStops
    @property
    def gradientstops(self):
        return self.GradientStops

    @property
    def GradientStyle(self):
        return self.fillformat.GradientStyle

    # Lower case alias for GradientStyle
    @property
    def gradientstyle(self):
        return self.GradientStyle

    @property
    def GradientVariant(self):
        return self.fillformat.GradientVariant

    # Lower case alias for GradientVariant
    @property
    def gradientvariant(self):
        return self.GradientVariant

    @property
    def Parent(self):
        return self.fillformat.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Pattern(self):
        return self.fillformat.Pattern

    # Lower case alias for Pattern
    @property
    def pattern(self):
        return self.Pattern

    @property
    def PictureEffects(self):
        return self.fillformat.PictureEffects

    # Lower case alias for PictureEffects
    @property
    def pictureeffects(self):
        return self.PictureEffects

    @property
    def PresetGradientType(self):
        return self.fillformat.PresetGradientType

    # Lower case alias for PresetGradientType
    @property
    def presetgradienttype(self):
        return self.PresetGradientType

    @property
    def PresetTexture(self):
        return self.fillformat.PresetTexture

    # Lower case alias for PresetTexture
    @property
    def presettexture(self):
        return self.PresetTexture

    @property
    def RotateWithObject(self):
        return self.fillformat.RotateWithObject

    # Lower case alias for RotateWithObject
    @property
    def rotatewithobject(self):
        return self.RotateWithObject

    @RotateWithObject.setter
    def RotateWithObject(self, value):
        self.fillformat.RotateWithObject = value

    # Lower case alias for RotateWithObject setter
    @rotatewithobject.setter
    def rotatewithobject(self, value):
        self.RotateWithObject = value

    @property
    def TextureAlignment(self):
        return self.fillformat.TextureAlignment

    # Lower case alias for TextureAlignment
    @property
    def texturealignment(self):
        return self.TextureAlignment

    @TextureAlignment.setter
    def TextureAlignment(self, value):
        self.fillformat.TextureAlignment = value

    # Lower case alias for TextureAlignment setter
    @texturealignment.setter
    def texturealignment(self, value):
        self.TextureAlignment = value

    @property
    def TextureHorizontalScale(self):
        return self.fillformat.TextureHorizontalScale

    # Lower case alias for TextureHorizontalScale
    @property
    def texturehorizontalscale(self):
        return self.TextureHorizontalScale

    @TextureHorizontalScale.setter
    def TextureHorizontalScale(self, value):
        self.fillformat.TextureHorizontalScale = value

    # Lower case alias for TextureHorizontalScale setter
    @texturehorizontalscale.setter
    def texturehorizontalscale(self, value):
        self.TextureHorizontalScale = value

    @property
    def TextureName(self):
        return self.fillformat.TextureName

    # Lower case alias for TextureName
    @property
    def texturename(self):
        return self.TextureName

    @property
    def TextureOffsetX(self):
        return self.fillformat.TextureOffsetX

    # Lower case alias for TextureOffsetX
    @property
    def textureoffsetx(self):
        return self.TextureOffsetX

    @TextureOffsetX.setter
    def TextureOffsetX(self, value):
        self.fillformat.TextureOffsetX = value

    # Lower case alias for TextureOffsetX setter
    @textureoffsetx.setter
    def textureoffsetx(self, value):
        self.TextureOffsetX = value

    @property
    def TextureOffsetY(self):
        return self.fillformat.TextureOffsetY

    # Lower case alias for TextureOffsetY
    @property
    def textureoffsety(self):
        return self.TextureOffsetY

    @TextureOffsetY.setter
    def TextureOffsetY(self, value):
        self.fillformat.TextureOffsetY = value

    # Lower case alias for TextureOffsetY setter
    @textureoffsety.setter
    def textureoffsety(self, value):
        self.TextureOffsetY = value

    @property
    def TextureTile(self):
        return self.fillformat.TextureTile

    # Lower case alias for TextureTile
    @property
    def texturetile(self):
        return self.TextureTile

    @TextureTile.setter
    def TextureTile(self, value):
        self.fillformat.TextureTile = value

    # Lower case alias for TextureTile setter
    @texturetile.setter
    def texturetile(self, value):
        self.TextureTile = value

    @property
    def TextureType(self):
        return self.fillformat.TextureType

    # Lower case alias for TextureType
    @property
    def texturetype(self):
        return self.TextureType

    @property
    def TextureVerticalScale(self):
        return self.fillformat.TextureVerticalScale

    # Lower case alias for TextureVerticalScale
    @property
    def textureverticalscale(self):
        return self.TextureVerticalScale

    @TextureVerticalScale.setter
    def TextureVerticalScale(self, value):
        self.fillformat.TextureVerticalScale = value

    # Lower case alias for TextureVerticalScale setter
    @textureverticalscale.setter
    def textureverticalscale(self, value):
        self.TextureVerticalScale = value

    @property
    def Transparency(self):
        return self.fillformat.Transparency

    # Lower case alias for Transparency
    @property
    def transparency(self):
        return self.Transparency

    @Transparency.setter
    def Transparency(self, value):
        self.fillformat.Transparency = value

    # Lower case alias for Transparency setter
    @transparency.setter
    def transparency(self, value):
        self.Transparency = value

    @property
    def Type(self):
        return self.fillformat.Type

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @property
    def Visible(self):
        return self.fillformat.Visible

    # Lower case alias for Visible
    @property
    def visible(self):
        return self.Visible

    @Visible.setter
    def Visible(self, value):
        self.fillformat.Visible = value

    # Lower case alias for Visible setter
    @visible.setter
    def visible(self, value):
        self.Visible = value

    def Background(self):
        self.fillformat.Background()

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


class FilterEffect:

    def __init__(self, filtereffect=None):
        self.filtereffect = filtereffect

    @property
    def Application(self):
        return Application(self.filtereffect.Application)

    @property
    def Parent(self):
        return self.filtereffect.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Reveal(self):
        return self.filtereffect.Reveal

    # Lower case alias for Reveal
    @property
    def reveal(self):
        return self.Reveal

    @Reveal.setter
    def Reveal(self, value):
        self.filtereffect.Reveal = value

    # Lower case alias for Reveal setter
    @reveal.setter
    def reveal(self, value):
        self.Reveal = value

    @property
    def Subtype(self):
        return self.filtereffect.Subtype

    # Lower case alias for Subtype
    @property
    def subtype(self):
        return self.Subtype

    @Subtype.setter
    def Subtype(self, value):
        self.filtereffect.Subtype = value

    # Lower case alias for Subtype setter
    @subtype.setter
    def subtype(self, value):
        self.Subtype = value

    @property
    def Type(self):
        return self.filtereffect.Type

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @Type.setter
    def Type(self, value):
        self.filtereffect.Type = value

    # Lower case alias for Type setter
    @type.setter
    def type(self, value):
        self.Type = value


class Floor:

    def __init__(self, floor=None):
        self.floor = floor

    @property
    def Application(self):
        return self.floor.Application

    @property
    def Creator(self):
        return self.floor.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.floor.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def Name(self):
        return self.floor.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.floor.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PictureType(self):
        return self.floor.PictureType

    # Lower case alias for PictureType
    @property
    def picturetype(self):
        return self.PictureType

    @PictureType.setter
    def PictureType(self, value):
        self.floor.PictureType = value

    # Lower case alias for PictureType setter
    @picturetype.setter
    def picturetype(self, value):
        self.PictureType = value

    @property
    def Thickness(self):
        return self.floor.Thickness

    # Lower case alias for Thickness
    @property
    def thickness(self):
        return self.Thickness

    @Thickness.setter
    def Thickness(self, value):
        self.floor.Thickness = value

    # Lower case alias for Thickness setter
    @thickness.setter
    def thickness(self, value):
        self.Thickness = value

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

    # Lower case alias for AutoRotateNumbers
    @property
    def autorotatenumbers(self):
        return self.AutoRotateNumbers

    @AutoRotateNumbers.setter
    def AutoRotateNumbers(self, value):
        self.font.AutoRotateNumbers = value

    # Lower case alias for AutoRotateNumbers setter
    @autorotatenumbers.setter
    def autorotatenumbers(self, value):
        self.AutoRotateNumbers = value

    @property
    def BaselineOffset(self):
        return self.font.BaselineOffset

    # Lower case alias for BaselineOffset
    @property
    def baselineoffset(self):
        return self.BaselineOffset

    @BaselineOffset.setter
    def BaselineOffset(self, value):
        self.font.BaselineOffset = value

    # Lower case alias for BaselineOffset setter
    @baselineoffset.setter
    def baselineoffset(self, value):
        self.BaselineOffset = value

    @property
    def Bold(self):
        return self.font.Bold

    # Lower case alias for Bold
    @property
    def bold(self):
        return self.Bold

    @Bold.setter
    def Bold(self, value):
        self.font.Bold = value

    # Lower case alias for Bold setter
    @bold.setter
    def bold(self, value):
        self.Bold = value

    @property
    def Color(self):
        return Font(self.font.Color)

    # Lower case alias for Color
    @property
    def color(self):
        return self.Color

    @Color.setter
    def Color(self, value):
        self.font.Color = value

    # Lower case alias for Color setter
    @color.setter
    def color(self, value):
        self.Color = value

    @property
    def Embeddable(self):
        return self.font.Embeddable

    # Lower case alias for Embeddable
    @property
    def embeddable(self):
        return self.Embeddable

    @property
    def Embedded(self):
        return self.font.Embedded

    # Lower case alias for Embedded
    @property
    def embedded(self):
        return self.Embedded

    @property
    def Emboss(self):
        return self.font.Emboss

    # Lower case alias for Emboss
    @property
    def emboss(self):
        return self.Emboss

    @Emboss.setter
    def Emboss(self, value):
        self.font.Emboss = value

    # Lower case alias for Emboss setter
    @emboss.setter
    def emboss(self, value):
        self.Emboss = value

    @property
    def Name(self):
        return self.font.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.font.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def NameAscii(self):
        return self.font.NameAscii

    # Lower case alias for NameAscii
    @property
    def nameascii(self):
        return self.NameAscii

    @NameAscii.setter
    def NameAscii(self, value):
        self.font.NameAscii = value

    # Lower case alias for NameAscii setter
    @nameascii.setter
    def nameascii(self, value):
        self.NameAscii = value

    @property
    def NameComplexScript(self):
        return self.font.NameComplexScript

    # Lower case alias for NameComplexScript
    @property
    def namecomplexscript(self):
        return self.NameComplexScript

    @NameComplexScript.setter
    def NameComplexScript(self, value):
        self.font.NameComplexScript = value

    # Lower case alias for NameComplexScript setter
    @namecomplexscript.setter
    def namecomplexscript(self, value):
        self.NameComplexScript = value

    @property
    def NameFarEast(self):
        return self.font.NameFarEast

    # Lower case alias for NameFarEast
    @property
    def namefareast(self):
        return self.NameFarEast

    @NameFarEast.setter
    def NameFarEast(self, value):
        self.font.NameFarEast = value

    # Lower case alias for NameFarEast setter
    @namefareast.setter
    def namefareast(self, value):
        self.NameFarEast = value

    @property
    def NameOther(self):
        return self.font.NameOther

    # Lower case alias for NameOther
    @property
    def nameother(self):
        return self.NameOther

    @NameOther.setter
    def NameOther(self, value):
        self.font.NameOther = value

    # Lower case alias for NameOther setter
    @nameother.setter
    def nameother(self, value):
        self.NameOther = value

    @property
    def Parent(self):
        return self.font.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Shadow(self):
        return self.font.Shadow

    # Lower case alias for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.font.Shadow = value

    # Lower case alias for Shadow setter
    @shadow.setter
    def shadow(self, value):
        self.Shadow = value

    @property
    def Size(self):
        return self.font.Size

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @Size.setter
    def Size(self, value):
        self.font.Size = value

    # Lower case alias for Size setter
    @size.setter
    def size(self, value):
        self.Size = value

    @property
    def Subscript(self):
        return self.font.Subscript

    # Lower case alias for Subscript
    @property
    def subscript(self):
        return self.Subscript

    @Subscript.setter
    def Subscript(self, value):
        self.font.Subscript = value

    # Lower case alias for Subscript setter
    @subscript.setter
    def subscript(self, value):
        self.Subscript = value

    @property
    def Superscript(self):
        return self.font.Superscript

    # Lower case alias for Superscript
    @property
    def superscript(self):
        return self.Superscript

    @Superscript.setter
    def Superscript(self, value):
        self.font.Superscript = value

    # Lower case alias for Superscript setter
    @superscript.setter
    def superscript(self, value):
        self.Superscript = value

    @property
    def Underline(self):
        return self.font.Underline

    # Lower case alias for Underline
    @property
    def underline(self):
        return self.Underline

    @Underline.setter
    def Underline(self, value):
        self.font.Underline = value

    # Lower case alias for Underline setter
    @underline.setter
    def underline(self, value):
        self.Underline = value


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.fonts.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.fonts.Item(*arguments)

    def Replace(self, Original=None, Replacement=None):
        arguments = com_arguments([Original, Replacement])
        self.fonts.Replace(*arguments)


class FreeformBuilder:

    def __init__(self, freeformbuilder=None):
        self.freeformbuilder = freeformbuilder

    @property
    def Application(self):
        return Application(self.freeformbuilder.Application)

    @property
    def Creator(self):
        return self.freeformbuilder.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Parent(self):
        return self.freeformbuilder.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def AddNodes(self, SegmentType=None, EditingType=None, X1=None, Y1=None, X2=None, Y2=None, X3=None, Y3=None):
        arguments = com_arguments([SegmentType, EditingType, X1, Y1, X2, Y2, X3, Y3])
        self.freeformbuilder.AddNodes(*arguments)

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

    # Lower case alias for Border
    @property
    def border(self):
        return self.Border

    @property
    def Creator(self):
        return self.gridlines.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Name(self):
        return self.gridlines.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.gridlines.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.groupshapes.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Parent(self):
        return self.groupshapes.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.groupshapes.Item(*arguments)

    def Range(self, Index=None):
        arguments = com_arguments([Index])
        return self.groupshapes.Range(*arguments)


class HeaderFooter:

    def __init__(self, headerfooter=None):
        self.headerfooter = headerfooter

    @property
    def Application(self):
        return Application(self.headerfooter.Application)

    @property
    def Format(self):
        return self.headerfooter.Format

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @Format.setter
    def Format(self, value):
        self.headerfooter.Format = value

    # Lower case alias for Format setter
    @format.setter
    def format(self, value):
        self.Format = value

    @property
    def Parent(self):
        return self.headerfooter.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Text(self):
        return self.headerfooter.Text

    # Lower case alias for Text
    @property
    def text(self):
        return self.Text

    @Text.setter
    def Text(self, value):
        self.headerfooter.Text = value

    # Lower case alias for Text setter
    @text.setter
    def text(self, value):
        self.Text = value

    @property
    def UseFormat(self):
        return self.headerfooter.UseFormat

    # Lower case alias for UseFormat
    @property
    def useformat(self):
        return self.UseFormat

    @UseFormat.setter
    def UseFormat(self, value):
        self.headerfooter.UseFormat = value

    # Lower case alias for UseFormat setter
    @useformat.setter
    def useformat(self, value):
        self.UseFormat = value

    @property
    def Visible(self):
        return self.headerfooter.Visible

    # Lower case alias for Visible
    @property
    def visible(self):
        return self.Visible

    @Visible.setter
    def Visible(self, value):
        self.headerfooter.Visible = value

    # Lower case alias for Visible setter
    @visible.setter
    def visible(self, value):
        self.Visible = value


class HeadersFooters:

    def __init__(self, headersfooters=None):
        self.headersfooters = headersfooters

    @property
    def Application(self):
        return Application(self.headersfooters.Application)

    @property
    def DateAndTime(self):
        return HeaderFooter(self.headersfooters.DateAndTime)

    # Lower case alias for DateAndTime
    @property
    def dateandtime(self):
        return self.DateAndTime

    @property
    def DisplayOnTitleSlide(self):
        return self.headersfooters.DisplayOnTitleSlide

    # Lower case alias for DisplayOnTitleSlide
    @property
    def displayontitleslide(self):
        return self.DisplayOnTitleSlide

    @DisplayOnTitleSlide.setter
    def DisplayOnTitleSlide(self, value):
        self.headersfooters.DisplayOnTitleSlide = value

    # Lower case alias for DisplayOnTitleSlide setter
    @displayontitleslide.setter
    def displayontitleslide(self, value):
        self.DisplayOnTitleSlide = value

    @property
    def Footer(self):
        return HeaderFooter(self.headersfooters.Footer)

    # Lower case alias for Footer
    @property
    def footer(self):
        return self.Footer

    @property
    def Header(self):
        return HeaderFooter(self.headersfooters.Header)

    # Lower case alias for Header
    @property
    def header(self):
        return self.Header

    @property
    def Parent(self):
        return self.headersfooters.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SlideNumber(self):
        return HeaderFooter(self.headersfooters.SlideNumber)

    # Lower case alias for SlideNumber
    @property
    def slidenumber(self):
        return self.SlideNumber

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

    # Lower case alias for Border
    @property
    def border(self):
        return self.Border

    @property
    def Creator(self):
        return self.hilolines.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.hilolines.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def Name(self):
        return self.hilolines.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.hilolines.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

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

    # Lower case alias for Address
    @property
    def address(self):
        return self.Address

    @Address.setter
    def Address(self, value):
        self.hyperlink.Address = value

    # Lower case alias for Address setter
    @address.setter
    def address(self, value):
        self.Address = value

    @property
    def Application(self):
        return Application(self.hyperlink.Application)

    @property
    def EmailSubject(self):
        return self.hyperlink.EmailSubject

    # Lower case alias for EmailSubject
    @property
    def emailsubject(self):
        return self.EmailSubject

    @EmailSubject.setter
    def EmailSubject(self, value):
        self.hyperlink.EmailSubject = value

    # Lower case alias for EmailSubject setter
    @emailsubject.setter
    def emailsubject(self, value):
        self.EmailSubject = value

    @property
    def Parent(self):
        return self.hyperlink.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ScreenTip(self):
        return self.hyperlink.ScreenTip

    # Lower case alias for ScreenTip
    @property
    def screentip(self):
        return self.ScreenTip

    @ScreenTip.setter
    def ScreenTip(self, value):
        self.hyperlink.ScreenTip = value

    # Lower case alias for ScreenTip setter
    @screentip.setter
    def screentip(self, value):
        self.ScreenTip = value

    @property
    def ShowAndReturn(self):
        return self.hyperlink.ShowAndReturn

    # Lower case alias for ShowAndReturn
    @property
    def showandreturn(self):
        return self.ShowAndReturn

    @ShowAndReturn.setter
    def ShowAndReturn(self, value):
        self.hyperlink.ShowAndReturn = value

    # Lower case alias for ShowAndReturn setter
    @showandreturn.setter
    def showandreturn(self, value):
        self.ShowAndReturn = value

    @property
    def SubAddress(self):
        return self.hyperlink.SubAddress

    # Lower case alias for SubAddress
    @property
    def subaddress(self):
        return self.SubAddress

    @SubAddress.setter
    def SubAddress(self, value):
        self.hyperlink.SubAddress = value

    # Lower case alias for SubAddress setter
    @subaddress.setter
    def subaddress(self, value):
        self.SubAddress = value

    @property
    def TextToDisplay(self):
        return self.hyperlink.TextToDisplay

    # Lower case alias for TextToDisplay
    @property
    def texttodisplay(self):
        return self.TextToDisplay

    @TextToDisplay.setter
    def TextToDisplay(self, value):
        self.hyperlink.TextToDisplay = value

    # Lower case alias for TextToDisplay setter
    @texttodisplay.setter
    def texttodisplay(self, value):
        self.TextToDisplay = value

    @property
    def Type(self):
        return self.hyperlink.Type

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    def AddToFavorites(self):
        self.hyperlink.AddToFavorites()

    def CreateNewDocument(self, FileName=None, EditNow=None, Overwrite=None):
        arguments = com_arguments([FileName, EditNow, Overwrite])
        return self.hyperlink.CreateNewDocument(*arguments)

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.hyperlinks.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.hyperlinks.Item(*arguments)


class Interior:

    def __init__(self, interior=None):
        self.interior = interior

    @property
    def Application(self):
        return self.interior.Application

    @property
    def Color(self):
        return self.interior.Color

    # Lower case alias for Color
    @property
    def color(self):
        return self.Color

    @Color.setter
    def Color(self, value):
        self.interior.Color = value

    # Lower case alias for Color setter
    @color.setter
    def color(self, value):
        self.Color = value

    @property
    def ColorIndex(self):
        return self.interior.ColorIndex

    # Lower case alias for ColorIndex
    @property
    def colorindex(self):
        return self.ColorIndex

    @ColorIndex.setter
    def ColorIndex(self, value):
        self.interior.ColorIndex = value

    # Lower case alias for ColorIndex setter
    @colorindex.setter
    def colorindex(self, value):
        self.ColorIndex = value

    @property
    def Creator(self):
        return self.interior.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def InvertIfNegative(self):
        return self.interior.InvertIfNegative

    # Lower case alias for InvertIfNegative
    @property
    def invertifnegative(self):
        return self.InvertIfNegative

    @InvertIfNegative.setter
    def InvertIfNegative(self, value):
        self.interior.InvertIfNegative = value

    # Lower case alias for InvertIfNegative setter
    @invertifnegative.setter
    def invertifnegative(self, value):
        self.InvertIfNegative = value

    @property
    def Parent(self):
        return self.interior.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Pattern(self):
        return XlPattern(self.interior.Pattern)

    # Lower case alias for Pattern
    @property
    def pattern(self):
        return self.Pattern

    @Pattern.setter
    def Pattern(self, value):
        self.interior.Pattern = value

    # Lower case alias for Pattern setter
    @pattern.setter
    def pattern(self, value):
        self.Pattern = value

    @property
    def PatternColor(self):
        return self.interior.PatternColor

    # Lower case alias for PatternColor
    @property
    def patterncolor(self):
        return self.PatternColor

    @PatternColor.setter
    def PatternColor(self, value):
        self.interior.PatternColor = value

    # Lower case alias for PatternColor setter
    @patterncolor.setter
    def patterncolor(self, value):
        self.PatternColor = value

    @property
    def PatternColorIndex(self):
        return XlColorIndex(self.interior.PatternColorIndex)

    # Lower case alias for PatternColorIndex
    @property
    def patterncolorindex(self):
        return self.PatternColorIndex

    @PatternColorIndex.setter
    def PatternColorIndex(self, value):
        self.interior.PatternColorIndex = value

    # Lower case alias for PatternColorIndex setter
    @patterncolorindex.setter
    def patterncolorindex(self, value):
        self.PatternColorIndex = value


class LeaderLines:

    def __init__(self, leaderlines=None):
        self.leaderlines = leaderlines

    @property
    def Application(self):
        return self.leaderlines.Application

    @property
    def Border(self):
        return ChartBorder(self.leaderlines.Border)

    # Lower case alias for Border
    @property
    def border(self):
        return self.Border

    @property
    def Creator(self):
        return self.leaderlines.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.leaderlines.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def Parent(self):
        return self.leaderlines.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

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

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.legend.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def Height(self):
        return self.legend.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @Height.setter
    def Height(self, value):
        self.legend.Height = value

    # Lower case alias for Height setter
    @height.setter
    def height(self, value):
        self.Height = value

    @property
    def IncludeInLayout(self):
        return self.legend.IncludeInLayout

    # Lower case alias for IncludeInLayout
    @property
    def includeinlayout(self):
        return self.IncludeInLayout

    @IncludeInLayout.setter
    def IncludeInLayout(self, value):
        self.legend.IncludeInLayout = value

    # Lower case alias for IncludeInLayout setter
    @includeinlayout.setter
    def includeinlayout(self, value):
        self.IncludeInLayout = value

    @property
    def Left(self):
        return self.legend.Left

    # Lower case alias for Left
    @property
    def left(self):
        return self.Left

    @property
    def Name(self):
        return self.legend.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.legend.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return XlLegendPosition(self.legend.Position)

    # Lower case alias for Position
    @property
    def position(self):
        return self.Position

    @Position.setter
    def Position(self, value):
        self.legend.Position = value

    # Lower case alias for Position setter
    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Shadow(self):
        return self.legend.Shadow

    # Lower case alias for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.legend.Shadow = value

    # Lower case alias for Shadow setter
    @shadow.setter
    def shadow(self, value):
        self.Shadow = value

    @property
    def Top(self):
        return self.legend.Top

    # Lower case alias for Top
    @property
    def top(self):
        return self.Top

    @Top.setter
    def Top(self, value):
        self.legend.Top = value

    # Lower case alias for Top setter
    @top.setter
    def top(self, value):
        self.Top = value

    @property
    def Width(self):
        return self.legend.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

    @Width.setter
    def Width(self, value):
        self.legend.Width = value

    # Lower case alias for Width setter
    @width.setter
    def width(self, value):
        self.Width = value

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.legendentries.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Parent(self):
        return self.legendentries.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

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

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Font(self):
        return ChartFont(self.legendentry.Font)

    # Lower case alias for Font
    @property
    def font(self):
        return self.Font

    @property
    def Format(self):
        return ChartFormat(self.legendentry.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def Height(self):
        return self.legendentry.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @property
    def Index(self):
        return self.legendentry.Index

    # Lower case alias for Index
    @property
    def index(self):
        return self.Index

    @property
    def Left(self):
        return self.legendentry.Left

    # Lower case alias for Left
    @property
    def left(self):
        return self.Left

    @property
    def LegendKey(self):
        return LegendKey(self.legendentry.LegendKey)

    # Lower case alias for LegendKey
    @property
    def legendkey(self):
        return self.LegendKey

    @property
    def Parent(self):
        return self.legendentry.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Top(self):
        return self.legendentry.Top

    # Lower case alias for Top
    @property
    def top(self):
        return self.Top

    @property
    def Width(self):
        return self.legendentry.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

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

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.legendkey.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def Height(self):
        return self.legendkey.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @property
    def InvertIfNegative(self):
        return self.legendkey.InvertIfNegative

    # Lower case alias for InvertIfNegative
    @property
    def invertifnegative(self):
        return self.InvertIfNegative

    @InvertIfNegative.setter
    def InvertIfNegative(self, value):
        self.legendkey.InvertIfNegative = value

    # Lower case alias for InvertIfNegative setter
    @invertifnegative.setter
    def invertifnegative(self, value):
        self.InvertIfNegative = value

    @property
    def Left(self):
        return self.legendkey.Left

    # Lower case alias for Left
    @property
    def left(self):
        return self.Left

    @property
    def MarkerBackgroundColor(self):
        return self.legendkey.MarkerBackgroundColor

    # Lower case alias for MarkerBackgroundColor
    @property
    def markerbackgroundcolor(self):
        return self.MarkerBackgroundColor

    @MarkerBackgroundColor.setter
    def MarkerBackgroundColor(self, value):
        self.legendkey.MarkerBackgroundColor = value

    # Lower case alias for MarkerBackgroundColor setter
    @markerbackgroundcolor.setter
    def markerbackgroundcolor(self, value):
        self.MarkerBackgroundColor = value

    @property
    def MarkerBackgroundColorIndex(self):
        return XlColorIndex(self.legendkey.MarkerBackgroundColorIndex)

    # Lower case alias for MarkerBackgroundColorIndex
    @property
    def markerbackgroundcolorindex(self):
        return self.MarkerBackgroundColorIndex

    @MarkerBackgroundColorIndex.setter
    def MarkerBackgroundColorIndex(self, value):
        self.legendkey.MarkerBackgroundColorIndex = value

    # Lower case alias for MarkerBackgroundColorIndex setter
    @markerbackgroundcolorindex.setter
    def markerbackgroundcolorindex(self, value):
        self.MarkerBackgroundColorIndex = value

    @property
    def MarkerForegroundColor(self):
        return self.legendkey.MarkerForegroundColor

    # Lower case alias for MarkerForegroundColor
    @property
    def markerforegroundcolor(self):
        return self.MarkerForegroundColor

    @MarkerForegroundColor.setter
    def MarkerForegroundColor(self, value):
        self.legendkey.MarkerForegroundColor = value

    # Lower case alias for MarkerForegroundColor setter
    @markerforegroundcolor.setter
    def markerforegroundcolor(self, value):
        self.MarkerForegroundColor = value

    @property
    def MarkerForegroundColorIndex(self):
        return XlColorIndex(self.legendkey.MarkerForegroundColorIndex)

    # Lower case alias for MarkerForegroundColorIndex
    @property
    def markerforegroundcolorindex(self):
        return self.MarkerForegroundColorIndex

    @MarkerForegroundColorIndex.setter
    def MarkerForegroundColorIndex(self, value):
        self.legendkey.MarkerForegroundColorIndex = value

    # Lower case alias for MarkerForegroundColorIndex setter
    @markerforegroundcolorindex.setter
    def markerforegroundcolorindex(self, value):
        self.MarkerForegroundColorIndex = value

    @property
    def MarkerSize(self):
        return self.legendkey.MarkerSize

    # Lower case alias for MarkerSize
    @property
    def markersize(self):
        return self.MarkerSize

    @MarkerSize.setter
    def MarkerSize(self, value):
        self.legendkey.MarkerSize = value

    # Lower case alias for MarkerSize setter
    @markersize.setter
    def markersize(self, value):
        self.MarkerSize = value

    @property
    def MarkerStyle(self):
        return XlMarkerStyle(self.legendkey.MarkerStyle)

    # Lower case alias for MarkerStyle
    @property
    def markerstyle(self):
        return self.MarkerStyle

    @MarkerStyle.setter
    def MarkerStyle(self, value):
        self.legendkey.MarkerStyle = value

    # Lower case alias for MarkerStyle setter
    @markerstyle.setter
    def markerstyle(self, value):
        self.MarkerStyle = value

    @property
    def Parent(self):
        return self.legendkey.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PictureType(self):
        return XlChartPictureType(self.legendkey.PictureType)

    # Lower case alias for PictureType
    @property
    def picturetype(self):
        return self.PictureType

    @PictureType.setter
    def PictureType(self, value):
        self.legendkey.PictureType = value

    # Lower case alias for PictureType setter
    @picturetype.setter
    def picturetype(self, value):
        self.PictureType = value

    @property
    def PictureUnit2(self):
        return self.legendkey.PictureUnit2

    # Lower case alias for PictureUnit2
    @property
    def pictureunit2(self):
        return self.PictureUnit2

    @PictureUnit2.setter
    def PictureUnit2(self, value):
        self.legendkey.PictureUnit2 = value

    # Lower case alias for PictureUnit2 setter
    @pictureunit2.setter
    def pictureunit2(self, value):
        self.PictureUnit2 = value

    @property
    def Shadow(self):
        return self.legendkey.Shadow

    # Lower case alias for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.legendkey.Shadow = value

    # Lower case alias for Shadow setter
    @shadow.setter
    def shadow(self, value):
        self.Shadow = value

    @property
    def Smooth(self):
        return self.legendkey.Smooth

    # Lower case alias for Smooth
    @property
    def smooth(self):
        return self.Smooth

    @Smooth.setter
    def Smooth(self, value):
        self.legendkey.Smooth = value

    # Lower case alias for Smooth setter
    @smooth.setter
    def smooth(self, value):
        self.Smooth = value

    @property
    def Top(self):
        return self.legendkey.Top

    # Lower case alias for Top
    @property
    def top(self):
        return self.Top

    @property
    def Width(self):
        return self.legendkey.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

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

    # Lower case alias for BackColor
    @property
    def backcolor(self):
        return self.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.lineformat.BackColor = value

    # Lower case alias for BackColor setter
    @backcolor.setter
    def backcolor(self, value):
        self.BackColor = value

    @property
    def BeginArrowheadLength(self):
        return self.lineformat.BeginArrowheadLength

    # Lower case alias for BeginArrowheadLength
    @property
    def beginarrowheadlength(self):
        return self.BeginArrowheadLength

    @BeginArrowheadLength.setter
    def BeginArrowheadLength(self, value):
        self.lineformat.BeginArrowheadLength = value

    # Lower case alias for BeginArrowheadLength setter
    @beginarrowheadlength.setter
    def beginarrowheadlength(self, value):
        self.BeginArrowheadLength = value

    @property
    def BeginArrowheadStyle(self):
        return self.lineformat.BeginArrowheadStyle

    # Lower case alias for BeginArrowheadStyle
    @property
    def beginarrowheadstyle(self):
        return self.BeginArrowheadStyle

    @BeginArrowheadStyle.setter
    def BeginArrowheadStyle(self, value):
        self.lineformat.BeginArrowheadStyle = value

    # Lower case alias for BeginArrowheadStyle setter
    @beginarrowheadstyle.setter
    def beginarrowheadstyle(self, value):
        self.BeginArrowheadStyle = value

    @property
    def BeginArrowheadWidth(self):
        return self.lineformat.BeginArrowheadWidth

    # Lower case alias for BeginArrowheadWidth
    @property
    def beginarrowheadwidth(self):
        return self.BeginArrowheadWidth

    @BeginArrowheadWidth.setter
    def BeginArrowheadWidth(self, value):
        self.lineformat.BeginArrowheadWidth = value

    # Lower case alias for BeginArrowheadWidth setter
    @beginarrowheadwidth.setter
    def beginarrowheadwidth(self, value):
        self.BeginArrowheadWidth = value

    @property
    def Creator(self):
        return self.lineformat.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def DashStyle(self):
        return self.lineformat.DashStyle

    # Lower case alias for DashStyle
    @property
    def dashstyle(self):
        return self.DashStyle

    @DashStyle.setter
    def DashStyle(self, value):
        self.lineformat.DashStyle = value

    # Lower case alias for DashStyle setter
    @dashstyle.setter
    def dashstyle(self, value):
        self.DashStyle = value

    @property
    def EndArrowheadLength(self):
        return self.lineformat.EndArrowheadLength

    # Lower case alias for EndArrowheadLength
    @property
    def endarrowheadlength(self):
        return self.EndArrowheadLength

    @EndArrowheadLength.setter
    def EndArrowheadLength(self, value):
        self.lineformat.EndArrowheadLength = value

    # Lower case alias for EndArrowheadLength setter
    @endarrowheadlength.setter
    def endarrowheadlength(self, value):
        self.EndArrowheadLength = value

    @property
    def EndArrowheadStyle(self):
        return self.lineformat.EndArrowheadStyle

    # Lower case alias for EndArrowheadStyle
    @property
    def endarrowheadstyle(self):
        return self.EndArrowheadStyle

    @EndArrowheadStyle.setter
    def EndArrowheadStyle(self, value):
        self.lineformat.EndArrowheadStyle = value

    # Lower case alias for EndArrowheadStyle setter
    @endarrowheadstyle.setter
    def endarrowheadstyle(self, value):
        self.EndArrowheadStyle = value

    @property
    def EndArrowheadWidth(self):
        return self.lineformat.EndArrowheadWidth

    # Lower case alias for EndArrowheadWidth
    @property
    def endarrowheadwidth(self):
        return self.EndArrowheadWidth

    @EndArrowheadWidth.setter
    def EndArrowheadWidth(self, value):
        self.lineformat.EndArrowheadWidth = value

    # Lower case alias for EndArrowheadWidth setter
    @endarrowheadwidth.setter
    def endarrowheadwidth(self, value):
        self.EndArrowheadWidth = value

    @property
    def ForeColor(self):
        return ColorFormat(self.lineformat.ForeColor)

    # Lower case alias for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.lineformat.ForeColor = value

    # Lower case alias for ForeColor setter
    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def InsetPen(self):
        return self.lineformat.InsetPen

    # Lower case alias for InsetPen
    @property
    def insetpen(self):
        return self.InsetPen

    @InsetPen.setter
    def InsetPen(self, value):
        self.lineformat.InsetPen = value

    # Lower case alias for InsetPen setter
    @insetpen.setter
    def insetpen(self, value):
        self.InsetPen = value

    @property
    def Parent(self):
        return self.lineformat.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Pattern(self):
        return self.lineformat.Pattern

    # Lower case alias for Pattern
    @property
    def pattern(self):
        return self.Pattern

    @Pattern.setter
    def Pattern(self, value):
        self.lineformat.Pattern = value

    # Lower case alias for Pattern setter
    @pattern.setter
    def pattern(self, value):
        self.Pattern = value

    @property
    def Style(self):
        return self.lineformat.Style

    # Lower case alias for Style
    @property
    def style(self):
        return self.Style

    @Style.setter
    def Style(self, value):
        self.lineformat.Style = value

    # Lower case alias for Style setter
    @style.setter
    def style(self, value):
        self.Style = value

    @property
    def Transparency(self):
        return self.lineformat.Transparency

    # Lower case alias for Transparency
    @property
    def transparency(self):
        return self.Transparency

    @Transparency.setter
    def Transparency(self, value):
        self.lineformat.Transparency = value

    # Lower case alias for Transparency setter
    @transparency.setter
    def transparency(self, value):
        self.Transparency = value

    @property
    def Visible(self):
        return self.lineformat.Visible

    # Lower case alias for Visible
    @property
    def visible(self):
        return self.Visible

    @Visible.setter
    def Visible(self, value):
        self.lineformat.Visible = value

    # Lower case alias for Visible setter
    @visible.setter
    def visible(self, value):
        self.Visible = value

    @property
    def Weight(self):
        return self.lineformat.Weight

    # Lower case alias for Weight
    @property
    def weight(self):
        return self.Weight

    @Weight.setter
    def Weight(self, value):
        self.lineformat.Weight = value

    # Lower case alias for Weight setter
    @weight.setter
    def weight(self, value):
        self.Weight = value


class LinkFormat:

    def __init__(self, linkformat=None):
        self.linkformat = linkformat

    @property
    def Application(self):
        return Application(self.linkformat.Application)

    @property
    def AutoUpdate(self):
        return self.linkformat.AutoUpdate

    # Lower case alias for AutoUpdate
    @property
    def autoupdate(self):
        return self.AutoUpdate

    @AutoUpdate.setter
    def AutoUpdate(self, value):
        self.linkformat.AutoUpdate = value

    # Lower case alias for AutoUpdate setter
    @autoupdate.setter
    def autoupdate(self, value):
        self.AutoUpdate = value

    @property
    def Parent(self):
        return self.linkformat.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SourceFullName(self):
        return self.linkformat.SourceFullName

    # Lower case alias for SourceFullName
    @property
    def sourcefullname(self):
        return self.SourceFullName

    @SourceFullName.setter
    def SourceFullName(self, value):
        self.linkformat.SourceFullName = value

    # Lower case alias for SourceFullName setter
    @sourcefullname.setter
    def sourcefullname(self, value):
        self.SourceFullName = value

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

    # Lower case alias for Background
    @property
    def background(self):
        return self.Background

    @property
    def BackgroundStyle(self):
        return self.master.BackgroundStyle

    # Lower case alias for BackgroundStyle
    @property
    def backgroundstyle(self):
        return self.BackgroundStyle

    @BackgroundStyle.setter
    def BackgroundStyle(self, value):
        self.master.BackgroundStyle = value

    # Lower case alias for BackgroundStyle setter
    @backgroundstyle.setter
    def backgroundstyle(self, value):
        self.BackgroundStyle = value

    @property
    def ColorScheme(self):
        return ColorScheme(self.master.ColorScheme)

    # Lower case alias for ColorScheme
    @property
    def colorscheme(self):
        return self.ColorScheme

    @ColorScheme.setter
    def ColorScheme(self, value):
        self.master.ColorScheme = value

    # Lower case alias for ColorScheme setter
    @colorscheme.setter
    def colorscheme(self, value):
        self.ColorScheme = value

    @property
    def CustomerData(self):
        return CustomerData(self.master.CustomerData)

    # Lower case alias for CustomerData
    @property
    def customerdata(self):
        return self.CustomerData

    @property
    def CustomLayouts(self):
        return CustomLayouts(self.master.CustomLayouts)

    # Lower case alias for CustomLayouts
    @property
    def customlayouts(self):
        return self.CustomLayouts

    @property
    def Design(self):
        return Design(self.master.Design)

    # Lower case alias for Design
    @property
    def design(self):
        return self.Design

    @property
    def HeadersFooters(self):
        return HeadersFooters(self.master.HeadersFooters)

    # Lower case alias for HeadersFooters
    @property
    def headersfooters(self):
        return self.HeadersFooters

    @property
    def Height(self):
        return self.master.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @Height.setter
    def Height(self, value):
        self.master.Height = value

    # Lower case alias for Height setter
    @height.setter
    def height(self, value):
        self.Height = value

    @property
    def Hyperlinks(self):
        return Hyperlinks(self.master.Hyperlinks)

    # Lower case alias for Hyperlinks
    @property
    def hyperlinks(self):
        return self.Hyperlinks

    @property
    def Name(self):
        return self.master.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.master.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.master.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Shapes(self):
        return Shapes(self.master.Shapes)

    # Lower case alias for Shapes
    @property
    def shapes(self):
        return self.Shapes

    @property
    def SlideShowTransition(self):
        return SlideShowTransition(self.master.SlideShowTransition)

    # Lower case alias for SlideShowTransition
    @property
    def slideshowtransition(self):
        return self.SlideShowTransition

    @property
    def TextStyles(self):
        return TextStyles(self.master.TextStyles)

    # Lower case alias for TextStyles
    @property
    def textstyles(self):
        return self.TextStyles

    @property
    def Theme(self):
        return self.master.Theme

    # Lower case alias for Theme
    @property
    def theme(self):
        return self.Theme

    @property
    def TimeLine(self):
        return TimeLine(self.master.TimeLine)

    # Lower case alias for TimeLine
    @property
    def timeline(self):
        return self.TimeLine

    @property
    def Width(self):
        return self.master.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

    def ApplyTheme(self, themeName=None):
        arguments = com_arguments([themeName])
        self.master.ApplyTheme(*arguments)

    def Delete(self):
        self.master.Delete()


class MediaBookmark:

    def __init__(self, mediabookmark=None):
        self.mediabookmark = mediabookmark

    @property
    def Index(self):
        return self.mediabookmark.Index

    # Lower case alias for Index
    @property
    def index(self):
        return self.Index

    @property
    def Name(self):
        return self.mediabookmark.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Position(self):
        return self.mediabookmark.Position

    # Lower case alias for Position
    @property
    def position(self):
        return self.Position

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    def Add(self, Position=None, Name=None):
        arguments = com_arguments([Position, Name])
        return MediaBookmark(self.mediabookmarks.Add(*arguments))

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.mediabookmarks.Item(*arguments)


class MediaFormat:

    def __init__(self, mediaformat=None):
        self.mediaformat = mediaformat

    @property
    def Application(self):
        return Application(self.mediaformat.Application)

    @property
    def AudioCompressionType(self):
        return self.mediaformat.AudioCompressionType

    # Lower case alias for AudioCompressionType
    @property
    def audiocompressiontype(self):
        return self.AudioCompressionType

    @property
    def AudioSamplingRate(self):
        return self.mediaformat.AudioSamplingRate

    # Lower case alias for AudioSamplingRate
    @property
    def audiosamplingrate(self):
        return self.AudioSamplingRate

    @property
    def EndPoint(self):
        return self.mediaformat.EndPoint

    # Lower case alias for EndPoint
    @property
    def endpoint(self):
        return self.EndPoint

    @EndPoint.setter
    def EndPoint(self, value):
        self.mediaformat.EndPoint = value

    # Lower case alias for EndPoint setter
    @endpoint.setter
    def endpoint(self, value):
        self.EndPoint = value

    @property
    def FadeInDuration(self):
        return self.mediaformat.FadeInDuration

    # Lower case alias for FadeInDuration
    @property
    def fadeinduration(self):
        return self.FadeInDuration

    @FadeInDuration.setter
    def FadeInDuration(self, value):
        self.mediaformat.FadeInDuration = value

    # Lower case alias for FadeInDuration setter
    @fadeinduration.setter
    def fadeinduration(self, value):
        self.FadeInDuration = value

    @property
    def FadeOutDuration(self):
        return self.mediaformat.FadeOutDuration

    # Lower case alias for FadeOutDuration
    @property
    def fadeoutduration(self):
        return self.FadeOutDuration

    @FadeOutDuration.setter
    def FadeOutDuration(self, value):
        self.mediaformat.FadeOutDuration = value

    # Lower case alias for FadeOutDuration setter
    @fadeoutduration.setter
    def fadeoutduration(self, value):
        self.FadeOutDuration = value

    @property
    def IsEmbedded(self):
        return self.mediaformat.IsEmbedded

    # Lower case alias for IsEmbedded
    @property
    def isembedded(self):
        return self.IsEmbedded

    @property
    def IsLinked(self):
        return self.mediaformat.IsLinked

    # Lower case alias for IsLinked
    @property
    def islinked(self):
        return self.IsLinked

    @property
    def Length(self):
        return self.mediaformat.Length

    # Lower case alias for Length
    @property
    def length(self):
        return self.Length

    @property
    def MediaBookmarks(self):
        return MediaBookmarks(self.mediaformat.MediaBookmarks)

    # Lower case alias for MediaBookmarks
    @property
    def mediabookmarks(self):
        return self.MediaBookmarks

    @property
    def Muted(self):
        return self.mediaformat.Muted

    # Lower case alias for Muted
    @property
    def muted(self):
        return self.Muted

    @Muted.setter
    def Muted(self, value):
        self.mediaformat.Muted = value

    # Lower case alias for Muted setter
    @muted.setter
    def muted(self, value):
        self.Muted = value

    @property
    def Parent(self):
        return self.mediaformat.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ResamplingStatus(self):
        return self.mediaformat.ResamplingStatus

    # Lower case alias for ResamplingStatus
    @property
    def resamplingstatus(self):
        return self.ResamplingStatus

    @property
    def SampleHeight(self):
        return self.mediaformat.SampleHeight

    # Lower case alias for SampleHeight
    @property
    def sampleheight(self):
        return self.SampleHeight

    @property
    def SampleWidth(self):
        return self.mediaformat.SampleWidth

    # Lower case alias for SampleWidth
    @property
    def samplewidth(self):
        return self.SampleWidth

    @property
    def StartPoint(self):
        return self.mediaformat.StartPoint

    # Lower case alias for StartPoint
    @property
    def startpoint(self):
        return self.StartPoint

    @StartPoint.setter
    def StartPoint(self, value):
        self.mediaformat.StartPoint = value

    # Lower case alias for StartPoint setter
    @startpoint.setter
    def startpoint(self, value):
        self.StartPoint = value

    @property
    def VideoCompressionType(self):
        return self.mediaformat.VideoCompressionType

    # Lower case alias for VideoCompressionType
    @property
    def videocompressiontype(self):
        return self.VideoCompressionType

    @property
    def VideoFrameRate(self):
        return self.mediaformat.VideoFrameRate

    # Lower case alias for VideoFrameRate
    @property
    def videoframerate(self):
        return self.VideoFrameRate

    @property
    def Volume(self):
        return self.mediaformat.Volume

    # Lower case alias for Volume
    @property
    def volume(self):
        return self.Volume

    @Volume.setter
    def Volume(self, value):
        self.mediaformat.Volume = value

    # Lower case alias for Volume setter
    @volume.setter
    def volume(self, value):
        self.Volume = value

    def Resample(self, Trim=None, SampleHeight=None, SampleWidth=None, VideoFrameRate=None, AudioSamplingRate=None, VideoBitRate=None):
        arguments = com_arguments([Trim, SampleHeight, SampleWidth, VideoFrameRate, AudioSamplingRate, VideoBitRate])
        return self.mediaformat.Resample(*arguments)

    def ResampleFromProfile(self, profile=None):
        arguments = com_arguments([profile])
        return self.mediaformat.ResampleFromProfile(*arguments)

    def SetDisplayPicture(self, Position=None):
        arguments = com_arguments([Position])
        return self.mediaformat.SetDisplayPicture(*arguments)

    def SetDisplayPictureFromFile(self, FilePath=None):
        arguments = com_arguments([FilePath])
        return self.mediaformat.SetDisplayPictureFromFile(*arguments)


class Model3DFormat:

    def __init__(self, model3dformat=None):
        self.model3dformat = model3dformat

    @property
    def Application(self):
        return Application(self.model3dformat.Application)

    @property
    def AutoFit(self):
        return self.model3dformat.AutoFit

    # Lower case alias for AutoFit
    @property
    def autofit(self):
        return self.AutoFit

    @AutoFit.setter
    def AutoFit(self, value):
        self.model3dformat.AutoFit = value

    # Lower case alias for AutoFit setter
    @autofit.setter
    def autofit(self, value):
        self.AutoFit = value

    @property
    def CameraPositionX(self):
        return self.model3dformat.CameraPositionX

    # Lower case alias for CameraPositionX
    @property
    def camerapositionx(self):
        return self.CameraPositionX

    @CameraPositionX.setter
    def CameraPositionX(self, value):
        self.model3dformat.CameraPositionX = value

    # Lower case alias for CameraPositionX setter
    @camerapositionx.setter
    def camerapositionx(self, value):
        self.CameraPositionX = value

    @property
    def CameraPositionY(self):
        return self.model3dformat.CameraPositionY

    # Lower case alias for CameraPositionY
    @property
    def camerapositiony(self):
        return self.CameraPositionY

    @CameraPositionY.setter
    def CameraPositionY(self, value):
        self.model3dformat.CameraPositionY = value

    # Lower case alias for CameraPositionY setter
    @camerapositiony.setter
    def camerapositiony(self, value):
        self.CameraPositionY = value

    @property
    def CameraPositionZ(self):
        return self.model3dformat.CameraPositionZ

    # Lower case alias for CameraPositionZ
    @property
    def camerapositionz(self):
        return self.CameraPositionZ

    @CameraPositionZ.setter
    def CameraPositionZ(self, value):
        self.model3dformat.CameraPositionZ = value

    # Lower case alias for CameraPositionZ setter
    @camerapositionz.setter
    def camerapositionz(self, value):
        self.CameraPositionZ = value

    @property
    def Creator(self):
        return self.model3dformat.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def FieldOfView(self):
        return self.model3dformat.FieldOfView

    # Lower case alias for FieldOfView
    @property
    def fieldofview(self):
        return self.FieldOfView

    @FieldOfView.setter
    def FieldOfView(self, value):
        self.model3dformat.FieldOfView = value

    # Lower case alias for FieldOfView setter
    @fieldofview.setter
    def fieldofview(self, value):
        self.FieldOfView = value

    @property
    def LookAtPointX(self):
        return self.model3dformat.LookAtPointX

    # Lower case alias for LookAtPointX
    @property
    def lookatpointx(self):
        return self.LookAtPointX

    @LookAtPointX.setter
    def LookAtPointX(self, value):
        self.model3dformat.LookAtPointX = value

    # Lower case alias for LookAtPointX setter
    @lookatpointx.setter
    def lookatpointx(self, value):
        self.LookAtPointX = value

    @property
    def LookAtPointY(self):
        return self.model3dformat.LookAtPointY

    # Lower case alias for LookAtPointY
    @property
    def lookatpointy(self):
        return self.LookAtPointY

    @LookAtPointY.setter
    def LookAtPointY(self, value):
        self.model3dformat.LookAtPointY = value

    # Lower case alias for LookAtPointY setter
    @lookatpointy.setter
    def lookatpointy(self, value):
        self.LookAtPointY = value

    @property
    def LookAtPointZ(self):
        return self.model3dformat.LookAtPointZ

    # Lower case alias for LookAtPointZ
    @property
    def lookatpointz(self):
        return self.LookAtPointZ

    @LookAtPointZ.setter
    def LookAtPointZ(self, value):
        self.model3dformat.LookAtPointZ = value

    # Lower case alias for LookAtPointZ setter
    @lookatpointz.setter
    def lookatpointz(self, value):
        self.LookAtPointZ = value

    @property
    def Parent(self):
        return self.model3dformat.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RotationX(self):
        return self.model3dformat.RotationX

    # Lower case alias for RotationX
    @property
    def rotationx(self):
        return self.RotationX

    @RotationX.setter
    def RotationX(self, value):
        self.model3dformat.RotationX = value

    # Lower case alias for RotationX setter
    @rotationx.setter
    def rotationx(self, value):
        self.RotationX = value

    @property
    def RotationY(self):
        return self.model3dformat.RotationY

    # Lower case alias for RotationY
    @property
    def rotationy(self):
        return self.RotationY

    @RotationY.setter
    def RotationY(self, value):
        self.model3dformat.RotationY = value

    # Lower case alias for RotationY setter
    @rotationy.setter
    def rotationy(self, value):
        self.RotationY = value

    @property
    def RotationZ(self):
        return self.model3dformat.RotationZ

    # Lower case alias for RotationZ
    @property
    def rotationz(self):
        return self.RotationZ

    @RotationZ.setter
    def RotationZ(self, value):
        self.model3dformat.RotationZ = value

    # Lower case alias for RotationZ setter
    @rotationz.setter
    def rotationz(self, value):
        self.RotationZ = value

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


class MotionEffect:

    def __init__(self, motioneffect=None):
        self.motioneffect = motioneffect

    @property
    def Application(self):
        return Application(self.motioneffect.Application)

    @property
    def ByX(self):
        return self.motioneffect.ByX

    # Lower case alias for ByX
    @property
    def byx(self):
        return self.ByX

    @ByX.setter
    def ByX(self, value):
        self.motioneffect.ByX = value

    # Lower case alias for ByX setter
    @byx.setter
    def byx(self, value):
        self.ByX = value

    @property
    def ByY(self):
        return self.motioneffect.ByY

    # Lower case alias for ByY
    @property
    def byy(self):
        return self.ByY

    @ByY.setter
    def ByY(self, value):
        self.motioneffect.ByY = value

    # Lower case alias for ByY setter
    @byy.setter
    def byy(self, value):
        self.ByY = value

    @property
    def FromX(self):
        return self.motioneffect.FromX

    # Lower case alias for FromX
    @property
    def fromx(self):
        return self.FromX

    @FromX.setter
    def FromX(self, value):
        self.motioneffect.FromX = value

    # Lower case alias for FromX setter
    @fromx.setter
    def fromx(self, value):
        self.FromX = value

    @property
    def FromY(self):
        return MotionEffect(self.motioneffect.FromY)

    # Lower case alias for FromY
    @property
    def fromy(self):
        return self.FromY

    @FromY.setter
    def FromY(self, value):
        self.motioneffect.FromY = value

    # Lower case alias for FromY setter
    @fromy.setter
    def fromy(self, value):
        self.FromY = value

    @property
    def Parent(self):
        return self.motioneffect.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Path(self):
        return self.motioneffect.Path

    # Lower case alias for Path
    @property
    def path(self):
        return self.Path

    @Path.setter
    def Path(self, value):
        self.motioneffect.Path = value

    # Lower case alias for Path setter
    @path.setter
    def path(self, value):
        self.Path = value

    @property
    def ToX(self):
        return self.motioneffect.ToX

    # Lower case alias for ToX
    @property
    def tox(self):
        return self.ToX

    @ToX.setter
    def ToX(self, value):
        self.motioneffect.ToX = value

    # Lower case alias for ToX setter
    @tox.setter
    def tox(self, value):
        self.ToX = value

    @property
    def ToY(self):
        return MotionEffect(self.motioneffect.ToY)

    # Lower case alias for ToY
    @property
    def toy(self):
        return self.ToY

    @ToY.setter
    def ToY(self, value):
        self.motioneffect.ToY = value

    # Lower case alias for ToY setter
    @toy.setter
    def toy(self, value):
        self.ToY = value


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Name(self):
        return self.namedslideshow.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.namedslideshow.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SlideIDs(self):
        return self.namedslideshow.SlideIDs

    # Lower case alias for SlideIDs
    @property
    def slideids(self):
        return self.SlideIDs

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.namedslideshows.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Add(self, Name=None, SafeArrayOfSlideIDs=None):
        arguments = com_arguments([Name, SafeArrayOfSlideIDs])
        return NamedSlideShow(self.namedslideshows.Add(*arguments))

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.namedslideshows.Item(*arguments)


class ObjectVerbs:

    def __init__(self, objectverbs=None):
        self.objectverbs = objectverbs

    @property
    def Application(self):
        return Application(self.objectverbs.Application)

    @property
    def Count(self):
        return self.objectverbs.Count

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.objectverbs.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.objectverbs.Item(*arguments)


class OLEFormat:

    def __init__(self, oleformat=None):
        self.oleformat = oleformat

    @property
    def Application(self):
        return Application(self.oleformat.Application)

    @property
    def FollowColors(self):
        return self.oleformat.FollowColors

    # Lower case alias for FollowColors
    @property
    def followcolors(self):
        return self.FollowColors

    @FollowColors.setter
    def FollowColors(self, value):
        self.oleformat.FollowColors = value

    # Lower case alias for FollowColors setter
    @followcolors.setter
    def followcolors(self, value):
        self.FollowColors = value

    @property
    def Object(self):
        return self.oleformat.Object

    # Lower case alias for Object
    @property
    def object(self):
        return self.Object

    @property
    def ObjectVerbs(self):
        return ObjectVerbs(self.oleformat.ObjectVerbs)

    # Lower case alias for ObjectVerbs
    @property
    def objectverbs(self):
        return self.ObjectVerbs

    @property
    def Parent(self):
        return self.oleformat.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ProgID(self):
        return self.oleformat.ProgID

    # Lower case alias for ProgID
    @property
    def progid(self):
        return self.ProgID

    def Activate(self):
        self.oleformat.Activate()

    def DoVerb(self, Index=None):
        arguments = com_arguments([Index])
        self.oleformat.DoVerb(*arguments)


class Options:

    def __init__(self, options=None):
        self.options = options

    @property
    def DisplayPasteOptions(self):
        return self.options.DisplayPasteOptions

    # Lower case alias for DisplayPasteOptions
    @property
    def displaypasteoptions(self):
        return self.DisplayPasteOptions

    @DisplayPasteOptions.setter
    def DisplayPasteOptions(self, value):
        self.options.DisplayPasteOptions = value

    # Lower case alias for DisplayPasteOptions setter
    @displaypasteoptions.setter
    def displaypasteoptions(self, value):
        self.DisplayPasteOptions = value

    @property
    def ShowCoauthoringMergeChanges(self):
        return self.options.ShowCoauthoringMergeChanges

    # Lower case alias for ShowCoauthoringMergeChanges
    @property
    def showcoauthoringmergechanges(self):
        return self.ShowCoauthoringMergeChanges


class PageSetup:

    def __init__(self, pagesetup=None):
        self.pagesetup = pagesetup

    @property
    def Application(self):
        return Application(self.pagesetup.Application)

    @property
    def FirstSlideNumber(self):
        return self.pagesetup.FirstSlideNumber

    # Lower case alias for FirstSlideNumber
    @property
    def firstslidenumber(self):
        return self.FirstSlideNumber

    @FirstSlideNumber.setter
    def FirstSlideNumber(self, value):
        self.pagesetup.FirstSlideNumber = value

    # Lower case alias for FirstSlideNumber setter
    @firstslidenumber.setter
    def firstslidenumber(self, value):
        self.FirstSlideNumber = value

    @property
    def NotesOrientation(self):
        return self.pagesetup.NotesOrientation

    # Lower case alias for NotesOrientation
    @property
    def notesorientation(self):
        return self.NotesOrientation

    @NotesOrientation.setter
    def NotesOrientation(self, value):
        self.pagesetup.NotesOrientation = value

    # Lower case alias for NotesOrientation setter
    @notesorientation.setter
    def notesorientation(self, value):
        self.NotesOrientation = value

    @property
    def Parent(self):
        return self.pagesetup.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SlideHeight(self):
        return self.pagesetup.SlideHeight

    # Lower case alias for SlideHeight
    @property
    def slideheight(self):
        return self.SlideHeight

    @SlideHeight.setter
    def SlideHeight(self, value):
        self.pagesetup.SlideHeight = value

    # Lower case alias for SlideHeight setter
    @slideheight.setter
    def slideheight(self, value):
        self.SlideHeight = value

    @property
    def SlideOrientation(self):
        return self.pagesetup.SlideOrientation

    # Lower case alias for SlideOrientation
    @property
    def slideorientation(self):
        return self.SlideOrientation

    @SlideOrientation.setter
    def SlideOrientation(self, value):
        self.pagesetup.SlideOrientation = value

    # Lower case alias for SlideOrientation setter
    @slideorientation.setter
    def slideorientation(self, value):
        self.SlideOrientation = value

    @property
    def SlideSize(self):
        return self.pagesetup.SlideSize

    # Lower case alias for SlideSize
    @property
    def slidesize(self):
        return self.SlideSize

    @SlideSize.setter
    def SlideSize(self, value):
        self.pagesetup.SlideSize = value

    # Lower case alias for SlideSize setter
    @slidesize.setter
    def slidesize(self, value):
        self.SlideSize = value

    @property
    def SlideWidth(self):
        return self.pagesetup.SlideWidth

    # Lower case alias for SlideWidth
    @property
    def slidewidth(self):
        return self.SlideWidth

    @SlideWidth.setter
    def SlideWidth(self, value):
        self.pagesetup.SlideWidth = value

    # Lower case alias for SlideWidth setter
    @slidewidth.setter
    def slidewidth(self, value):
        self.SlideWidth = value


class Pane:

    def __init__(self, pane=None):
        self.pane = pane

    @property
    def Active(self):
        return self.pane.Active

    # Lower case alias for Active
    @property
    def active(self):
        return self.Active

    @property
    def Application(self):
        return Application(self.pane.Application)

    @property
    def Parent(self):
        return self.pane.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ViewType(self):
        return self.pane.ViewType

    # Lower case alias for ViewType
    @property
    def viewtype(self):
        return self.ViewType

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

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.panes.Item(*arguments)


class ParagraphFormat:

    def __init__(self, paragraphformat=None):
        self.paragraphformat = paragraphformat

    @property
    def Alignment(self):
        return self.paragraphformat.Alignment

    # Lower case alias for Alignment
    @property
    def alignment(self):
        return self.Alignment

    @Alignment.setter
    def Alignment(self, value):
        self.paragraphformat.Alignment = value

    # Lower case alias for Alignment setter
    @alignment.setter
    def alignment(self, value):
        self.Alignment = value

    @property
    def Application(self):
        return Application(self.paragraphformat.Application)

    @property
    def BaseLineAlignment(self):
        return self.paragraphformat.BaseLineAlignment

    # Lower case alias for BaseLineAlignment
    @property
    def baselinealignment(self):
        return self.BaseLineAlignment

    @BaseLineAlignment.setter
    def BaseLineAlignment(self, value):
        self.paragraphformat.BaseLineAlignment = value

    # Lower case alias for BaseLineAlignment setter
    @baselinealignment.setter
    def baselinealignment(self, value):
        self.BaseLineAlignment = value

    @property
    def Bullet(self):
        return BulletFormat(self.paragraphformat.Bullet)

    # Lower case alias for Bullet
    @property
    def bullet(self):
        return self.Bullet

    @property
    def FarEastLineBreakControl(self):
        return self.paragraphformat.FarEastLineBreakControl

    # Lower case alias for FarEastLineBreakControl
    @property
    def fareastlinebreakcontrol(self):
        return self.FarEastLineBreakControl

    @FarEastLineBreakControl.setter
    def FarEastLineBreakControl(self, value):
        self.paragraphformat.FarEastLineBreakControl = value

    # Lower case alias for FarEastLineBreakControl setter
    @fareastlinebreakcontrol.setter
    def fareastlinebreakcontrol(self, value):
        self.FarEastLineBreakControl = value

    @property
    def HangingPunctuation(self):
        return self.paragraphformat.HangingPunctuation

    # Lower case alias for HangingPunctuation
    @property
    def hangingpunctuation(self):
        return self.HangingPunctuation

    @HangingPunctuation.setter
    def HangingPunctuation(self, value):
        self.paragraphformat.HangingPunctuation = value

    # Lower case alias for HangingPunctuation setter
    @hangingpunctuation.setter
    def hangingpunctuation(self, value):
        self.HangingPunctuation = value

    @property
    def LineRuleAfter(self):
        return self.paragraphformat.LineRuleAfter

    # Lower case alias for LineRuleAfter
    @property
    def lineruleafter(self):
        return self.LineRuleAfter

    @LineRuleAfter.setter
    def LineRuleAfter(self, value):
        self.paragraphformat.LineRuleAfter = value

    # Lower case alias for LineRuleAfter setter
    @lineruleafter.setter
    def lineruleafter(self, value):
        self.LineRuleAfter = value

    @property
    def LineRuleBefore(self):
        return self.paragraphformat.LineRuleBefore

    # Lower case alias for LineRuleBefore
    @property
    def linerulebefore(self):
        return self.LineRuleBefore

    @LineRuleBefore.setter
    def LineRuleBefore(self, value):
        self.paragraphformat.LineRuleBefore = value

    # Lower case alias for LineRuleBefore setter
    @linerulebefore.setter
    def linerulebefore(self, value):
        self.LineRuleBefore = value

    @property
    def LineRuleWithin(self):
        return self.paragraphformat.LineRuleWithin

    # Lower case alias for LineRuleWithin
    @property
    def linerulewithin(self):
        return self.LineRuleWithin

    @LineRuleWithin.setter
    def LineRuleWithin(self, value):
        self.paragraphformat.LineRuleWithin = value

    # Lower case alias for LineRuleWithin setter
    @linerulewithin.setter
    def linerulewithin(self, value):
        self.LineRuleWithin = value

    @property
    def Parent(self):
        return self.paragraphformat.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SpaceAfter(self):
        return self.paragraphformat.SpaceAfter

    # Lower case alias for SpaceAfter
    @property
    def spaceafter(self):
        return self.SpaceAfter

    @SpaceAfter.setter
    def SpaceAfter(self, value):
        self.paragraphformat.SpaceAfter = value

    # Lower case alias for SpaceAfter setter
    @spaceafter.setter
    def spaceafter(self, value):
        self.SpaceAfter = value

    @property
    def SpaceBefore(self):
        return self.paragraphformat.SpaceBefore

    # Lower case alias for SpaceBefore
    @property
    def spacebefore(self):
        return self.SpaceBefore

    @SpaceBefore.setter
    def SpaceBefore(self, value):
        self.paragraphformat.SpaceBefore = value

    # Lower case alias for SpaceBefore setter
    @spacebefore.setter
    def spacebefore(self, value):
        self.SpaceBefore = value

    @property
    def SpaceWithin(self):
        return self.paragraphformat.SpaceWithin

    # Lower case alias for SpaceWithin
    @property
    def spacewithin(self):
        return self.SpaceWithin

    @SpaceWithin.setter
    def SpaceWithin(self, value):
        self.paragraphformat.SpaceWithin = value

    # Lower case alias for SpaceWithin setter
    @spacewithin.setter
    def spacewithin(self, value):
        self.SpaceWithin = value

    @property
    def TextDirection(self):
        return self.paragraphformat.TextDirection

    # Lower case alias for TextDirection
    @property
    def textdirection(self):
        return self.TextDirection

    @TextDirection.setter
    def TextDirection(self, value):
        self.paragraphformat.TextDirection = value

    # Lower case alias for TextDirection setter
    @textdirection.setter
    def textdirection(self, value):
        self.TextDirection = value

    @property
    def WordWrap(self):
        return self.paragraphformat.WordWrap

    # Lower case alias for WordWrap
    @property
    def wordwrap(self):
        return self.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.paragraphformat.WordWrap = value

    # Lower case alias for WordWrap setter
    @wordwrap.setter
    def wordwrap(self, value):
        self.WordWrap = value


class PictureFormat:

    def __init__(self, pictureformat=None):
        self.pictureformat = pictureformat

    @property
    def Application(self):
        return Application(self.pictureformat.Application)

    @property
    def Brightness(self):
        return self.pictureformat.Brightness

    # Lower case alias for Brightness
    @property
    def brightness(self):
        return self.Brightness

    @Brightness.setter
    def Brightness(self, value):
        self.pictureformat.Brightness = value

    # Lower case alias for Brightness setter
    @brightness.setter
    def brightness(self, value):
        self.Brightness = value

    @property
    def ColorType(self):
        return self.pictureformat.ColorType

    # Lower case alias for ColorType
    @property
    def colortype(self):
        return self.ColorType

    @ColorType.setter
    def ColorType(self, value):
        self.pictureformat.ColorType = value

    # Lower case alias for ColorType setter
    @colortype.setter
    def colortype(self, value):
        self.ColorType = value

    @property
    def Contrast(self):
        return self.pictureformat.Contrast

    # Lower case alias for Contrast
    @property
    def contrast(self):
        return self.Contrast

    @Contrast.setter
    def Contrast(self, value):
        self.pictureformat.Contrast = value

    # Lower case alias for Contrast setter
    @contrast.setter
    def contrast(self, value):
        self.Contrast = value

    @property
    def Creator(self):
        return self.pictureformat.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Crop(self):
        return self.pictureformat.Crop

    # Lower case alias for Crop
    @property
    def crop(self):
        return self.Crop

    @Crop.setter
    def Crop(self, value):
        self.pictureformat.Crop = value

    # Lower case alias for Crop setter
    @crop.setter
    def crop(self, value):
        self.Crop = value

    @property
    def CropBottom(self):
        return self.pictureformat.CropBottom

    # Lower case alias for CropBottom
    @property
    def cropbottom(self):
        return self.CropBottom

    @CropBottom.setter
    def CropBottom(self, value):
        self.pictureformat.CropBottom = value

    # Lower case alias for CropBottom setter
    @cropbottom.setter
    def cropbottom(self, value):
        self.CropBottom = value

    @property
    def CropLeft(self):
        return self.pictureformat.CropLeft

    # Lower case alias for CropLeft
    @property
    def cropleft(self):
        return self.CropLeft

    @CropLeft.setter
    def CropLeft(self, value):
        self.pictureformat.CropLeft = value

    # Lower case alias for CropLeft setter
    @cropleft.setter
    def cropleft(self, value):
        self.CropLeft = value

    @property
    def CropRight(self):
        return self.pictureformat.CropRight

    # Lower case alias for CropRight
    @property
    def cropright(self):
        return self.CropRight

    @CropRight.setter
    def CropRight(self, value):
        self.pictureformat.CropRight = value

    # Lower case alias for CropRight setter
    @cropright.setter
    def cropright(self, value):
        self.CropRight = value

    @property
    def CropTop(self):
        return self.pictureformat.CropTop

    # Lower case alias for CropTop
    @property
    def croptop(self):
        return self.CropTop

    @CropTop.setter
    def CropTop(self, value):
        self.pictureformat.CropTop = value

    # Lower case alias for CropTop setter
    @croptop.setter
    def croptop(self, value):
        self.CropTop = value

    @property
    def Parent(self):
        return self.pictureformat.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def TransparencyColor(self):
        return self.pictureformat.TransparencyColor

    # Lower case alias for TransparencyColor
    @property
    def transparencycolor(self):
        return self.TransparencyColor

    @TransparencyColor.setter
    def TransparencyColor(self, value):
        self.pictureformat.TransparencyColor = value

    # Lower case alias for TransparencyColor setter
    @transparencycolor.setter
    def transparencycolor(self, value):
        self.TransparencyColor = value

    @property
    def TransparentBackground(self):
        return self.pictureformat.TransparentBackground

    # Lower case alias for TransparentBackground
    @property
    def transparentbackground(self):
        return self.TransparentBackground

    @TransparentBackground.setter
    def TransparentBackground(self, value):
        self.pictureformat.TransparentBackground = value

    # Lower case alias for TransparentBackground setter
    @transparentbackground.setter
    def transparentbackground(self, value):
        self.TransparentBackground = value

    def IncrementBrightness(self, Increment=None):
        arguments = com_arguments([Increment])
        self.pictureformat.IncrementBrightness(*arguments)

    def IncrementContrast(self, Increment=None):
        arguments = com_arguments([Increment])
        self.pictureformat.IncrementContrast(*arguments)


class PlaceholderFormat:

    def __init__(self, placeholderformat=None):
        self.placeholderformat = placeholderformat

    @property
    def Application(self):
        return Application(self.placeholderformat.Application)

    @property
    def ContainedType(self):
        return self.placeholderformat.ContainedType

    # Lower case alias for ContainedType
    @property
    def containedtype(self):
        return self.ContainedType

    @property
    def Name(self):
        return self.placeholderformat.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.placeholderformat.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.placeholderformat.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Type(self):
        return self.placeholderformat.Type

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.placeholders.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def FindByName(self, Index=None):
        arguments = com_arguments([Index])
        return self.placeholders.FindByName(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.placeholders.Item(*arguments)


class Player:

    def __init__(self, player=None):
        self.player = player

    @property
    def Application(self):
        return Application(self.player.Application)

    @property
    def CurrentPosition(self):
        return self.player.CurrentPosition

    # Lower case alias for CurrentPosition
    @property
    def currentposition(self):
        return self.CurrentPosition

    @CurrentPosition.setter
    def CurrentPosition(self, value):
        self.player.CurrentPosition = value

    # Lower case alias for CurrentPosition setter
    @currentposition.setter
    def currentposition(self, value):
        self.CurrentPosition = value

    @property
    def Parent(self):
        return self.player.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def State(self):
        return self.player.State

    # Lower case alias for State
    @property
    def state(self):
        return self.State

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

    # Lower case alias for ActionVerb
    @property
    def actionverb(self):
        return self.ActionVerb

    @ActionVerb.setter
    def ActionVerb(self, value):
        self.playsettings.ActionVerb = value

    # Lower case alias for ActionVerb setter
    @actionverb.setter
    def actionverb(self, value):
        self.ActionVerb = value

    @property
    def Application(self):
        return Application(self.playsettings.Application)

    @property
    def HideWhileNotPlaying(self):
        return self.playsettings.HideWhileNotPlaying

    # Lower case alias for HideWhileNotPlaying
    @property
    def hidewhilenotplaying(self):
        return self.HideWhileNotPlaying

    @HideWhileNotPlaying.setter
    def HideWhileNotPlaying(self, value):
        self.playsettings.HideWhileNotPlaying = value

    # Lower case alias for HideWhileNotPlaying setter
    @hidewhilenotplaying.setter
    def hidewhilenotplaying(self, value):
        self.HideWhileNotPlaying = value

    @property
    def LoopUntilStopped(self):
        return self.playsettings.LoopUntilStopped

    # Lower case alias for LoopUntilStopped
    @property
    def loopuntilstopped(self):
        return self.LoopUntilStopped

    @LoopUntilStopped.setter
    def LoopUntilStopped(self, value):
        self.playsettings.LoopUntilStopped = value

    # Lower case alias for LoopUntilStopped setter
    @loopuntilstopped.setter
    def loopuntilstopped(self, value):
        self.LoopUntilStopped = value

    @property
    def Parent(self):
        return self.playsettings.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PauseAnimation(self):
        return self.playsettings.PauseAnimation

    # Lower case alias for PauseAnimation
    @property
    def pauseanimation(self):
        return self.PauseAnimation

    @PauseAnimation.setter
    def PauseAnimation(self, value):
        self.playsettings.PauseAnimation = value

    # Lower case alias for PauseAnimation setter
    @pauseanimation.setter
    def pauseanimation(self, value):
        self.PauseAnimation = value

    @property
    def PlayOnEntry(self):
        return self.playsettings.PlayOnEntry

    # Lower case alias for PlayOnEntry
    @property
    def playonentry(self):
        return self.PlayOnEntry

    @PlayOnEntry.setter
    def PlayOnEntry(self, value):
        self.playsettings.PlayOnEntry = value

    # Lower case alias for PlayOnEntry setter
    @playonentry.setter
    def playonentry(self, value):
        self.PlayOnEntry = value

    @property
    def RewindMovie(self):
        return self.playsettings.RewindMovie

    # Lower case alias for RewindMovie
    @property
    def rewindmovie(self):
        return self.RewindMovie

    @RewindMovie.setter
    def RewindMovie(self, value):
        self.playsettings.RewindMovie = value

    # Lower case alias for RewindMovie setter
    @rewindmovie.setter
    def rewindmovie(self, value):
        self.RewindMovie = value

    @property
    def StopAfterSlides(self):
        return self.playsettings.StopAfterSlides

    # Lower case alias for StopAfterSlides
    @property
    def stopafterslides(self):
        return self.StopAfterSlides

    @StopAfterSlides.setter
    def StopAfterSlides(self, value):
        self.playsettings.StopAfterSlides = value

    # Lower case alias for StopAfterSlides setter
    @stopafterslides.setter
    def stopafterslides(self, value):
        self.StopAfterSlides = value


class PlotArea:

    def __init__(self, plotarea=None):
        self.plotarea = plotarea

    @property
    def Application(self):
        return self.plotarea.Application

    @property
    def Creator(self):
        return self.plotarea.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.plotarea.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def Height(self):
        return self.plotarea.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @Height.setter
    def Height(self, value):
        self.plotarea.Height = value

    # Lower case alias for Height setter
    @height.setter
    def height(self, value):
        self.Height = value

    @property
    def InsideHeight(self):
        return self.plotarea.InsideHeight

    # Lower case alias for InsideHeight
    @property
    def insideheight(self):
        return self.InsideHeight

    @InsideHeight.setter
    def InsideHeight(self, value):
        self.plotarea.InsideHeight = value

    # Lower case alias for InsideHeight setter
    @insideheight.setter
    def insideheight(self, value):
        self.InsideHeight = value

    @property
    def InsideLeft(self):
        return self.plotarea.InsideLeft

    # Lower case alias for InsideLeft
    @property
    def insideleft(self):
        return self.InsideLeft

    @InsideLeft.setter
    def InsideLeft(self, value):
        self.plotarea.InsideLeft = value

    # Lower case alias for InsideLeft setter
    @insideleft.setter
    def insideleft(self, value):
        self.InsideLeft = value

    @property
    def InsideTop(self):
        return self.plotarea.InsideTop

    # Lower case alias for InsideTop
    @property
    def insidetop(self):
        return self.InsideTop

    @InsideTop.setter
    def InsideTop(self, value):
        self.plotarea.InsideTop = value

    # Lower case alias for InsideTop setter
    @insidetop.setter
    def insidetop(self, value):
        self.InsideTop = value

    @property
    def InsideWidth(self):
        return self.plotarea.InsideWidth

    # Lower case alias for InsideWidth
    @property
    def insidewidth(self):
        return self.InsideWidth

    @InsideWidth.setter
    def InsideWidth(self, value):
        self.plotarea.InsideWidth = value

    # Lower case alias for InsideWidth setter
    @insidewidth.setter
    def insidewidth(self, value):
        self.InsideWidth = value

    @property
    def Left(self):
        return self.plotarea.Left

    # Lower case alias for Left
    @property
    def left(self):
        return self.Left

    @Left.setter
    def Left(self, value):
        self.plotarea.Left = value

    # Lower case alias for Left setter
    @left.setter
    def left(self, value):
        self.Left = value

    @property
    def Name(self):
        return self.plotarea.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.plotarea.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return XlChartElementPosition(self.plotarea.Position)

    # Lower case alias for Position
    @property
    def position(self):
        return self.Position

    @Position.setter
    def Position(self, value):
        self.plotarea.Position = value

    # Lower case alias for Position setter
    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Top(self):
        return self.plotarea.Top

    # Lower case alias for Top
    @property
    def top(self):
        return self.Top

    @Top.setter
    def Top(self, value):
        self.plotarea.Top = value

    # Lower case alias for Top setter
    @top.setter
    def top(self, value):
        self.Top = value

    @property
    def Width(self):
        return self.plotarea.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

    @Width.setter
    def Width(self, value):
        self.plotarea.Width = value

    # Lower case alias for Width setter
    @width.setter
    def width(self, value):
        self.Width = value

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

    # Lower case alias for ApplyPictToEnd
    @property
    def applypicttoend(self):
        return self.ApplyPictToEnd

    @ApplyPictToEnd.setter
    def ApplyPictToEnd(self, value):
        self.point.ApplyPictToEnd = value

    # Lower case alias for ApplyPictToEnd setter
    @applypicttoend.setter
    def applypicttoend(self, value):
        self.ApplyPictToEnd = value

    @property
    def ApplyPictToFront(self):
        return self.point.ApplyPictToFront

    # Lower case alias for ApplyPictToFront
    @property
    def applypicttofront(self):
        return self.ApplyPictToFront

    @ApplyPictToFront.setter
    def ApplyPictToFront(self, value):
        self.point.ApplyPictToFront = value

    # Lower case alias for ApplyPictToFront setter
    @applypicttofront.setter
    def applypicttofront(self, value):
        self.ApplyPictToFront = value

    @property
    def ApplyPictToSides(self):
        return self.point.ApplyPictToSides

    # Lower case alias for ApplyPictToSides
    @property
    def applypicttosides(self):
        return self.ApplyPictToSides

    @ApplyPictToSides.setter
    def ApplyPictToSides(self, value):
        self.point.ApplyPictToSides = value

    # Lower case alias for ApplyPictToSides setter
    @applypicttosides.setter
    def applypicttosides(self, value):
        self.ApplyPictToSides = value

    @property
    def Creator(self):
        return self.point.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def DataLabel(self):
        return DataLabel(self.point.DataLabel)

    # Lower case alias for DataLabel
    @property
    def datalabel(self):
        return self.DataLabel

    @property
    def Explosion(self):
        return self.point.Explosion

    # Lower case alias for Explosion
    @property
    def explosion(self):
        return self.Explosion

    @Explosion.setter
    def Explosion(self, value):
        self.point.Explosion = value

    # Lower case alias for Explosion setter
    @explosion.setter
    def explosion(self, value):
        self.Explosion = value

    @property
    def Format(self):
        return ChartFormat(self.point.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def Has3DEffect(self):
        return self.point.Has3DEffect

    # Lower case alias for Has3DEffect
    @property
    def has3deffect(self):
        return self.Has3DEffect

    @Has3DEffect.setter
    def Has3DEffect(self, value):
        self.point.Has3DEffect = value

    # Lower case alias for Has3DEffect setter
    @has3deffect.setter
    def has3deffect(self, value):
        self.Has3DEffect = value

    @property
    def HasDataLabel(self):
        return self.point.HasDataLabel

    # Lower case alias for HasDataLabel
    @property
    def hasdatalabel(self):
        return self.HasDataLabel

    @HasDataLabel.setter
    def HasDataLabel(self, value):
        self.point.HasDataLabel = value

    # Lower case alias for HasDataLabel setter
    @hasdatalabel.setter
    def hasdatalabel(self, value):
        self.HasDataLabel = value

    @property
    def Height(self):
        return self.point.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @property
    def InvertIfNegative(self):
        return self.point.InvertIfNegative

    # Lower case alias for InvertIfNegative
    @property
    def invertifnegative(self):
        return self.InvertIfNegative

    @InvertIfNegative.setter
    def InvertIfNegative(self, value):
        self.point.InvertIfNegative = value

    # Lower case alias for InvertIfNegative setter
    @invertifnegative.setter
    def invertifnegative(self, value):
        self.InvertIfNegative = value

    @property
    def Left(self):
        return self.point.Left

    # Lower case alias for Left
    @property
    def left(self):
        return self.Left

    @property
    def MarkerBackgroundColor(self):
        return self.point.MarkerBackgroundColor

    # Lower case alias for MarkerBackgroundColor
    @property
    def markerbackgroundcolor(self):
        return self.MarkerBackgroundColor

    @MarkerBackgroundColor.setter
    def MarkerBackgroundColor(self, value):
        self.point.MarkerBackgroundColor = value

    # Lower case alias for MarkerBackgroundColor setter
    @markerbackgroundcolor.setter
    def markerbackgroundcolor(self, value):
        self.MarkerBackgroundColor = value

    @property
    def MarkerBackgroundColorIndex(self):
        return XlColorIndex(self.point.MarkerBackgroundColorIndex)

    # Lower case alias for MarkerBackgroundColorIndex
    @property
    def markerbackgroundcolorindex(self):
        return self.MarkerBackgroundColorIndex

    @MarkerBackgroundColorIndex.setter
    def MarkerBackgroundColorIndex(self, value):
        self.point.MarkerBackgroundColorIndex = value

    # Lower case alias for MarkerBackgroundColorIndex setter
    @markerbackgroundcolorindex.setter
    def markerbackgroundcolorindex(self, value):
        self.MarkerBackgroundColorIndex = value

    @property
    def MarkerForegroundColor(self):
        return self.point.MarkerForegroundColor

    # Lower case alias for MarkerForegroundColor
    @property
    def markerforegroundcolor(self):
        return self.MarkerForegroundColor

    @MarkerForegroundColor.setter
    def MarkerForegroundColor(self, value):
        self.point.MarkerForegroundColor = value

    # Lower case alias for MarkerForegroundColor setter
    @markerforegroundcolor.setter
    def markerforegroundcolor(self, value):
        self.MarkerForegroundColor = value

    @property
    def MarkerForegroundColorIndex(self):
        return XlColorIndex(self.point.MarkerForegroundColorIndex)

    # Lower case alias for MarkerForegroundColorIndex
    @property
    def markerforegroundcolorindex(self):
        return self.MarkerForegroundColorIndex

    @MarkerForegroundColorIndex.setter
    def MarkerForegroundColorIndex(self, value):
        self.point.MarkerForegroundColorIndex = value

    # Lower case alias for MarkerForegroundColorIndex setter
    @markerforegroundcolorindex.setter
    def markerforegroundcolorindex(self, value):
        self.MarkerForegroundColorIndex = value

    @property
    def MarkerSize(self):
        return self.point.MarkerSize

    # Lower case alias for MarkerSize
    @property
    def markersize(self):
        return self.MarkerSize

    @MarkerSize.setter
    def MarkerSize(self, value):
        self.point.MarkerSize = value

    # Lower case alias for MarkerSize setter
    @markersize.setter
    def markersize(self, value):
        self.MarkerSize = value

    @property
    def MarkerStyle(self):
        return XlMarkerStyle(self.point.MarkerStyle)

    # Lower case alias for MarkerStyle
    @property
    def markerstyle(self):
        return self.MarkerStyle

    @MarkerStyle.setter
    def MarkerStyle(self, value):
        self.point.MarkerStyle = value

    # Lower case alias for MarkerStyle setter
    @markerstyle.setter
    def markerstyle(self, value):
        self.MarkerStyle = value

    @property
    def Name(self):
        return self.point.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.point.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PictureType(self):
        return XlChartPictureType(self.point.PictureType)

    # Lower case alias for PictureType
    @property
    def picturetype(self):
        return self.PictureType

    @PictureType.setter
    def PictureType(self, value):
        self.point.PictureType = value

    # Lower case alias for PictureType setter
    @picturetype.setter
    def picturetype(self, value):
        self.PictureType = value

    @property
    def PictureUnit2(self):
        return self.point.PictureUnit2

    # Lower case alias for PictureUnit2
    @property
    def pictureunit2(self):
        return self.PictureUnit2

    @PictureUnit2.setter
    def PictureUnit2(self, value):
        self.point.PictureUnit2 = value

    # Lower case alias for PictureUnit2 setter
    @pictureunit2.setter
    def pictureunit2(self, value):
        self.PictureUnit2 = value

    @property
    def SecondaryPlot(self):
        return self.point.SecondaryPlot

    # Lower case alias for SecondaryPlot
    @property
    def secondaryplot(self):
        return self.SecondaryPlot

    @SecondaryPlot.setter
    def SecondaryPlot(self, value):
        self.point.SecondaryPlot = value

    # Lower case alias for SecondaryPlot setter
    @secondaryplot.setter
    def secondaryplot(self, value):
        self.SecondaryPlot = value

    @property
    def Shadow(self):
        return self.point.Shadow

    # Lower case alias for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.point.Shadow = value

    # Lower case alias for Shadow setter
    @shadow.setter
    def shadow(self, value):
        self.Shadow = value

    @property
    def Top(self):
        return self.point.Top

    # Lower case alias for Top
    @property
    def top(self):
        return self.Top

    @property
    def Width(self):
        return self.point.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.points.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Parent(self):
        return self.points.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Point(self.points.Item(*arguments))


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

    # Lower case alias for AutoSaveOn
    @property
    def autosaveon(self):
        return self.AutoSaveOn

    @AutoSaveOn.setter
    def AutoSaveOn(self, value):
        self.presentation.AutoSaveOn = value

    # Lower case alias for AutoSaveOn setter
    @autosaveon.setter
    def autosaveon(self, value):
        self.AutoSaveOn = value

    @property
    def Broadcast(self):
        return Broadcast(self.presentation.Broadcast)

    # Lower case alias for Broadcast
    @property
    def broadcast(self):
        return self.Broadcast

    @property
    def BuiltInDocumentProperties(self):
        return self.presentation.BuiltInDocumentProperties

    # Lower case alias for BuiltInDocumentProperties
    @property
    def builtindocumentproperties(self):
        return self.BuiltInDocumentProperties

    @property
    def Coauthoring(self):
        return Coauthoring(self.presentation.Coauthoring)

    # Lower case alias for Coauthoring
    @property
    def coauthoring(self):
        return self.Coauthoring

    @property
    def ColorSchemes(self):
        return ColorSchemes(self.presentation.ColorSchemes)

    # Lower case alias for ColorSchemes
    @property
    def colorschemes(self):
        return self.ColorSchemes

    @property
    def CommandBars(self):
        return self.presentation.CommandBars

    # Lower case alias for CommandBars
    @property
    def commandbars(self):
        return self.CommandBars

    @property
    def Container(self):
        return self.presentation.Container

    # Lower case alias for Container
    @property
    def container(self):
        return self.Container

    @property
    def ContentTypeProperties(self):
        return self.presentation.ContentTypeProperties

    # Lower case alias for ContentTypeProperties
    @property
    def contenttypeproperties(self):
        return self.ContentTypeProperties

    @property
    def CreateVideoStatus(self):
        return Presentation(self.presentation.CreateVideoStatus)

    # Lower case alias for CreateVideoStatus
    @property
    def createvideostatus(self):
        return self.CreateVideoStatus

    @property
    def CustomDocumentProperties(self):
        return self.presentation.CustomDocumentProperties

    # Lower case alias for CustomDocumentProperties
    @property
    def customdocumentproperties(self):
        return self.CustomDocumentProperties

    @property
    def CustomerData(self):
        return CustomerData(self.presentation.CustomerData)

    # Lower case alias for CustomerData
    @property
    def customerdata(self):
        return self.CustomerData

    @property
    def CustomXMLParts(self):
        return self.presentation.CustomXMLParts

    # Lower case alias for CustomXMLParts
    @property
    def customxmlparts(self):
        return self.CustomXMLParts

    @property
    def DefaultLanguageID(self):
        return self.presentation.DefaultLanguageID

    # Lower case alias for DefaultLanguageID
    @property
    def defaultlanguageid(self):
        return self.DefaultLanguageID

    @DefaultLanguageID.setter
    def DefaultLanguageID(self, value):
        self.presentation.DefaultLanguageID = value

    # Lower case alias for DefaultLanguageID setter
    @defaultlanguageid.setter
    def defaultlanguageid(self, value):
        self.DefaultLanguageID = value

    @property
    def DefaultShape(self):
        return Shape(self.presentation.DefaultShape)

    # Lower case alias for DefaultShape
    @property
    def defaultshape(self):
        return self.DefaultShape

    @property
    def Designs(self):
        return Designs(self.presentation.Designs)

    # Lower case alias for Designs
    @property
    def designs(self):
        return self.Designs

    @property
    def DisplayComments(self):
        return self.presentation.DisplayComments

    # Lower case alias for DisplayComments
    @property
    def displaycomments(self):
        return self.DisplayComments

    @DisplayComments.setter
    def DisplayComments(self, value):
        self.presentation.DisplayComments = value

    # Lower case alias for DisplayComments setter
    @displaycomments.setter
    def displaycomments(self, value):
        self.DisplayComments = value

    @property
    def DocumentInspectors(self):
        return self.presentation.DocumentInspectors

    # Lower case alias for DocumentInspectors
    @property
    def documentinspectors(self):
        return self.DocumentInspectors

    @property
    def DocumentLibraryVersions(self):
        return self.presentation.DocumentLibraryVersions

    # Lower case alias for DocumentLibraryVersions
    @property
    def documentlibraryversions(self):
        return self.DocumentLibraryVersions

    @property
    def EncryptionProvider(self):
        return self.presentation.EncryptionProvider

    # Lower case alias for EncryptionProvider
    @property
    def encryptionprovider(self):
        return self.EncryptionProvider

    @EncryptionProvider.setter
    def EncryptionProvider(self, value):
        self.presentation.EncryptionProvider = value

    # Lower case alias for EncryptionProvider setter
    @encryptionprovider.setter
    def encryptionprovider(self, value):
        self.EncryptionProvider = value

    @property
    def EnvelopeVisible(self):
        return self.presentation.EnvelopeVisible

    # Lower case alias for EnvelopeVisible
    @property
    def envelopevisible(self):
        return self.EnvelopeVisible

    @EnvelopeVisible.setter
    def EnvelopeVisible(self, value):
        self.presentation.EnvelopeVisible = value

    # Lower case alias for EnvelopeVisible setter
    @envelopevisible.setter
    def envelopevisible(self, value):
        self.EnvelopeVisible = value

    @property
    def ExtraColors(self):
        return ExtraColors(self.presentation.ExtraColors)

    # Lower case alias for ExtraColors
    @property
    def extracolors(self):
        return self.ExtraColors

    @property
    def FarEastLineBreakLanguage(self):
        return self.presentation.FarEastLineBreakLanguage

    # Lower case alias for FarEastLineBreakLanguage
    @property
    def fareastlinebreaklanguage(self):
        return self.FarEastLineBreakLanguage

    @FarEastLineBreakLanguage.setter
    def FarEastLineBreakLanguage(self, value):
        self.presentation.FarEastLineBreakLanguage = value

    # Lower case alias for FarEastLineBreakLanguage setter
    @fareastlinebreaklanguage.setter
    def fareastlinebreaklanguage(self, value):
        self.FarEastLineBreakLanguage = value

    @property
    def FarEastLineBreakLevel(self):
        return self.presentation.FarEastLineBreakLevel

    # Lower case alias for FarEastLineBreakLevel
    @property
    def fareastlinebreaklevel(self):
        return self.FarEastLineBreakLevel

    @FarEastLineBreakLevel.setter
    def FarEastLineBreakLevel(self, value):
        self.presentation.FarEastLineBreakLevel = value

    # Lower case alias for FarEastLineBreakLevel setter
    @fareastlinebreaklevel.setter
    def fareastlinebreaklevel(self, value):
        self.FarEastLineBreakLevel = value

    @property
    def Final(self):
        return self.presentation.Final

    # Lower case alias for Final
    @property
    def final(self):
        return self.Final

    @Final.setter
    def Final(self, value):
        self.presentation.Final = value

    # Lower case alias for Final setter
    @final.setter
    def final(self, value):
        self.Final = value

    @property
    def Fonts(self):
        return Fonts(self.presentation.Fonts)

    # Lower case alias for Fonts
    @property
    def fonts(self):
        return self.Fonts

    @property
    def FullName(self):
        return self.presentation.FullName

    # Lower case alias for FullName
    @property
    def fullname(self):
        return self.FullName

    @property
    def GridDistance(self):
        return self.presentation.GridDistance

    # Lower case alias for GridDistance
    @property
    def griddistance(self):
        return self.GridDistance

    @GridDistance.setter
    def GridDistance(self, value):
        self.presentation.GridDistance = value

    # Lower case alias for GridDistance setter
    @griddistance.setter
    def griddistance(self, value):
        self.GridDistance = value

    @property
    def HandoutMaster(self):
        return Master(self.presentation.HandoutMaster)

    # Lower case alias for HandoutMaster
    @property
    def handoutmaster(self):
        return self.HandoutMaster

    @property
    def HasHandoutMaster(self):
        return self.presentation.HasHandoutMaster

    # Lower case alias for HasHandoutMaster
    @property
    def hashandoutmaster(self):
        return self.HasHandoutMaster

    @property
    def HasNotesMaster(self):
        return self.presentation.HasNotesMaster

    # Lower case alias for HasNotesMaster
    @property
    def hasnotesmaster(self):
        return self.HasNotesMaster

    @property
    def HasTitleMaster(self):
        return self.presentation.HasTitleMaster

    # Lower case alias for HasTitleMaster
    @property
    def hastitlemaster(self):
        return self.HasTitleMaster

    @property
    def HasVBProject(self):
        return self.presentation.HasVBProject

    # Lower case alias for HasVBProject
    @property
    def hasvbproject(self):
        return self.HasVBProject

    @property
    def InMergeMode(self):
        return self.presentation.InMergeMode

    # Lower case alias for InMergeMode
    @property
    def inmergemode(self):
        return self.InMergeMode

    @property
    def IsFullyDownloaded(self):
        return self.presentation.IsFullyDownloaded

    # Lower case alias for IsFullyDownloaded
    @property
    def isfullydownloaded(self):
        return self.IsFullyDownloaded

    @property
    def LayoutDirection(self):
        return self.presentation.LayoutDirection

    # Lower case alias for LayoutDirection
    @property
    def layoutdirection(self):
        return self.LayoutDirection

    @LayoutDirection.setter
    def LayoutDirection(self, value):
        self.presentation.LayoutDirection = value

    # Lower case alias for LayoutDirection setter
    @layoutdirection.setter
    def layoutdirection(self, value):
        self.LayoutDirection = value

    @property
    def Name(self):
        return self.presentation.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def NoLineBreakAfter(self):
        return self.presentation.NoLineBreakAfter

    # Lower case alias for NoLineBreakAfter
    @property
    def nolinebreakafter(self):
        return self.NoLineBreakAfter

    @NoLineBreakAfter.setter
    def NoLineBreakAfter(self, value):
        self.presentation.NoLineBreakAfter = value

    # Lower case alias for NoLineBreakAfter setter
    @nolinebreakafter.setter
    def nolinebreakafter(self, value):
        self.NoLineBreakAfter = value

    @property
    def NoLineBreakBefore(self):
        return self.presentation.NoLineBreakBefore

    # Lower case alias for NoLineBreakBefore
    @property
    def nolinebreakbefore(self):
        return self.NoLineBreakBefore

    @NoLineBreakBefore.setter
    def NoLineBreakBefore(self, value):
        self.presentation.NoLineBreakBefore = value

    # Lower case alias for NoLineBreakBefore setter
    @nolinebreakbefore.setter
    def nolinebreakbefore(self, value):
        self.NoLineBreakBefore = value

    @property
    def NotesMaster(self):
        return Master(self.presentation.NotesMaster)

    # Lower case alias for NotesMaster
    @property
    def notesmaster(self):
        return self.NotesMaster

    @property
    def PageSetup(self):
        return PageSetup(self.presentation.PageSetup)

    # Lower case alias for PageSetup
    @property
    def pagesetup(self):
        return self.PageSetup

    @property
    def Parent(self):
        return self.presentation.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Password(self):
        return self.presentation.Password

    # Lower case alias for Password
    @property
    def password(self):
        return self.Password

    @Password.setter
    def Password(self, value):
        self.presentation.Password = value

    # Lower case alias for Password setter
    @password.setter
    def password(self, value):
        self.Password = value

    @property
    def PasswordEncryptionAlgorithm(self):
        return self.presentation.PasswordEncryptionAlgorithm

    # Lower case alias for PasswordEncryptionAlgorithm
    @property
    def passwordencryptionalgorithm(self):
        return self.PasswordEncryptionAlgorithm

    @property
    def PasswordEncryptionFileProperties(self):
        return self.presentation.PasswordEncryptionFileProperties

    # Lower case alias for PasswordEncryptionFileProperties
    @property
    def passwordencryptionfileproperties(self):
        return self.PasswordEncryptionFileProperties

    @property
    def PasswordEncryptionKeyLength(self):
        return self.presentation.PasswordEncryptionKeyLength

    # Lower case alias for PasswordEncryptionKeyLength
    @property
    def passwordencryptionkeylength(self):
        return self.PasswordEncryptionKeyLength

    @property
    def PasswordEncryptionProvider(self):
        return self.presentation.PasswordEncryptionProvider

    # Lower case alias for PasswordEncryptionProvider
    @property
    def passwordencryptionprovider(self):
        return self.PasswordEncryptionProvider

    @property
    def Path(self):
        return Presentation(self.presentation.Path)

    # Lower case alias for Path
    @property
    def path(self):
        return self.Path

    @property
    def PrintOptions(self):
        return PrintOptions(self.presentation.PrintOptions)

    # Lower case alias for PrintOptions
    @property
    def printoptions(self):
        return self.PrintOptions

    @property
    def ReadOnly(self):
        return self.presentation.ReadOnly

    # Lower case alias for ReadOnly
    @property
    def readonly(self):
        return self.ReadOnly

    @property
    def ReadOnlyRecommended(self):
        return self.presentation.ReadOnlyRecommended

    # Lower case alias for ReadOnlyRecommended
    @property
    def readonlyrecommended(self):
        return self.ReadOnlyRecommended

    @property
    def RemovePersonalInformation(self):
        return self.presentation.RemovePersonalInformation

    # Lower case alias for RemovePersonalInformation
    @property
    def removepersonalinformation(self):
        return self.RemovePersonalInformation

    @RemovePersonalInformation.setter
    def RemovePersonalInformation(self, value):
        self.presentation.RemovePersonalInformation = value

    # Lower case alias for RemovePersonalInformation setter
    @removepersonalinformation.setter
    def removepersonalinformation(self, value):
        self.RemovePersonalInformation = value

    @property
    def Research(self):
        return Research(self.presentation.Research)

    # Lower case alias for Research
    @property
    def research(self):
        return self.Research

    @property
    def Saved(self):
        return self.presentation.Saved

    # Lower case alias for Saved
    @property
    def saved(self):
        return self.Saved

    @Saved.setter
    def Saved(self, value):
        self.presentation.Saved = value

    # Lower case alias for Saved setter
    @saved.setter
    def saved(self, value):
        self.Saved = value

    @property
    def SectionProperties(self):
        return SectionProperties(self.presentation.SectionProperties)

    # Lower case alias for SectionProperties
    @property
    def sectionproperties(self):
        return self.SectionProperties

    @property
    def SensitivityLabel(self):
        return self.presentation.SensitivityLabel

    # Lower case alias for SensitivityLabel
    @property
    def sensitivitylabel(self):
        return self.SensitivityLabel

    @property
    def ServerPolicy(self):
        return self.presentation.ServerPolicy

    # Lower case alias for ServerPolicy
    @property
    def serverpolicy(self):
        return self.ServerPolicy

    @property
    def SharedWorkspace(self):
        return self.presentation.SharedWorkspace

    # Lower case alias for SharedWorkspace
    @property
    def sharedworkspace(self):
        return self.SharedWorkspace

    @property
    def Signatures(self):
        return self.presentation.Signatures

    # Lower case alias for Signatures
    @property
    def signatures(self):
        return self.Signatures

    @property
    def SlideMaster(self):
        return Master(self.presentation.SlideMaster)

    # Lower case alias for SlideMaster
    @property
    def slidemaster(self):
        return self.SlideMaster

    @property
    def Slides(self):
        return Slides(self.presentation.Slides)

    # Lower case alias for Slides
    @property
    def slides(self):
        return self.Slides

    @property
    def SlideShowSettings(self):
        return SlideShowSettings(self.presentation.SlideShowSettings)

    # Lower case alias for SlideShowSettings
    @property
    def slideshowsettings(self):
        return self.SlideShowSettings

    @property
    def SlideShowWindow(self):
        return SlideShowWindow(self.presentation.SlideShowWindow)

    # Lower case alias for SlideShowWindow
    @property
    def slideshowwindow(self):
        return self.SlideShowWindow

    @property
    def SnapToGrid(self):
        return self.presentation.SnapToGrid

    # Lower case alias for SnapToGrid
    @property
    def snaptogrid(self):
        return self.SnapToGrid

    @SnapToGrid.setter
    def SnapToGrid(self, value):
        self.presentation.SnapToGrid = value

    # Lower case alias for SnapToGrid setter
    @snaptogrid.setter
    def snaptogrid(self, value):
        self.SnapToGrid = value

    @property
    def Sync(self):
        return self.presentation.Sync

    # Lower case alias for Sync
    @property
    def sync(self):
        return self.Sync

    @property
    def Tags(self):
        return Tags(self.presentation.Tags)

    # Lower case alias for Tags
    @property
    def tags(self):
        return self.Tags

    @property
    def TemplateName(self):
        return self.presentation.TemplateName

    # Lower case alias for TemplateName
    @property
    def templatename(self):
        return self.TemplateName

    @property
    def TitleMaster(self):
        return Master(self.presentation.TitleMaster)

    # Lower case alias for TitleMaster
    @property
    def titlemaster(self):
        return self.TitleMaster

    @property
    def VBASigned(self):
        return self.presentation.VBASigned

    # Lower case alias for VBASigned
    @property
    def vbasigned(self):
        return self.VBASigned

    @property
    def VBProject(self):
        return self.presentation.VBProject

    # Lower case alias for VBProject
    @property
    def vbproject(self):
        return self.VBProject

    @property
    def Windows(self):
        return DocumentWindows(self.presentation.Windows)

    # Lower case alias for Windows
    @property
    def windows(self):
        return self.Windows

    @property
    def WritePassword(self):
        return self.presentation.WritePassword

    # Lower case alias for WritePassword
    @property
    def writepassword(self):
        return self.WritePassword

    @WritePassword.setter
    def WritePassword(self, value):
        self.presentation.WritePassword = value

    # Lower case alias for WritePassword setter
    @writepassword.setter
    def writepassword(self, value):
        self.WritePassword = value

    def AcceptAll(self):
        return self.presentation.AcceptAll()

    def AddTitleMaster(self):
        return self.presentation.AddTitleMaster()

    def AddToFavorites(self):
        self.presentation.AddToFavorites()

    def ApplyTemplate(self, FileName=None):
        arguments = com_arguments([FileName])
        self.presentation.ApplyTemplate(*arguments)

    def ApplyTheme(self, themeName=None):
        arguments = com_arguments([themeName])
        self.presentation.ApplyTheme(*arguments)

    def CanCheckIn(self):
        return self.presentation.CanCheckIn()

    def CheckIn(self, SaveChanges=None, Comments=None, MakePublic=None):
        arguments = com_arguments([SaveChanges, Comments, MakePublic])
        self.presentation.CheckIn(*arguments)

    def CheckInWithVersion(self, SaveChanges=None, Comments=None, MakePublic=None, VersionType=None):
        arguments = com_arguments([SaveChanges, Comments, MakePublic, VersionType])
        self.presentation.CheckInWithVersion(*arguments)

    def Close(self):
        self.presentation.Close()

    def Convert2(self, FileName=None):
        arguments = com_arguments([FileName])
        self.presentation.Convert2(*arguments)

    def CreateVideo(self, FileName=None, UseTimingsAndNarrations=None, DefaultSlideDuration=None, VertResolution=None, FramesPerSecond=None, Quality=None):
        arguments = com_arguments([FileName, UseTimingsAndNarrations, DefaultSlideDuration, VertResolution, FramesPerSecond, Quality])
        self.presentation.CreateVideo(*arguments)

    def EndReview(self):
        return self.presentation.EndReview()

    def EnsureAllMediaUpgraded(self):
        self.presentation.EnsureAllMediaUpgraded()

    def Export(self, Path=None, FilterName=None, ScaleWidth=None, ScaleHeight=None):
        arguments = com_arguments([Path, FilterName, ScaleWidth, ScaleHeight])
        self.presentation.Export(*arguments)

    def ExportAsFixedFormat(self, Path=None, FixedFormatType=None, Intent=None, FrameSlides=None, HandoutOrder=None, OutputType=None, PrintHiddenSlides=None, PrintRange=None, RangeType=None, SlideShowName=None, IncludeDocProperties=None, KeepIRMSettings=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005_1=None, ExternalExporter=None):
        arguments = com_arguments([Path, FixedFormatType, Intent, FrameSlides, HandoutOrder, OutputType, PrintHiddenSlides, PrintRange, RangeType, SlideShowName, IncludeDocProperties, KeepIRMSettings, DocStructureTags, BitmapMissingFonts, UseISO19005_1, ExternalExporter])
        self.presentation.ExportAsFixedFormat(*arguments)

    def FollowHyperlink(self, Address=None, SubAddress=None, NewWindow=None, AddHistory=None, ExtraInfo=None, Method=None, HeaderInfo=None):
        arguments = com_arguments([Address, SubAddress, NewWindow, AddHistory, ExtraInfo, Method, HeaderInfo])
        return self.presentation.FollowHyperlink(*arguments)

    def GetWorkflowTasks(self):
        return self.presentation.GetWorkflowTasks()

    def GetWorkflowTemplates(self):
        return self.presentation.GetWorkflowTemplates()

    def LockServerFile(self):
        self.presentation.LockServerFile()

    def MergeWithBaseline(self, withPresentation=None, baselinePresentation=None):
        arguments = com_arguments([withPresentation, baselinePresentation])
        return self.presentation.MergeWithBaseline(*arguments)

    def NewWindow(self):
        return self.presentation.NewWindow()

    def PrintOut(self, From=None, To=None, PrintToFile=None, Copies=None, Collate=None):
        arguments = com_arguments([From, To, PrintToFile, Copies, Collate])
        self.presentation.PrintOut(*arguments)

    def PublishSlides(self, SlideLibraryUrl=None, Overwrite=None):
        arguments = com_arguments([SlideLibraryUrl, Overwrite])
        self.presentation.PublishSlides(*arguments)

    def RejectAll(self):
        return self.presentation.RejectAll()

    def RemoveDocumentInformation(self, Type=None):
        arguments = com_arguments([Type])
        self.presentation.RemoveDocumentInformation(*arguments)

    def Save(self):
        self.presentation.Save()

    def SaveAs(self, FileName=None, FileFormat=None, EmbedFonts=None):
        arguments = com_arguments([FileName, FileFormat, EmbedFonts])
        self.presentation.SaveAs(*arguments)

    def SaveCopyAs(self, FileName=None, FileFormat=None, EmbedTrueTypeFonts=None):
        arguments = com_arguments([FileName, FileFormat, EmbedTrueTypeFonts])
        self.presentation.SaveCopyAs(*arguments)

    def SaveCopyAs2(self, FileName=None, FileFormat=None, EmbedTrueTypeFonts=None, ReadOnlyRecommended=None):
        arguments = com_arguments([FileName, FileFormat, EmbedTrueTypeFonts, ReadOnlyRecommended])
        self.presentation.SaveCopyAs2(*arguments)

    def SendFaxOverInternet(self, Recipients=None, Subject=None, ShowMessage=None):
        arguments = com_arguments([Recipients, Subject, ShowMessage])
        self.presentation.SendFaxOverInternet(*arguments)

    def SetPasswordEncryptionOptions(self, PasswordEncryptionProvider=None, PasswordEncryptionAlgorithm=None, PasswordEncryptionKeyLength=None, PasswordEncryptionFileProperties=None):
        arguments = com_arguments([PasswordEncryptionProvider, PasswordEncryptionAlgorithm, PasswordEncryptionKeyLength, PasswordEncryptionFileProperties])
        self.presentation.SetPasswordEncryptionOptions(*arguments)

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.presentations.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Add(self, WithWindow=None):
        arguments = com_arguments([WithWindow])
        return Presentation(self.presentations.Add(*arguments))

    def CanCheckOut(self, FileName=None):
        arguments = com_arguments([FileName])
        return self.presentations.CanCheckOut(*arguments)

    def CheckOut(self, FileName=None):
        arguments = com_arguments([FileName])
        return self.presentations.CheckOut(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.presentations.Item(*arguments)

    def Open(self, FileName=None, ReadOnly=None, Untitled=None, WithWindow=None):
        arguments = com_arguments([FileName, ReadOnly, Untitled, WithWindow])
        return Presentation(self.presentations.Open(*arguments))

    def Open2007(self, FileName=None, ReadOnly=None, Untitled=None, WithWindow=None, OpenAndRepair=None):
        arguments = com_arguments([FileName, ReadOnly, Untitled, WithWindow, OpenAndRepair])
        return Presentation(self.presentations.Open2007(*arguments))


class PrintOptions:

    def __init__(self, printoptions=None):
        self.printoptions = printoptions

    @property
    def ActivePrinter(self):
        return self.printoptions.ActivePrinter

    # Lower case alias for ActivePrinter
    @property
    def activeprinter(self):
        return self.ActivePrinter

    @property
    def Application(self):
        return Application(self.printoptions.Application)

    @property
    def Collate(self):
        return self.printoptions.Collate

    # Lower case alias for Collate
    @property
    def collate(self):
        return self.Collate

    @Collate.setter
    def Collate(self, value):
        self.printoptions.Collate = value

    # Lower case alias for Collate setter
    @collate.setter
    def collate(self, value):
        self.Collate = value

    @property
    def FitToPage(self):
        return self.printoptions.FitToPage

    # Lower case alias for FitToPage
    @property
    def fittopage(self):
        return self.FitToPage

    @FitToPage.setter
    def FitToPage(self, value):
        self.printoptions.FitToPage = value

    # Lower case alias for FitToPage setter
    @fittopage.setter
    def fittopage(self, value):
        self.FitToPage = value

    @property
    def FrameSlides(self):
        return self.printoptions.FrameSlides

    # Lower case alias for FrameSlides
    @property
    def frameslides(self):
        return self.FrameSlides

    @FrameSlides.setter
    def FrameSlides(self, value):
        self.printoptions.FrameSlides = value

    # Lower case alias for FrameSlides setter
    @frameslides.setter
    def frameslides(self, value):
        self.FrameSlides = value

    @property
    def HandoutOrder(self):
        return self.printoptions.HandoutOrder

    # Lower case alias for HandoutOrder
    @property
    def handoutorder(self):
        return self.HandoutOrder

    @HandoutOrder.setter
    def HandoutOrder(self, value):
        self.printoptions.HandoutOrder = value

    # Lower case alias for HandoutOrder setter
    @handoutorder.setter
    def handoutorder(self, value):
        self.HandoutOrder = value

    @property
    def HighQuality(self):
        return self.printoptions.HighQuality

    # Lower case alias for HighQuality
    @property
    def highquality(self):
        return self.HighQuality

    @HighQuality.setter
    def HighQuality(self, value):
        self.printoptions.HighQuality = value

    # Lower case alias for HighQuality setter
    @highquality.setter
    def highquality(self, value):
        self.HighQuality = value

    @property
    def NumberOfCopies(self):
        return self.printoptions.NumberOfCopies

    # Lower case alias for NumberOfCopies
    @property
    def numberofcopies(self):
        return self.NumberOfCopies

    @NumberOfCopies.setter
    def NumberOfCopies(self, value):
        self.printoptions.NumberOfCopies = value

    # Lower case alias for NumberOfCopies setter
    @numberofcopies.setter
    def numberofcopies(self, value):
        self.NumberOfCopies = value

    @property
    def OutputType(self):
        return self.printoptions.OutputType

    # Lower case alias for OutputType
    @property
    def outputtype(self):
        return self.OutputType

    @OutputType.setter
    def OutputType(self, value):
        self.printoptions.OutputType = value

    # Lower case alias for OutputType setter
    @outputtype.setter
    def outputtype(self, value):
        self.OutputType = value

    @property
    def Parent(self):
        return self.printoptions.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PrintColorType(self):
        return self.printoptions.PrintColorType

    # Lower case alias for PrintColorType
    @property
    def printcolortype(self):
        return self.PrintColorType

    @PrintColorType.setter
    def PrintColorType(self, value):
        self.printoptions.PrintColorType = value

    # Lower case alias for PrintColorType setter
    @printcolortype.setter
    def printcolortype(self, value):
        self.PrintColorType = value

    @property
    def PrintComments(self):
        return self.printoptions.PrintComments

    # Lower case alias for PrintComments
    @property
    def printcomments(self):
        return self.PrintComments

    @PrintComments.setter
    def PrintComments(self, value):
        self.printoptions.PrintComments = value

    # Lower case alias for PrintComments setter
    @printcomments.setter
    def printcomments(self, value):
        self.PrintComments = value

    @property
    def PrintFontsAsGraphics(self):
        return self.printoptions.PrintFontsAsGraphics

    # Lower case alias for PrintFontsAsGraphics
    @property
    def printfontsasgraphics(self):
        return self.PrintFontsAsGraphics

    @PrintFontsAsGraphics.setter
    def PrintFontsAsGraphics(self, value):
        self.printoptions.PrintFontsAsGraphics = value

    # Lower case alias for PrintFontsAsGraphics setter
    @printfontsasgraphics.setter
    def printfontsasgraphics(self, value):
        self.PrintFontsAsGraphics = value

    @property
    def PrintHiddenSlides(self):
        return self.printoptions.PrintHiddenSlides

    # Lower case alias for PrintHiddenSlides
    @property
    def printhiddenslides(self):
        return self.PrintHiddenSlides

    @PrintHiddenSlides.setter
    def PrintHiddenSlides(self, value):
        self.printoptions.PrintHiddenSlides = value

    # Lower case alias for PrintHiddenSlides setter
    @printhiddenslides.setter
    def printhiddenslides(self, value):
        self.PrintHiddenSlides = value

    @property
    def PrintInBackground(self):
        return self.printoptions.PrintInBackground

    # Lower case alias for PrintInBackground
    @property
    def printinbackground(self):
        return self.PrintInBackground

    @PrintInBackground.setter
    def PrintInBackground(self, value):
        self.printoptions.PrintInBackground = value

    # Lower case alias for PrintInBackground setter
    @printinbackground.setter
    def printinbackground(self, value):
        self.PrintInBackground = value

    @property
    def Ranges(self):
        return PrintRanges(self.printoptions.Ranges)

    # Lower case alias for Ranges
    @property
    def ranges(self):
        return self.Ranges

    @property
    def RangeType(self):
        return self.printoptions.RangeType

    # Lower case alias for RangeType
    @property
    def rangetype(self):
        return self.RangeType

    @RangeType.setter
    def RangeType(self, value):
        self.printoptions.RangeType = value

    # Lower case alias for RangeType setter
    @rangetype.setter
    def rangetype(self, value):
        self.RangeType = value

    @property
    def sectionIndex(self):
        return PrintOptions(self.printoptions.sectionIndex)

    # Lower case alias for sectionIndex
    @property
    def sectionindex(self):
        return self.sectionIndex

    @sectionIndex.setter
    def sectionIndex(self, value):
        self.printoptions.sectionIndex = value

    # Lower case alias for sectionIndex setter
    @sectionindex.setter
    def sectionindex(self, value):
        self.sectionIndex = value

    @property
    def SlideShowName(self):
        return self.printoptions.SlideShowName

    # Lower case alias for SlideShowName
    @property
    def slideshowname(self):
        return self.SlideShowName

    @SlideShowName.setter
    def SlideShowName(self, value):
        self.printoptions.SlideShowName = value

    # Lower case alias for SlideShowName setter
    @slideshowname.setter
    def slideshowname(self, value):
        self.SlideShowName = value


class PrintRange:

    def __init__(self, printrange=None):
        self.printrange = printrange

    @property
    def Application(self):
        return Application(self.printrange.Application)

    @property
    def End(self):
        return self.printrange.End

    # Lower case alias for End
    @property
    def end(self):
        return self.End

    @property
    def Parent(self):
        return self.printrange.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Start(self):
        return self.printrange.Start

    # Lower case alias for Start
    @property
    def start(self):
        return self.Start

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.printranges.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Add(self, Start=None, End=None):
        arguments = com_arguments([Start, End])
        return PrintRange(self.printranges.Add(*arguments))

    def ClearAll(self):
        return self.printranges.ClearAll()

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.printranges.Item(*arguments)


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

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Points(self):
        return AnimationPoints(self.propertyeffect.Points)

    # Lower case alias for Points
    @property
    def points(self):
        return self.Points

    @property
    def Property(self):
        return self.propertyeffect.Property

    @Property.setter
    def Property(self, value):
        self.propertyeffect.Property = value

    @property
    def To(self):
        return self.propertyeffect.To

    # Lower case alias for To
    @property
    def to(self):
        return self.To

    @To.setter
    def To(self, value):
        self.propertyeffect.To = value

    # Lower case alias for To setter
    @to.setter
    def to(self, value):
        self.To = value


class ProtectedViewWindow:

    def __init__(self, protectedviewwindow=None):
        self.protectedviewwindow = protectedviewwindow

    @property
    def Active(self):
        return self.protectedviewwindow.Active

    # Lower case alias for Active
    @property
    def active(self):
        return self.Active

    @property
    def Application(self):
        return Application(self.protectedviewwindow.Application)

    @property
    def Caption(self):
        return self.protectedviewwindow.Caption

    # Lower case alias for Caption
    @property
    def caption(self):
        return self.Caption

    @property
    def Height(self):
        return self.protectedviewwindow.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @Height.setter
    def Height(self, value):
        self.protectedviewwindow.Height = value

    # Lower case alias for Height setter
    @height.setter
    def height(self, value):
        self.Height = value

    @property
    def Left(self):
        return self.protectedviewwindow.Left

    # Lower case alias for Left
    @property
    def left(self):
        return self.Left

    @Left.setter
    def Left(self, value):
        self.protectedviewwindow.Left = value

    # Lower case alias for Left setter
    @left.setter
    def left(self, value):
        self.Left = value

    @property
    def Parent(self):
        return self.protectedviewwindow.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Presentation(self):
        return Presentation(self.protectedviewwindow.Presentation)

    # Lower case alias for Presentation
    @property
    def presentation(self):
        return self.Presentation

    @property
    def SourceName(self):
        return ProtectedViewWindow(self.protectedviewwindow.SourceName)

    # Lower case alias for SourceName
    @property
    def sourcename(self):
        return self.SourceName

    @property
    def SourcePath(self):
        return ProtectedViewWindow(self.protectedviewwindow.SourcePath)

    # Lower case alias for SourcePath
    @property
    def sourcepath(self):
        return self.SourcePath

    @property
    def Top(self):
        return self.protectedviewwindow.Top

    # Lower case alias for Top
    @property
    def top(self):
        return self.Top

    @Top.setter
    def Top(self, value):
        self.protectedviewwindow.Top = value

    # Lower case alias for Top setter
    @top.setter
    def top(self, value):
        self.Top = value

    @property
    def Width(self):
        return self.protectedviewwindow.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

    @Width.setter
    def Width(self, value):
        self.protectedviewwindow.Width = value

    # Lower case alias for Width setter
    @width.setter
    def width(self, value):
        self.Width = value

    @property
    def WindowState(self):
        return self.protectedviewwindow.WindowState

    # Lower case alias for WindowState
    @property
    def windowstate(self):
        return self.WindowState

    @WindowState.setter
    def WindowState(self, value):
        self.protectedviewwindow.WindowState = value

    # Lower case alias for WindowState setter
    @windowstate.setter
    def windowstate(self, value):
        self.WindowState = value

    def Activate(self):
        self.protectedviewwindow.Activate()

    def Close(self):
        self.protectedviewwindow.Close()

    def Edit(self, ModifyPassword=None):
        arguments = com_arguments([ModifyPassword])
        return self.protectedviewwindow.Edit(*arguments)


class ProtectedViewWindows:

    def __init__(self, protectedviewwindows=None):
        self.protectedviewwindows = protectedviewwindows

    @property
    def Application(self):
        return Application(self.protectedviewwindows.Application)

    @property
    def Count(self):
        return self.protectedviewwindows.Count

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.protectedviewwindows.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.protectedviewwindows.Item(*arguments)

    def Open(self, FileName=None, ReadPassword=None, OpenAndRepair=None):
        arguments = com_arguments([FileName, ReadPassword, OpenAndRepair])
        return self.protectedviewwindows.Open(*arguments)


class PublishObject:

    def __init__(self, publishobject=None):
        self.publishobject = publishobject

    @property
    def Application(self):
        return Application(self.publishobject.Application)

    @property
    def FileName(self):
        return self.publishobject.FileName

    # Lower case alias for FileName
    @property
    def filename(self):
        return self.FileName

    @FileName.setter
    def FileName(self, value):
        self.publishobject.FileName = value

    # Lower case alias for FileName setter
    @filename.setter
    def filename(self, value):
        self.FileName = value

    @property
    def HTMLVersion(self):
        return self.publishobject.HTMLVersion

    # Lower case alias for HTMLVersion
    @property
    def htmlversion(self):
        return self.HTMLVersion

    @HTMLVersion.setter
    def HTMLVersion(self, value):
        self.publishobject.HTMLVersion = value

    # Lower case alias for HTMLVersion setter
    @htmlversion.setter
    def htmlversion(self, value):
        self.HTMLVersion = value

    @property
    def Parent(self):
        return self.publishobject.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RangeEnd(self):
        return self.publishobject.RangeEnd

    # Lower case alias for RangeEnd
    @property
    def rangeend(self):
        return self.RangeEnd

    @RangeEnd.setter
    def RangeEnd(self, value):
        self.publishobject.RangeEnd = value

    # Lower case alias for RangeEnd setter
    @rangeend.setter
    def rangeend(self, value):
        self.RangeEnd = value

    @property
    def RangeStart(self):
        return self.publishobject.RangeStart

    # Lower case alias for RangeStart
    @property
    def rangestart(self):
        return self.RangeStart

    @RangeStart.setter
    def RangeStart(self, value):
        self.publishobject.RangeStart = value

    # Lower case alias for RangeStart setter
    @rangestart.setter
    def rangestart(self, value):
        self.RangeStart = value

    @property
    def SlideShowName(self):
        return self.publishobject.SlideShowName

    # Lower case alias for SlideShowName
    @property
    def slideshowname(self):
        return self.SlideShowName

    @SlideShowName.setter
    def SlideShowName(self, value):
        self.publishobject.SlideShowName = value

    # Lower case alias for SlideShowName setter
    @slideshowname.setter
    def slideshowname(self, value):
        self.SlideShowName = value

    @property
    def SourceType(self):
        return self.publishobject.SourceType

    # Lower case alias for SourceType
    @property
    def sourcetype(self):
        return self.SourceType

    @SourceType.setter
    def SourceType(self, value):
        self.publishobject.SourceType = value

    # Lower case alias for SourceType setter
    @sourcetype.setter
    def sourcetype(self, value):
        self.SourceType = value

    @property
    def SpeakerNotes(self):
        return self.publishobject.SpeakerNotes

    # Lower case alias for SpeakerNotes
    @property
    def speakernotes(self):
        return self.SpeakerNotes

    @SpeakerNotes.setter
    def SpeakerNotes(self, value):
        self.publishobject.SpeakerNotes = value

    # Lower case alias for SpeakerNotes setter
    @speakernotes.setter
    def speakernotes(self, value):
        self.SpeakerNotes = value

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.publishobjects.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.publishobjects.Item(*arguments)


class ResampleMediaTasks:

    def __init__(self, resamplemediatasks=None):
        self.resamplemediatasks = resamplemediatasks

    def __call__(self, item):
        return ResampleMediaTask(self.resamplemediatasks(item))

    @property
    def Count(self):
        return self.resamplemediatasks.Count

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def PercentComplete(self):
        return self.resamplemediatasks.PercentComplete

    # Lower case alias for PercentComplete
    @property
    def percentcomplete(self):
        return self.PercentComplete

    def Cancel(self):
        return self.resamplemediatasks.Cancel()

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.resamplemediatasks.Item(*arguments)

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

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def IsResearchService(self, ServiceID=None):
        arguments = com_arguments([ServiceID])
        return self.research.IsResearchService(*arguments)

    def Query(self, ServiceID=None, QueryString=None, QueryLanguage=None, UseSelection=None, RequeryContextXML=None, NewQueryContextXML=None, LaunchQuery=None):
        arguments = com_arguments([ServiceID, QueryString, QueryLanguage, UseSelection, RequeryContextXML, NewQueryContextXML, LaunchQuery])
        self.research.Query(*arguments)

    def SetLanguagePair(self, Language1=None, Language2=None):
        arguments = com_arguments([Language1, Language2])
        self.research.SetLanguagePair(*arguments)


class RGBColor:

    def __init__(self, rgbcolor=None):
        self.rgbcolor = rgbcolor

    @property
    def Application(self):
        return Application(self.rgbcolor.Application)

    @property
    def Parent(self):
        return self.rgbcolor.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RGB(self):
        return PpColorSchemeIndex(self.rgbcolor.RGB)

    # Lower case alias for RGB
    @property
    def rgb(self):
        return self.RGB

    @RGB.setter
    def RGB(self, value):
        self.rgbcolor.RGB = value

    # Lower case alias for RGB setter
    @rgb.setter
    def rgb(self, value):
        self.RGB = value


class RotationEffect:

    def __init__(self, rotationeffect=None):
        self.rotationeffect = rotationeffect

    @property
    def Application(self):
        return Application(self.rotationeffect.Application)

    @property
    def By(self):
        return self.rotationeffect.By

    # Lower case alias for By
    @property
    def by(self):
        return self.By

    @By.setter
    def By(self, value):
        self.rotationeffect.By = value

    # Lower case alias for By setter
    @by.setter
    def by(self, value):
        self.By = value

    @property
    def From(self):
        return self.rotationeffect.From

    @From.setter
    def From(self, value):
        self.rotationeffect.From = value

    @property
    def Parent(self):
        return self.rotationeffect.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def To(self):
        return self.rotationeffect.To

    # Lower case alias for To
    @property
    def to(self):
        return self.To

    @To.setter
    def To(self, value):
        self.rotationeffect.To = value

    # Lower case alias for To setter
    @to.setter
    def to(self, value):
        self.To = value


class Row:

    def __init__(self, row=None):
        self.row = row

    @property
    def Application(self):
        return Application(self.row.Application)

    def Cells(self, RowIndex=None, ColumnIndex=None):
        arguments = com_arguments([RowIndex, ColumnIndex])
        if callable(self.row.Cells):
            return CellRange(self.row.Cells(*arguments))
        else:
            return CellRange(self.row.GetCells(*arguments))

    # Lower case alias for Cells
    def cells(self, RowIndex=None, ColumnIndex=None):
        arguments = [RowIndex, ColumnIndex]
        return self.Cells(*arguments)

    @property
    def Height(self):
        return self.row.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @Height.setter
    def Height(self, value):
        self.row.Height = value

    # Lower case alias for Height setter
    @height.setter
    def height(self, value):
        self.Height = value

    @property
    def Parent(self):
        return self.row.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.rows.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Add(self, BeforeRow=None):
        arguments = com_arguments([BeforeRow])
        return Row(self.rows.Add(*arguments))

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.rows.Item(*arguments)


class Ruler:

    def __init__(self, ruler=None):
        self.ruler = ruler

    @property
    def Application(self):
        return Application(self.ruler.Application)

    @property
    def Levels(self):
        return RulerLevels(self.ruler.Levels)

    # Lower case alias for Levels
    @property
    def levels(self):
        return self.Levels

    @property
    def Parent(self):
        return self.ruler.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def TabStops(self):
        return TabStops(self.ruler.TabStops)

    # Lower case alias for TabStops
    @property
    def tabstops(self):
        return self.TabStops


class RulerLevel:

    def __init__(self, rulerlevel=None):
        self.rulerlevel = rulerlevel

    @property
    def Application(self):
        return Application(self.rulerlevel.Application)

    @property
    def FirstMargin(self):
        return self.rulerlevel.FirstMargin

    # Lower case alias for FirstMargin
    @property
    def firstmargin(self):
        return self.FirstMargin

    @FirstMargin.setter
    def FirstMargin(self, value):
        self.rulerlevel.FirstMargin = value

    # Lower case alias for FirstMargin setter
    @firstmargin.setter
    def firstmargin(self, value):
        self.FirstMargin = value

    @property
    def LeftMargin(self):
        return self.rulerlevel.LeftMargin

    # Lower case alias for LeftMargin
    @property
    def leftmargin(self):
        return self.LeftMargin

    @LeftMargin.setter
    def LeftMargin(self, value):
        self.rulerlevel.LeftMargin = value

    # Lower case alias for LeftMargin setter
    @leftmargin.setter
    def leftmargin(self, value):
        self.LeftMargin = value

    @property
    def Parent(self):
        return self.rulerlevel.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.rulerlevels.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.rulerlevels.Item(*arguments)


class ScaleEffect:

    def __init__(self, scaleeffect=None):
        self.scaleeffect = scaleeffect

    @property
    def Application(self):
        return Application(self.scaleeffect.Application)

    @property
    def ByX(self):
        return self.scaleeffect.ByX

    # Lower case alias for ByX
    @property
    def byx(self):
        return self.ByX

    @ByX.setter
    def ByX(self, value):
        self.scaleeffect.ByX = value

    # Lower case alias for ByX setter
    @byx.setter
    def byx(self, value):
        self.ByX = value

    @property
    def ByY(self):
        return self.scaleeffect.ByY

    # Lower case alias for ByY
    @property
    def byy(self):
        return self.ByY

    @ByY.setter
    def ByY(self, value):
        self.scaleeffect.ByY = value

    # Lower case alias for ByY setter
    @byy.setter
    def byy(self, value):
        self.ByY = value

    @property
    def FromX(self):
        return self.scaleeffect.FromX

    # Lower case alias for FromX
    @property
    def fromx(self):
        return self.FromX

    @FromX.setter
    def FromX(self, value):
        self.scaleeffect.FromX = value

    # Lower case alias for FromX setter
    @fromx.setter
    def fromx(self, value):
        self.FromX = value

    @property
    def FromY(self):
        return ScaleEffect(self.scaleeffect.FromY)

    # Lower case alias for FromY
    @property
    def fromy(self):
        return self.FromY

    @FromY.setter
    def FromY(self, value):
        self.scaleeffect.FromY = value

    # Lower case alias for FromY setter
    @fromy.setter
    def fromy(self, value):
        self.FromY = value

    @property
    def Parent(self):
        return self.scaleeffect.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ToX(self):
        return self.scaleeffect.ToX

    # Lower case alias for ToX
    @property
    def tox(self):
        return self.ToX

    @ToX.setter
    def ToX(self, value):
        self.scaleeffect.ToX = value

    # Lower case alias for ToX setter
    @tox.setter
    def tox(self, value):
        self.ToX = value

    @property
    def ToY(self):
        return ScaleEffect(self.scaleeffect.ToY)

    # Lower case alias for ToY
    @property
    def toy(self):
        return self.ToY

    @ToY.setter
    def ToY(self, value):
        self.scaleeffect.ToY = value

    # Lower case alias for ToY setter
    @toy.setter
    def toy(self, value):
        self.ToY = value


class SectionProperties:

    def __init__(self, sectionproperties=None):
        self.sectionproperties = sectionproperties

    @property
    def Application(self):
        return Application(self.sectionproperties.Application)

    @property
    def Count(self):
        return self.sectionproperties.Count

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.sectionproperties.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def AddBeforeSlide(self, SlideIndex=None, sectionName=None):
        arguments = com_arguments([SlideIndex, sectionName])
        return self.sectionproperties.AddBeforeSlide(*arguments)

    def AddSection(self, sectionIndex=None, sectionName=None):
        arguments = com_arguments([sectionIndex, sectionName])
        return self.sectionproperties.AddSection(*arguments)

    def Delete(self, sectionIndex=None, deleteSlides=None):
        arguments = com_arguments([sectionIndex, deleteSlides])
        self.sectionproperties.Delete(*arguments)

    def FirstSlide(self, sectionIndex=None):
        arguments = com_arguments([sectionIndex])
        return self.sectionproperties.FirstSlide(*arguments)

    def Move(self, sectionIndex=None, toPos=None):
        arguments = com_arguments([sectionIndex, toPos])
        self.sectionproperties.Move(*arguments)

    def Name(self, sectionIndex=None):
        arguments = com_arguments([sectionIndex])
        return self.sectionproperties.Name(*arguments)

    def Rename(self, sectionIndex=None, sectionName=None):
        arguments = com_arguments([sectionIndex, sectionName])
        self.sectionproperties.Rename(*arguments)

    def SectionID(self, sectionIndex=None):
        arguments = com_arguments([sectionIndex])
        return self.sectionproperties.SectionID(*arguments)

    def SlidesCount(self, sectionIndex=None):
        arguments = com_arguments([sectionIndex])
        return self.sectionproperties.SlidesCount(*arguments)


class Selection:

    def __init__(self, selection=None):
        self.selection = selection

    @property
    def Application(self):
        return Application(self.selection.Application)

    @property
    def ChildShapeRange(self):
        return ShapeRange(self.selection.ChildShapeRange)

    # Lower case alias for ChildShapeRange
    @property
    def childshaperange(self):
        return self.ChildShapeRange

    @property
    def HasChildShapeRange(self):
        return self.selection.HasChildShapeRange

    # Lower case alias for HasChildShapeRange
    @property
    def haschildshaperange(self):
        return self.HasChildShapeRange

    @property
    def Parent(self):
        return self.selection.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ShapeRange(self):
        return ShapeRange(self.selection.ShapeRange)

    # Lower case alias for ShapeRange
    @property
    def shaperange(self):
        return self.ShapeRange

    @property
    def SlideRange(self):
        return SlideRange(self.selection.SlideRange)

    # Lower case alias for SlideRange
    @property
    def sliderange(self):
        return self.SlideRange

    @property
    def TextRange(self):
        return TextRange(self.selection.TextRange)

    # Lower case alias for TextRange
    @property
    def textrange(self):
        return self.TextRange

    @property
    def TextRange2(self):
        return self.selection.TextRange2

    # Lower case alias for TextRange2
    @property
    def textrange2(self):
        return self.TextRange2

    @property
    def Type(self):
        return self.selection.Type

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.sequence.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def AddEffect(self, Shape=None, effectId=None, Level=None, trigger=None, Index=None):
        arguments = com_arguments([Shape, effectId, Level, trigger, Index])
        return self.sequence.AddEffect(*arguments)

    def AddTriggerEffect(self, pShape=None, effectId=None, trigger=None, pTriggerShape=None, bookmark=None, Level=None):
        arguments = com_arguments([pShape, effectId, trigger, pTriggerShape, bookmark, Level])
        return self.sequence.AddTriggerEffect(*arguments)

    def Clone(self, Effect=None, Index=None):
        arguments = com_arguments([Effect, Index])
        return self.sequence.Clone(*arguments)

    def ConvertToAfterEffect(self, Effect=None, After=None, DimColor=None, DimSchemeColor=None):
        arguments = com_arguments([Effect, After, DimColor, DimSchemeColor])
        return self.sequence.ConvertToAfterEffect(*arguments)

    def ConvertToAnimateBackground(self, Effect=None, AnimateBackground=None):
        arguments = com_arguments([Effect, AnimateBackground])
        return self.sequence.ConvertToAnimateBackground(*arguments)

    def ConvertToAnimateInReverse(self, Effect=None, animateInReverse=None):
        arguments = com_arguments([Effect, animateInReverse])
        return self.sequence.ConvertToAnimateInReverse(*arguments)

    def ConvertToBuildLevel(self, Effect=None, Level=None):
        arguments = com_arguments([Effect, Level])
        return self.sequence.ConvertToBuildLevel(*arguments)

    def ConvertToTextUnitEffect(self, Effect=None, unitEffect=None):
        arguments = com_arguments([Effect, unitEffect])
        return self.sequence.ConvertToTextUnitEffect(*arguments)

    def FindFirstAnimationFor(self, Shape=None):
        arguments = com_arguments([Shape])
        return self.sequence.FindFirstAnimationFor(*arguments)

    def FindFirstAnimationForClick(self, click=None):
        arguments = com_arguments([click])
        return self.sequence.FindFirstAnimationForClick(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.sequence.Item(*arguments)


class Sequences:

    def __init__(self, sequences=None):
        self.sequences = sequences

    @property
    def Application(self):
        return Application(self.sequences.Application)

    @property
    def Count(self):
        return self.sequences.Count

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.sequences.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Add(self, Index=None):
        arguments = com_arguments([Index])
        return self.sequences.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.sequences.Item(*arguments)


class Series:

    def __init__(self, series=None):
        self.series = series

    @property
    def Application(self):
        return self.series.Application

    @property
    def ApplyPictToEnd(self):
        return self.series.ApplyPictToEnd

    # Lower case alias for ApplyPictToEnd
    @property
    def applypicttoend(self):
        return self.ApplyPictToEnd

    @ApplyPictToEnd.setter
    def ApplyPictToEnd(self, value):
        self.series.ApplyPictToEnd = value

    # Lower case alias for ApplyPictToEnd setter
    @applypicttoend.setter
    def applypicttoend(self, value):
        self.ApplyPictToEnd = value

    @property
    def ApplyPictToFront(self):
        return self.series.ApplyPictToFront

    # Lower case alias for ApplyPictToFront
    @property
    def applypicttofront(self):
        return self.ApplyPictToFront

    @ApplyPictToFront.setter
    def ApplyPictToFront(self, value):
        self.series.ApplyPictToFront = value

    # Lower case alias for ApplyPictToFront setter
    @applypicttofront.setter
    def applypicttofront(self, value):
        self.ApplyPictToFront = value

    @property
    def ApplyPictToSides(self):
        return self.series.ApplyPictToSides

    # Lower case alias for ApplyPictToSides
    @property
    def applypicttosides(self):
        return self.ApplyPictToSides

    @ApplyPictToSides.setter
    def ApplyPictToSides(self, value):
        self.series.ApplyPictToSides = value

    # Lower case alias for ApplyPictToSides setter
    @applypicttosides.setter
    def applypicttosides(self, value):
        self.ApplyPictToSides = value

    @property
    def AxisGroup(self):
        return XlAxisGroup(self.series.AxisGroup)

    # Lower case alias for AxisGroup
    @property
    def axisgroup(self):
        return self.AxisGroup

    @AxisGroup.setter
    def AxisGroup(self, value):
        self.series.AxisGroup = value

    # Lower case alias for AxisGroup setter
    @axisgroup.setter
    def axisgroup(self, value):
        self.AxisGroup = value

    @property
    def BarShape(self):
        return XlBarShape(self.series.BarShape)

    # Lower case alias for BarShape
    @property
    def barshape(self):
        return self.BarShape

    @BarShape.setter
    def BarShape(self, value):
        self.series.BarShape = value

    # Lower case alias for BarShape setter
    @barshape.setter
    def barshape(self, value):
        self.BarShape = value

    @property
    def BubbleSizes(self):
        return self.series.BubbleSizes

    # Lower case alias for BubbleSizes
    @property
    def bubblesizes(self):
        return self.BubbleSizes

    @BubbleSizes.setter
    def BubbleSizes(self, value):
        self.series.BubbleSizes = value

    # Lower case alias for BubbleSizes setter
    @bubblesizes.setter
    def bubblesizes(self, value):
        self.BubbleSizes = value

    @property
    def ChartType(self):
        return self.series.ChartType

    # Lower case alias for ChartType
    @property
    def charttype(self):
        return self.ChartType

    @ChartType.setter
    def ChartType(self, value):
        self.series.ChartType = value

    # Lower case alias for ChartType setter
    @charttype.setter
    def charttype(self, value):
        self.ChartType = value

    @property
    def Creator(self):
        return self.series.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def ErrorBars(self):
        return ErrorBars(self.series.ErrorBars)

    # Lower case alias for ErrorBars
    @property
    def errorbars(self):
        return self.ErrorBars

    @property
    def Explosion(self):
        return self.series.Explosion

    # Lower case alias for Explosion
    @property
    def explosion(self):
        return self.Explosion

    @Explosion.setter
    def Explosion(self, value):
        self.series.Explosion = value

    # Lower case alias for Explosion setter
    @explosion.setter
    def explosion(self, value):
        self.Explosion = value

    @property
    def Format(self):
        return ChartFormat(self.series.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def Formula(self):
        return self.series.Formula

    # Lower case alias for Formula
    @property
    def formula(self):
        return self.Formula

    @Formula.setter
    def Formula(self, value):
        self.series.Formula = value

    # Lower case alias for Formula setter
    @formula.setter
    def formula(self, value):
        self.Formula = value

    @property
    def FormulaLocal(self):
        return self.series.FormulaLocal

    # Lower case alias for FormulaLocal
    @property
    def formulalocal(self):
        return self.FormulaLocal

    @FormulaLocal.setter
    def FormulaLocal(self, value):
        self.series.FormulaLocal = value

    # Lower case alias for FormulaLocal setter
    @formulalocal.setter
    def formulalocal(self, value):
        self.FormulaLocal = value

    @property
    def FormulaR1C1(self):
        return self.series.FormulaR1C1

    # Lower case alias for FormulaR1C1
    @property
    def formular1c1(self):
        return self.FormulaR1C1

    @FormulaR1C1.setter
    def FormulaR1C1(self, value):
        self.series.FormulaR1C1 = value

    # Lower case alias for FormulaR1C1 setter
    @formular1c1.setter
    def formular1c1(self, value):
        self.FormulaR1C1 = value

    @property
    def FormulaR1C1Local(self):
        return self.series.FormulaR1C1Local

    # Lower case alias for FormulaR1C1Local
    @property
    def formular1c1local(self):
        return self.FormulaR1C1Local

    @FormulaR1C1Local.setter
    def FormulaR1C1Local(self, value):
        self.series.FormulaR1C1Local = value

    # Lower case alias for FormulaR1C1Local setter
    @formular1c1local.setter
    def formular1c1local(self, value):
        self.FormulaR1C1Local = value

    @property
    def Has3DEffect(self):
        return self.series.Has3DEffect

    # Lower case alias for Has3DEffect
    @property
    def has3deffect(self):
        return self.Has3DEffect

    @Has3DEffect.setter
    def Has3DEffect(self, value):
        self.series.Has3DEffect = value

    # Lower case alias for Has3DEffect setter
    @has3deffect.setter
    def has3deffect(self, value):
        self.Has3DEffect = value

    @property
    def HasDataLabels(self):
        return self.series.HasDataLabels

    # Lower case alias for HasDataLabels
    @property
    def hasdatalabels(self):
        return self.HasDataLabels

    @HasDataLabels.setter
    def HasDataLabels(self, value):
        self.series.HasDataLabels = value

    # Lower case alias for HasDataLabels setter
    @hasdatalabels.setter
    def hasdatalabels(self, value):
        self.HasDataLabels = value

    @property
    def HasErrorBars(self):
        return self.series.HasErrorBars

    # Lower case alias for HasErrorBars
    @property
    def haserrorbars(self):
        return self.HasErrorBars

    @HasErrorBars.setter
    def HasErrorBars(self, value):
        self.series.HasErrorBars = value

    # Lower case alias for HasErrorBars setter
    @haserrorbars.setter
    def haserrorbars(self, value):
        self.HasErrorBars = value

    @property
    def HasLeaderLines(self):
        return self.series.HasLeaderLines

    # Lower case alias for HasLeaderLines
    @property
    def hasleaderlines(self):
        return self.HasLeaderLines

    @HasLeaderLines.setter
    def HasLeaderLines(self, value):
        self.series.HasLeaderLines = value

    # Lower case alias for HasLeaderLines setter
    @hasleaderlines.setter
    def hasleaderlines(self, value):
        self.HasLeaderLines = value

    @property
    def InvertColor(self):
        return self.series.InvertColor

    # Lower case alias for InvertColor
    @property
    def invertcolor(self):
        return self.InvertColor

    @InvertColor.setter
    def InvertColor(self, value):
        self.series.InvertColor = value

    # Lower case alias for InvertColor setter
    @invertcolor.setter
    def invertcolor(self, value):
        self.InvertColor = value

    @property
    def InvertColorIndex(self):
        return self.series.InvertColorIndex

    # Lower case alias for InvertColorIndex
    @property
    def invertcolorindex(self):
        return self.InvertColorIndex

    @InvertColorIndex.setter
    def InvertColorIndex(self, value):
        self.series.InvertColorIndex = value

    # Lower case alias for InvertColorIndex setter
    @invertcolorindex.setter
    def invertcolorindex(self, value):
        self.InvertColorIndex = value

    @property
    def InvertIfNegative(self):
        return self.series.InvertIfNegative

    # Lower case alias for InvertIfNegative
    @property
    def invertifnegative(self):
        return self.InvertIfNegative

    @InvertIfNegative.setter
    def InvertIfNegative(self, value):
        self.series.InvertIfNegative = value

    # Lower case alias for InvertIfNegative setter
    @invertifnegative.setter
    def invertifnegative(self, value):
        self.InvertIfNegative = value

    @property
    def LeaderLines(self):
        return LeaderLines(self.series.LeaderLines)

    # Lower case alias for LeaderLines
    @property
    def leaderlines(self):
        return self.LeaderLines

    @property
    def MarkerBackgroundColor(self):
        return self.series.MarkerBackgroundColor

    # Lower case alias for MarkerBackgroundColor
    @property
    def markerbackgroundcolor(self):
        return self.MarkerBackgroundColor

    @MarkerBackgroundColor.setter
    def MarkerBackgroundColor(self, value):
        self.series.MarkerBackgroundColor = value

    # Lower case alias for MarkerBackgroundColor setter
    @markerbackgroundcolor.setter
    def markerbackgroundcolor(self, value):
        self.MarkerBackgroundColor = value

    @property
    def MarkerBackgroundColorIndex(self):
        return XlColorIndex(self.series.MarkerBackgroundColorIndex)

    # Lower case alias for MarkerBackgroundColorIndex
    @property
    def markerbackgroundcolorindex(self):
        return self.MarkerBackgroundColorIndex

    @MarkerBackgroundColorIndex.setter
    def MarkerBackgroundColorIndex(self, value):
        self.series.MarkerBackgroundColorIndex = value

    # Lower case alias for MarkerBackgroundColorIndex setter
    @markerbackgroundcolorindex.setter
    def markerbackgroundcolorindex(self, value):
        self.MarkerBackgroundColorIndex = value

    @property
    def MarkerForegroundColor(self):
        return self.series.MarkerForegroundColor

    # Lower case alias for MarkerForegroundColor
    @property
    def markerforegroundcolor(self):
        return self.MarkerForegroundColor

    @MarkerForegroundColor.setter
    def MarkerForegroundColor(self, value):
        self.series.MarkerForegroundColor = value

    # Lower case alias for MarkerForegroundColor setter
    @markerforegroundcolor.setter
    def markerforegroundcolor(self, value):
        self.MarkerForegroundColor = value

    @property
    def MarkerForegroundColorIndex(self):
        return XlColorIndex(self.series.MarkerForegroundColorIndex)

    # Lower case alias for MarkerForegroundColorIndex
    @property
    def markerforegroundcolorindex(self):
        return self.MarkerForegroundColorIndex

    @MarkerForegroundColorIndex.setter
    def MarkerForegroundColorIndex(self, value):
        self.series.MarkerForegroundColorIndex = value

    # Lower case alias for MarkerForegroundColorIndex setter
    @markerforegroundcolorindex.setter
    def markerforegroundcolorindex(self, value):
        self.MarkerForegroundColorIndex = value

    @property
    def MarkerSize(self):
        return self.series.MarkerSize

    # Lower case alias for MarkerSize
    @property
    def markersize(self):
        return self.MarkerSize

    @MarkerSize.setter
    def MarkerSize(self, value):
        self.series.MarkerSize = value

    # Lower case alias for MarkerSize setter
    @markersize.setter
    def markersize(self, value):
        self.MarkerSize = value

    @property
    def MarkerStyle(self):
        return XlMarkerStyle(self.series.MarkerStyle)

    # Lower case alias for MarkerStyle
    @property
    def markerstyle(self):
        return self.MarkerStyle

    @MarkerStyle.setter
    def MarkerStyle(self, value):
        self.series.MarkerStyle = value

    # Lower case alias for MarkerStyle setter
    @markerstyle.setter
    def markerstyle(self, value):
        self.MarkerStyle = value

    @property
    def Name(self):
        return self.series.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.series.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.series.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PictureType(self):
        return XlChartPictureType(self.series.PictureType)

    # Lower case alias for PictureType
    @property
    def picturetype(self):
        return self.PictureType

    @PictureType.setter
    def PictureType(self, value):
        self.series.PictureType = value

    # Lower case alias for PictureType setter
    @picturetype.setter
    def picturetype(self, value):
        self.PictureType = value

    @property
    def PictureUnit2(self):
        return self.series.PictureUnit2

    # Lower case alias for PictureUnit2
    @property
    def pictureunit2(self):
        return self.PictureUnit2

    @PictureUnit2.setter
    def PictureUnit2(self, value):
        self.series.PictureUnit2 = value

    # Lower case alias for PictureUnit2 setter
    @pictureunit2.setter
    def pictureunit2(self, value):
        self.PictureUnit2 = value

    @property
    def PlotColorIndex(self):
        return self.series.PlotColorIndex

    # Lower case alias for PlotColorIndex
    @property
    def plotcolorindex(self):
        return self.PlotColorIndex

    @property
    def PlotOrder(self):
        return self.series.PlotOrder

    # Lower case alias for PlotOrder
    @property
    def plotorder(self):
        return self.PlotOrder

    @PlotOrder.setter
    def PlotOrder(self, value):
        self.series.PlotOrder = value

    # Lower case alias for PlotOrder setter
    @plotorder.setter
    def plotorder(self, value):
        self.PlotOrder = value

    @property
    def Shadow(self):
        return self.series.Shadow

    # Lower case alias for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.series.Shadow = value

    # Lower case alias for Shadow setter
    @shadow.setter
    def shadow(self, value):
        self.Shadow = value

    @property
    def Smooth(self):
        return self.series.Smooth

    # Lower case alias for Smooth
    @property
    def smooth(self):
        return self.Smooth

    @Smooth.setter
    def Smooth(self, value):
        self.series.Smooth = value

    # Lower case alias for Smooth setter
    @smooth.setter
    def smooth(self, value):
        self.Smooth = value

    @property
    def Type(self):
        return self.series.Type

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @Type.setter
    def Type(self, value):
        self.series.Type = value

    # Lower case alias for Type setter
    @type.setter
    def type(self, value):
        self.Type = value

    @property
    def Values(self):
        return self.series.Values

    # Lower case alias for Values
    @property
    def values(self):
        return self.Values

    @Values.setter
    def Values(self, value):
        self.series.Values = value

    # Lower case alias for Values setter
    @values.setter
    def values(self, value):
        self.Values = value

    @property
    def XValues(self):
        return self.series.XValues

    # Lower case alias for XValues
    @property
    def xvalues(self):
        return self.XValues

    @XValues.setter
    def XValues(self, value):
        self.series.XValues = value

    # Lower case alias for XValues setter
    @xvalues.setter
    def xvalues(self, value):
        self.XValues = value

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.seriescollection.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Parent(self):
        return self.seriescollection.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

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

    # Lower case alias for Border
    @property
    def border(self):
        return self.Border

    @property
    def Creator(self):
        return self.serieslines.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.serieslines.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def Name(self):
        return self.serieslines.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.serieslines.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

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

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Property(self):
        return self.seteffect.Property

    @Property.setter
    def Property(self, value):
        self.seteffect.Property = value

    @property
    def To(self):
        return self.seteffect.To

    # Lower case alias for To
    @property
    def to(self):
        return self.To

    @To.setter
    def To(self, value):
        self.seteffect.To = value

    # Lower case alias for To setter
    @to.setter
    def to(self, value):
        self.To = value


class ShadowFormat:

    def __init__(self, shadowformat=None):
        self.shadowformat = shadowformat

    @property
    def Application(self):
        return Application(self.shadowformat.Application)

    @property
    def Blur(self):
        return self.shadowformat.Blur

    # Lower case alias for Blur
    @property
    def blur(self):
        return self.Blur

    @Blur.setter
    def Blur(self, value):
        self.shadowformat.Blur = value

    # Lower case alias for Blur setter
    @blur.setter
    def blur(self, value):
        self.Blur = value

    @property
    def Creator(self):
        return self.shadowformat.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def ForeColor(self):
        return ColorFormat(self.shadowformat.ForeColor)

    # Lower case alias for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.shadowformat.ForeColor = value

    # Lower case alias for ForeColor setter
    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def Obscured(self):
        return self.shadowformat.Obscured

    # Lower case alias for Obscured
    @property
    def obscured(self):
        return self.Obscured

    @Obscured.setter
    def Obscured(self, value):
        self.shadowformat.Obscured = value

    # Lower case alias for Obscured setter
    @obscured.setter
    def obscured(self, value):
        self.Obscured = value

    @property
    def OffsetX(self):
        return self.shadowformat.OffsetX

    # Lower case alias for OffsetX
    @property
    def offsetx(self):
        return self.OffsetX

    @OffsetX.setter
    def OffsetX(self, value):
        self.shadowformat.OffsetX = value

    # Lower case alias for OffsetX setter
    @offsetx.setter
    def offsetx(self, value):
        self.OffsetX = value

    @property
    def OffsetY(self):
        return self.shadowformat.OffsetY

    # Lower case alias for OffsetY
    @property
    def offsety(self):
        return self.OffsetY

    @OffsetY.setter
    def OffsetY(self, value):
        self.shadowformat.OffsetY = value

    # Lower case alias for OffsetY setter
    @offsety.setter
    def offsety(self, value):
        self.OffsetY = value

    @property
    def Parent(self):
        return self.shadowformat.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RotateWithShape(self):
        return self.shadowformat.RotateWithShape

    # Lower case alias for RotateWithShape
    @property
    def rotatewithshape(self):
        return self.RotateWithShape

    @RotateWithShape.setter
    def RotateWithShape(self, value):
        self.shadowformat.RotateWithShape = value

    # Lower case alias for RotateWithShape setter
    @rotatewithshape.setter
    def rotatewithshape(self, value):
        self.RotateWithShape = value

    @property
    def Size(self):
        return self.shadowformat.Size

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @Size.setter
    def Size(self, value):
        self.shadowformat.Size = value

    # Lower case alias for Size setter
    @size.setter
    def size(self, value):
        self.Size = value

    @property
    def Style(self):
        return self.shadowformat.Style

    # Lower case alias for Style
    @property
    def style(self):
        return self.Style

    @Style.setter
    def Style(self, value):
        self.shadowformat.Style = value

    # Lower case alias for Style setter
    @style.setter
    def style(self, value):
        self.Style = value

    @property
    def Transparency(self):
        return self.shadowformat.Transparency

    # Lower case alias for Transparency
    @property
    def transparency(self):
        return self.Transparency

    @Transparency.setter
    def Transparency(self, value):
        self.shadowformat.Transparency = value

    # Lower case alias for Transparency setter
    @transparency.setter
    def transparency(self, value):
        self.Transparency = value

    @property
    def Type(self):
        return self.shadowformat.Type

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @Type.setter
    def Type(self, value):
        self.shadowformat.Type = value

    # Lower case alias for Type setter
    @type.setter
    def type(self, value):
        self.Type = value

    @property
    def Visible(self):
        return self.shadowformat.Visible

    # Lower case alias for Visible
    @property
    def visible(self):
        return self.Visible

    @Visible.setter
    def Visible(self, value):
        self.shadowformat.Visible = value

    # Lower case alias for Visible setter
    @visible.setter
    def visible(self, value):
        self.Visible = value

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
    def ActionSettings(self):
        return ActionSettings(self.shape.ActionSettings)

    # Lower case alias for ActionSettings
    @property
    def actionsettings(self):
        return self.ActionSettings

    @property
    def Adjustments(self):
        return Adjustments(self.shape.Adjustments)

    # Lower case alias for Adjustments
    @property
    def adjustments(self):
        return self.Adjustments

    @property
    def AlternativeText(self):
        return self.shape.AlternativeText

    # Lower case alias for AlternativeText
    @property
    def alternativetext(self):
        return self.AlternativeText

    @AlternativeText.setter
    def AlternativeText(self, value):
        self.shape.AlternativeText = value

    # Lower case alias for AlternativeText setter
    @alternativetext.setter
    def alternativetext(self, value):
        self.AlternativeText = value

    @property
    def AnimationSettings(self):
        return AnimationSettings(self.shape.AnimationSettings)

    # Lower case alias for AnimationSettings
    @property
    def animationsettings(self):
        return self.AnimationSettings

    @property
    def Application(self):
        return Application(self.shape.Application)

    @property
    def AutoShapeType(self):
        return Shape(self.shape.AutoShapeType)

    # Lower case alias for AutoShapeType
    @property
    def autoshapetype(self):
        return self.AutoShapeType

    @AutoShapeType.setter
    def AutoShapeType(self, value):
        self.shape.AutoShapeType = value

    # Lower case alias for AutoShapeType setter
    @autoshapetype.setter
    def autoshapetype(self, value):
        self.AutoShapeType = value

    @property
    def BackgroundStyle(self):
        return self.shape.BackgroundStyle

    # Lower case alias for BackgroundStyle
    @property
    def backgroundstyle(self):
        return self.BackgroundStyle

    @BackgroundStyle.setter
    def BackgroundStyle(self, value):
        self.shape.BackgroundStyle = value

    # Lower case alias for BackgroundStyle setter
    @backgroundstyle.setter
    def backgroundstyle(self, value):
        self.BackgroundStyle = value

    @property
    def BlackWhiteMode(self):
        return self.shape.BlackWhiteMode

    # Lower case alias for BlackWhiteMode
    @property
    def blackwhitemode(self):
        return self.BlackWhiteMode

    @BlackWhiteMode.setter
    def BlackWhiteMode(self, value):
        self.shape.BlackWhiteMode = value

    # Lower case alias for BlackWhiteMode setter
    @blackwhitemode.setter
    def blackwhitemode(self, value):
        self.BlackWhiteMode = value

    @property
    def Callout(self):
        return CalloutFormat(self.shape.Callout)

    # Lower case alias for Callout
    @property
    def callout(self):
        return self.Callout

    @property
    def Chart(self):
        return Chart(self.shape.Chart)

    # Lower case alias for Chart
    @property
    def chart(self):
        return self.Chart

    @property
    def Child(self):
        return self.shape.Child

    # Lower case alias for Child
    @property
    def child(self):
        return self.Child

    @property
    def ConnectionSiteCount(self):
        return self.shape.ConnectionSiteCount

    # Lower case alias for ConnectionSiteCount
    @property
    def connectionsitecount(self):
        return self.ConnectionSiteCount

    @property
    def Connector(self):
        return self.shape.Connector

    # Lower case alias for Connector
    @property
    def connector(self):
        return self.Connector

    @property
    def ConnectorFormat(self):
        return ConnectorFormat(self.shape.ConnectorFormat)

    # Lower case alias for ConnectorFormat
    @property
    def connectorformat(self):
        return self.ConnectorFormat

    @property
    def Creator(self):
        return self.shape.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def CustomerData(self):
        return CustomerData(self.shape.CustomerData)

    # Lower case alias for CustomerData
    @property
    def customerdata(self):
        return self.CustomerData

    @property
    def Decorative(self):
        return self.shape.Decorative

    # Lower case alias for Decorative
    @property
    def decorative(self):
        return self.Decorative

    @Decorative.setter
    def Decorative(self, value):
        self.shape.Decorative = value

    # Lower case alias for Decorative setter
    @decorative.setter
    def decorative(self, value):
        self.Decorative = value

    @property
    def Fill(self):
        return FillFormat(self.shape.Fill)

    # Lower case alias for Fill
    @property
    def fill(self):
        return self.Fill

    @property
    def Glow(self):
        return self.shape.Glow

    # Lower case alias for Glow
    @property
    def glow(self):
        return self.Glow

    @property
    def GraphicStyle(self):
        return self.shape.GraphicStyle

    # Lower case alias for GraphicStyle
    @property
    def graphicstyle(self):
        return self.GraphicStyle

    @GraphicStyle.setter
    def GraphicStyle(self, value):
        self.shape.GraphicStyle = value

    # Lower case alias for GraphicStyle setter
    @graphicstyle.setter
    def graphicstyle(self, value):
        self.GraphicStyle = value

    @property
    def GroupItems(self):
        return GroupShapes(self.shape.GroupItems)

    # Lower case alias for GroupItems
    @property
    def groupitems(self):
        return self.GroupItems

    @property
    def HasChart(self):
        return self.shape.HasChart

    # Lower case alias for HasChart
    @property
    def haschart(self):
        return self.HasChart

    @property
    def HasSmartArt(self):
        return self.shape.HasSmartArt

    # Lower case alias for HasSmartArt
    @property
    def hassmartart(self):
        return self.HasSmartArt

    @property
    def HasTable(self):
        return self.shape.HasTable

    # Lower case alias for HasTable
    @property
    def hastable(self):
        return self.HasTable

    @property
    def HasTextFrame(self):
        return self.shape.HasTextFrame

    # Lower case alias for HasTextFrame
    @property
    def hastextframe(self):
        return self.HasTextFrame

    @property
    def Height(self):
        return self.shape.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @Height.setter
    def Height(self, value):
        self.shape.Height = value

    # Lower case alias for Height setter
    @height.setter
    def height(self, value):
        self.Height = value

    @property
    def HorizontalFlip(self):
        return self.shape.HorizontalFlip

    # Lower case alias for HorizontalFlip
    @property
    def horizontalflip(self):
        return self.HorizontalFlip

    @property
    def Id(self):
        return self.shape.Id

    # Lower case alias for Id
    @property
    def id(self):
        return self.Id

    @property
    def Left(self):
        return self.shape.Left

    # Lower case alias for Left
    @property
    def left(self):
        return self.Left

    @Left.setter
    def Left(self, value):
        self.shape.Left = value

    # Lower case alias for Left setter
    @left.setter
    def left(self, value):
        self.Left = value

    @property
    def Line(self):
        return LineFormat(self.shape.Line)

    # Lower case alias for Line
    @property
    def line(self):
        return self.Line

    @property
    def LinkFormat(self):
        return LinkFormat(self.shape.LinkFormat)

    # Lower case alias for LinkFormat
    @property
    def linkformat(self):
        return self.LinkFormat

    @property
    def LockAspectRatio(self):
        return self.shape.LockAspectRatio

    # Lower case alias for LockAspectRatio
    @property
    def lockaspectratio(self):
        return self.LockAspectRatio

    @LockAspectRatio.setter
    def LockAspectRatio(self, value):
        self.shape.LockAspectRatio = value

    # Lower case alias for LockAspectRatio setter
    @lockaspectratio.setter
    def lockaspectratio(self, value):
        self.LockAspectRatio = value

    @property
    def MediaFormat(self):
        return self.shape.MediaFormat

    # Lower case alias for MediaFormat
    @property
    def mediaformat(self):
        return self.MediaFormat

    @property
    def MediaType(self):
        return self.shape.MediaType

    # Lower case alias for MediaType
    @property
    def mediatype(self):
        return self.MediaType

    @property
    def Model3D(self):
        return Model3DFormat(self.shape.Model3D)

    # Lower case alias for Model3D
    @property
    def model3d(self):
        return self.Model3D

    @property
    def Name(self):
        return self.shape.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.shape.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Nodes(self):
        return ShapeNodes(self.shape.Nodes)

    # Lower case alias for Nodes
    @property
    def nodes(self):
        return self.Nodes

    @property
    def OLEFormat(self):
        return OLEFormat(self.shape.OLEFormat)

    # Lower case alias for OLEFormat
    @property
    def oleformat(self):
        return self.OLEFormat

    @property
    def Parent(self):
        return self.shape.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ParentGroup(self):
        return Shape(self.shape.ParentGroup)

    # Lower case alias for ParentGroup
    @property
    def parentgroup(self):
        return self.ParentGroup

    @property
    def PictureFormat(self):
        return PictureFormat(self.shape.PictureFormat)

    # Lower case alias for PictureFormat
    @property
    def pictureformat(self):
        return self.PictureFormat

    @property
    def PlaceholderFormat(self):
        return PlaceholderFormat(self.shape.PlaceholderFormat)

    # Lower case alias for PlaceholderFormat
    @property
    def placeholderformat(self):
        return self.PlaceholderFormat

    @property
    def Reflection(self):
        return self.shape.Reflection

    # Lower case alias for Reflection
    @property
    def reflection(self):
        return self.Reflection

    @property
    def Rotation(self):
        return self.shape.Rotation

    # Lower case alias for Rotation
    @property
    def rotation(self):
        return self.Rotation

    @Rotation.setter
    def Rotation(self, value):
        self.shape.Rotation = value

    # Lower case alias for Rotation setter
    @rotation.setter
    def rotation(self, value):
        self.Rotation = value

    @property
    def Shadow(self):
        return ShadowFormat(self.shape.Shadow)

    # Lower case alias for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @property
    def ShapeStyle(self):
        return self.shape.ShapeStyle

    # Lower case alias for ShapeStyle
    @property
    def shapestyle(self):
        return self.ShapeStyle

    @ShapeStyle.setter
    def ShapeStyle(self, value):
        self.shape.ShapeStyle = value

    # Lower case alias for ShapeStyle setter
    @shapestyle.setter
    def shapestyle(self, value):
        self.ShapeStyle = value

    @property
    def SmartArt(self):
        return Shape(self.shape.SmartArt)

    # Lower case alias for SmartArt
    @property
    def smartart(self):
        return self.SmartArt

    @property
    def SoftEdge(self):
        return self.shape.SoftEdge

    # Lower case alias for SoftEdge
    @property
    def softedge(self):
        return self.SoftEdge

    @property
    def Table(self):
        return Table(self.shape.Table)

    # Lower case alias for Table
    @property
    def table(self):
        return self.Table

    @property
    def Tags(self):
        return Tags(self.shape.Tags)

    # Lower case alias for Tags
    @property
    def tags(self):
        return self.Tags

    @property
    def TextEffect(self):
        return TextEffectFormat(self.shape.TextEffect)

    # Lower case alias for TextEffect
    @property
    def texteffect(self):
        return self.TextEffect

    @property
    def TextFrame(self):
        return TextFrame(self.shape.TextFrame)

    # Lower case alias for TextFrame
    @property
    def textframe(self):
        return self.TextFrame

    @property
    def TextFrame2(self):
        return TextFrame2(self.shape.TextFrame2)

    # Lower case alias for TextFrame2
    @property
    def textframe2(self):
        return self.TextFrame2

    @property
    def ThreeD(self):
        return ThreeDFormat(self.shape.ThreeD)

    # Lower case alias for ThreeD
    @property
    def threed(self):
        return self.ThreeD

    @property
    def Title(self):
        return Shape(self.shape.Title)

    # Lower case alias for Title
    @property
    def title(self):
        return self.Title

    @property
    def Top(self):
        return self.shape.Top

    # Lower case alias for Top
    @property
    def top(self):
        return self.Top

    @Top.setter
    def Top(self, value):
        self.shape.Top = value

    # Lower case alias for Top setter
    @top.setter
    def top(self, value):
        self.Top = value

    @property
    def Type(self):
        return self.shape.Type

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @property
    def VerticalFlip(self):
        return self.shape.VerticalFlip

    # Lower case alias for VerticalFlip
    @property
    def verticalflip(self):
        return self.VerticalFlip

    @property
    def Vertices(self):
        return self.shape.Vertices

    # Lower case alias for Vertices
    @property
    def vertices(self):
        return self.Vertices

    @property
    def Visible(self):
        return self.shape.Visible

    # Lower case alias for Visible
    @property
    def visible(self):
        return self.Visible

    @Visible.setter
    def Visible(self, value):
        self.shape.Visible = value

    # Lower case alias for Visible setter
    @visible.setter
    def visible(self, value):
        self.Visible = value

    @property
    def Width(self):
        return self.shape.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

    @Width.setter
    def Width(self, value):
        self.shape.Width = value

    # Lower case alias for Width setter
    @width.setter
    def width(self, value):
        self.Width = value

    @property
    def ZOrderPosition(self):
        return self.shape.ZOrderPosition

    # Lower case alias for ZOrderPosition
    @property
    def zorderposition(self):
        return self.ZOrderPosition

    def Apply(self):
        self.shape.Apply()

    def ApplyAnimation(self):
        self.shape.ApplyAnimation()

    def ConvertTextToSmartArt(self, Layout=None):
        arguments = com_arguments([Layout])
        self.shape.ConvertTextToSmartArt(*arguments)

    def Copy(self):
        self.shape.Copy()

    def Cut(self):
        self.shape.Cut()

    def Delete(self):
        self.shape.Delete()

    def Duplicate(self):
        return self.shape.Duplicate()

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

    def PickupAnimation(self):
        self.shape.PickupAnimation()

    def RerouteConnections(self):
        self.shape.RerouteConnections()

    def ScaleHeight(self, Factor=None, RelativeToOriginalSize=None, fScale=None):
        arguments = com_arguments([Factor, RelativeToOriginalSize, fScale])
        self.shape.ScaleHeight(*arguments)

    def ScaleWidth(self, Factor=None, RelativeToOriginalSize=None, fScale=None):
        arguments = com_arguments([Factor, RelativeToOriginalSize, fScale])
        self.shape.ScaleWidth(*arguments)

    def Select(self, Replace=None):
        arguments = com_arguments([Replace])
        self.shape.Select(*arguments)

    def SetShapesDefaultProperties(self):
        self.shape.SetShapesDefaultProperties()

    def Ungroup(self):
        return self.shape.Ungroup()

    def UpgradeMedia(self):
        self.shape.UpgradeMedia()

    def ZOrder(self, ZOrderCmd=None):
        arguments = com_arguments([ZOrderCmd])
        self.shape.ZOrder(*arguments)


class ShapeNode:

    def __init__(self, shapenode=None):
        self.shapenode = shapenode

    @property
    def Application(self):
        return Application(self.shapenode.Application)

    @property
    def Creator(self):
        return self.shapenode.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def EditingType(self):
        return self.shapenode.EditingType

    # Lower case alias for EditingType
    @property
    def editingtype(self):
        return self.EditingType

    @property
    def Parent(self):
        return self.shapenode.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Points(self):
        return self.shapenode.Points

    # Lower case alias for Points
    @property
    def points(self):
        return self.Points

    @property
    def SegmentType(self):
        return self.shapenode.SegmentType

    # Lower case alias for SegmentType
    @property
    def segmenttype(self):
        return self.SegmentType


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.shapenodes.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Parent(self):
        return self.shapenodes.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Delete(self, Index=None):
        arguments = com_arguments([Index])
        self.shapenodes.Delete(*arguments)

    def Insert(self, Index=None, SegmentType=None, EditingType=None, X1=None, Y1=None, X2=None, Y2=None, X3=None, Y3=None):
        arguments = com_arguments([Index, SegmentType, EditingType, X1, Y1, X2, Y2, X3, Y3])
        self.shapenodes.Insert(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.shapenodes.Item(*arguments)

    def SetEditingType(self, Index=None, EditingType=None):
        arguments = com_arguments([Index, EditingType])
        self.shapenodes.SetEditingType(*arguments)

    def SetPosition(self, Index=None, X1=None, Y1=None):
        arguments = com_arguments([Index, X1, Y1])
        self.shapenodes.SetPosition(*arguments)

    def SetSegmentType(self, Index=None, SegmentType=None):
        arguments = com_arguments([Index, SegmentType])
        self.shapenodes.SetSegmentType(*arguments)


class ShapeRange:

    def __init__(self, shaperange=None):
        self.shaperange = shaperange

    @property
    def ActionSettings(self):
        return ActionSettings(self.shaperange.ActionSettings)

    # Lower case alias for ActionSettings
    @property
    def actionsettings(self):
        return self.ActionSettings

    @property
    def Adjustments(self):
        return Adjustments(self.shaperange.Adjustments)

    # Lower case alias for Adjustments
    @property
    def adjustments(self):
        return self.Adjustments

    @property
    def AlternativeText(self):
        return self.shaperange.AlternativeText

    # Lower case alias for AlternativeText
    @property
    def alternativetext(self):
        return self.AlternativeText

    @AlternativeText.setter
    def AlternativeText(self, value):
        self.shaperange.AlternativeText = value

    # Lower case alias for AlternativeText setter
    @alternativetext.setter
    def alternativetext(self, value):
        self.AlternativeText = value

    @property
    def AnimationSettings(self):
        return AnimationSettings(self.shaperange.AnimationSettings)

    # Lower case alias for AnimationSettings
    @property
    def animationsettings(self):
        return self.AnimationSettings

    @property
    def Application(self):
        return Application(self.shaperange.Application)

    @property
    def AutoShapeType(self):
        return ShapeRange(self.shaperange.AutoShapeType)

    # Lower case alias for AutoShapeType
    @property
    def autoshapetype(self):
        return self.AutoShapeType

    @AutoShapeType.setter
    def AutoShapeType(self, value):
        self.shaperange.AutoShapeType = value

    # Lower case alias for AutoShapeType setter
    @autoshapetype.setter
    def autoshapetype(self, value):
        self.AutoShapeType = value

    @property
    def BackgroundStyle(self):
        return self.shaperange.BackgroundStyle

    # Lower case alias for BackgroundStyle
    @property
    def backgroundstyle(self):
        return self.BackgroundStyle

    @BackgroundStyle.setter
    def BackgroundStyle(self, value):
        self.shaperange.BackgroundStyle = value

    # Lower case alias for BackgroundStyle setter
    @backgroundstyle.setter
    def backgroundstyle(self, value):
        self.BackgroundStyle = value

    @property
    def BlackWhiteMode(self):
        return self.shaperange.BlackWhiteMode

    # Lower case alias for BlackWhiteMode
    @property
    def blackwhitemode(self):
        return self.BlackWhiteMode

    @BlackWhiteMode.setter
    def BlackWhiteMode(self, value):
        self.shaperange.BlackWhiteMode = value

    # Lower case alias for BlackWhiteMode setter
    @blackwhitemode.setter
    def blackwhitemode(self, value):
        self.BlackWhiteMode = value

    @property
    def Callout(self):
        return CalloutFormat(self.shaperange.Callout)

    # Lower case alias for Callout
    @property
    def callout(self):
        return self.Callout

    @property
    def Chart(self):
        return Chart(self.shaperange.Chart)

    # Lower case alias for Chart
    @property
    def chart(self):
        return self.Chart

    @property
    def Child(self):
        return self.shaperange.Child

    # Lower case alias for Child
    @property
    def child(self):
        return self.Child

    @property
    def ConnectionSiteCount(self):
        return self.shaperange.ConnectionSiteCount

    # Lower case alias for ConnectionSiteCount
    @property
    def connectionsitecount(self):
        return self.ConnectionSiteCount

    @property
    def Connector(self):
        return self.shaperange.Connector

    # Lower case alias for Connector
    @property
    def connector(self):
        return self.Connector

    @property
    def ConnectorFormat(self):
        return ConnectorFormat(self.shaperange.ConnectorFormat)

    # Lower case alias for ConnectorFormat
    @property
    def connectorformat(self):
        return self.ConnectorFormat

    @property
    def Count(self):
        return self.shaperange.Count

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.shaperange.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def CustomerData(self):
        return CustomerData(self.shaperange.CustomerData)

    # Lower case alias for CustomerData
    @property
    def customerdata(self):
        return self.CustomerData

    @property
    def Decorative(self):
        return self.shaperange.Decorative

    # Lower case alias for Decorative
    @property
    def decorative(self):
        return self.Decorative

    @Decorative.setter
    def Decorative(self, value):
        self.shaperange.Decorative = value

    # Lower case alias for Decorative setter
    @decorative.setter
    def decorative(self, value):
        self.Decorative = value

    @property
    def Fill(self):
        return FillFormat(self.shaperange.Fill)

    # Lower case alias for Fill
    @property
    def fill(self):
        return self.Fill

    @property
    def Glow(self):
        return self.shaperange.Glow

    # Lower case alias for Glow
    @property
    def glow(self):
        return self.Glow

    @property
    def GraphicStyle(self):
        return self.shaperange.GraphicStyle

    # Lower case alias for GraphicStyle
    @property
    def graphicstyle(self):
        return self.GraphicStyle

    @GraphicStyle.setter
    def GraphicStyle(self, value):
        self.shaperange.GraphicStyle = value

    # Lower case alias for GraphicStyle setter
    @graphicstyle.setter
    def graphicstyle(self, value):
        self.GraphicStyle = value

    @property
    def GroupItems(self):
        return GroupShapes(self.shaperange.GroupItems)

    # Lower case alias for GroupItems
    @property
    def groupitems(self):
        return self.GroupItems

    @property
    def HasChart(self):
        return self.shaperange.HasChart

    # Lower case alias for HasChart
    @property
    def haschart(self):
        return self.HasChart

    @property
    def HasSmartArt(self):
        return self.shaperange.HasSmartArt

    # Lower case alias for HasSmartArt
    @property
    def hassmartart(self):
        return self.HasSmartArt

    @property
    def HasTable(self):
        return self.shaperange.HasTable

    # Lower case alias for HasTable
    @property
    def hastable(self):
        return self.HasTable

    @property
    def HasTextFrame(self):
        return self.shaperange.HasTextFrame

    # Lower case alias for HasTextFrame
    @property
    def hastextframe(self):
        return self.HasTextFrame

    @property
    def Height(self):
        return self.shaperange.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @Height.setter
    def Height(self, value):
        self.shaperange.Height = value

    # Lower case alias for Height setter
    @height.setter
    def height(self, value):
        self.Height = value

    @property
    def HorizontalFlip(self):
        return self.shaperange.HorizontalFlip

    # Lower case alias for HorizontalFlip
    @property
    def horizontalflip(self):
        return self.HorizontalFlip

    @property
    def Id(self):
        return self.shaperange.Id

    # Lower case alias for Id
    @property
    def id(self):
        return self.Id

    @property
    def Left(self):
        return self.shaperange.Left

    # Lower case alias for Left
    @property
    def left(self):
        return self.Left

    @Left.setter
    def Left(self, value):
        self.shaperange.Left = value

    # Lower case alias for Left setter
    @left.setter
    def left(self, value):
        self.Left = value

    @property
    def Line(self):
        return LineFormat(self.shaperange.Line)

    # Lower case alias for Line
    @property
    def line(self):
        return self.Line

    @property
    def LinkFormat(self):
        return LinkFormat(self.shaperange.LinkFormat)

    # Lower case alias for LinkFormat
    @property
    def linkformat(self):
        return self.LinkFormat

    @property
    def LockAspectRatio(self):
        return self.shaperange.LockAspectRatio

    # Lower case alias for LockAspectRatio
    @property
    def lockaspectratio(self):
        return self.LockAspectRatio

    @LockAspectRatio.setter
    def LockAspectRatio(self, value):
        self.shaperange.LockAspectRatio = value

    # Lower case alias for LockAspectRatio setter
    @lockaspectratio.setter
    def lockaspectratio(self, value):
        self.LockAspectRatio = value

    @property
    def MediaFormat(self):
        return MediaFormat(self.shaperange.MediaFormat)

    # Lower case alias for MediaFormat
    @property
    def mediaformat(self):
        return self.MediaFormat

    @property
    def MediaType(self):
        return self.shaperange.MediaType

    # Lower case alias for MediaType
    @property
    def mediatype(self):
        return self.MediaType

    @property
    def Model3D(self):
        return Model3DFormat(self.shaperange.Model3D)

    # Lower case alias for Model3D
    @property
    def model3d(self):
        return self.Model3D

    @property
    def Name(self):
        return self.shaperange.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.shaperange.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Nodes(self):
        return ShapeNodes(self.shaperange.Nodes)

    # Lower case alias for Nodes
    @property
    def nodes(self):
        return self.Nodes

    @property
    def OLEFormat(self):
        return OLEFormat(self.shaperange.OLEFormat)

    # Lower case alias for OLEFormat
    @property
    def oleformat(self):
        return self.OLEFormat

    @property
    def Parent(self):
        return self.shaperange.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ParentGroup(self):
        return Shape(self.shaperange.ParentGroup)

    # Lower case alias for ParentGroup
    @property
    def parentgroup(self):
        return self.ParentGroup

    @property
    def PictureFormat(self):
        return PictureFormat(self.shaperange.PictureFormat)

    # Lower case alias for PictureFormat
    @property
    def pictureformat(self):
        return self.PictureFormat

    @property
    def PlaceholderFormat(self):
        return PlaceholderFormat(self.shaperange.PlaceholderFormat)

    # Lower case alias for PlaceholderFormat
    @property
    def placeholderformat(self):
        return self.PlaceholderFormat

    @property
    def Reflection(self):
        return self.shaperange.Reflection

    # Lower case alias for Reflection
    @property
    def reflection(self):
        return self.Reflection

    @property
    def Rotation(self):
        return self.shaperange.Rotation

    # Lower case alias for Rotation
    @property
    def rotation(self):
        return self.Rotation

    @Rotation.setter
    def Rotation(self, value):
        self.shaperange.Rotation = value

    # Lower case alias for Rotation setter
    @rotation.setter
    def rotation(self, value):
        self.Rotation = value

    @property
    def Shadow(self):
        return ShadowFormat(self.shaperange.Shadow)

    # Lower case alias for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @property
    def ShapeStyle(self):
        return self.shaperange.ShapeStyle

    # Lower case alias for ShapeStyle
    @property
    def shapestyle(self):
        return self.ShapeStyle

    @property
    def SmartArt(self):
        return ShapeRange(self.shaperange.SmartArt)

    # Lower case alias for SmartArt
    @property
    def smartart(self):
        return self.SmartArt

    @property
    def SoftEdge(self):
        return self.shaperange.SoftEdge

    # Lower case alias for SoftEdge
    @property
    def softedge(self):
        return self.SoftEdge

    @property
    def Table(self):
        return Table(self.shaperange.Table)

    # Lower case alias for Table
    @property
    def table(self):
        return self.Table

    @property
    def Tags(self):
        return Tags(self.shaperange.Tags)

    # Lower case alias for Tags
    @property
    def tags(self):
        return self.Tags

    @property
    def TextEffect(self):
        return TextEffectFormat(self.shaperange.TextEffect)

    # Lower case alias for TextEffect
    @property
    def texteffect(self):
        return self.TextEffect

    @property
    def TextFrame(self):
        return TextFrame(self.shaperange.TextFrame)

    # Lower case alias for TextFrame
    @property
    def textframe(self):
        return self.TextFrame

    @property
    def TextFrame2(self):
        return TextFrame2(self.shaperange.TextFrame2)

    # Lower case alias for TextFrame2
    @property
    def textframe2(self):
        return self.TextFrame2

    @property
    def ThreeD(self):
        return ThreeDFormat(self.shaperange.ThreeD)

    # Lower case alias for ThreeD
    @property
    def threed(self):
        return self.ThreeD

    @property
    def Title(self):
        return Shape(self.shaperange.Title)

    # Lower case alias for Title
    @property
    def title(self):
        return self.Title

    @property
    def Top(self):
        return self.shaperange.Top

    # Lower case alias for Top
    @property
    def top(self):
        return self.Top

    @Top.setter
    def Top(self, value):
        self.shaperange.Top = value

    # Lower case alias for Top setter
    @top.setter
    def top(self, value):
        self.Top = value

    @property
    def Type(self):
        return self.shaperange.Type

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @property
    def VerticalFlip(self):
        return self.shaperange.VerticalFlip

    # Lower case alias for VerticalFlip
    @property
    def verticalflip(self):
        return self.VerticalFlip

    @property
    def Vertices(self):
        return self.shaperange.Vertices

    # Lower case alias for Vertices
    @property
    def vertices(self):
        return self.Vertices

    @property
    def Visible(self):
        return self.shaperange.Visible

    # Lower case alias for Visible
    @property
    def visible(self):
        return self.Visible

    @Visible.setter
    def Visible(self, value):
        self.shaperange.Visible = value

    # Lower case alias for Visible setter
    @visible.setter
    def visible(self, value):
        self.Visible = value

    @property
    def Width(self):
        return self.shaperange.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

    @Width.setter
    def Width(self, value):
        self.shaperange.Width = value

    # Lower case alias for Width setter
    @width.setter
    def width(self, value):
        self.Width = value

    @property
    def ZOrderPosition(self):
        return self.shaperange.ZOrderPosition

    # Lower case alias for ZOrderPosition
    @property
    def zorderposition(self):
        return self.ZOrderPosition

    def Align(self, AlignCmd=None, RelativeTo=None):
        arguments = com_arguments([AlignCmd, RelativeTo])
        self.shaperange.Align(*arguments)

    def Apply(self):
        self.shaperange.Apply()

    def ApplyAnimation(self):
        self.shaperange.ApplyAnimation()

    def ConvertTextToSmartArt(self, Layout=None):
        arguments = com_arguments([Layout])
        return self.shaperange.ConvertTextToSmartArt(*arguments)

    def Copy(self):
        self.shaperange.Copy()

    def Cut(self):
        self.shaperange.Cut()

    def Delete(self):
        self.shaperange.Delete()

    def Distribute(self, DistributeCmd=None, RelativeTo=None):
        arguments = com_arguments([DistributeCmd, RelativeTo])
        return self.shaperange.Distribute(*arguments)

    def Duplicate(self):
        return self.shaperange.Duplicate()

    def Flip(self, FlipCmd=None):
        arguments = com_arguments([FlipCmd])
        self.shaperange.Flip(*arguments)

    def Group(self):
        return self.shaperange.Group()

    def IncrementLeft(self, Increment=None):
        arguments = com_arguments([Increment])
        self.shaperange.IncrementLeft(*arguments)

    def IncrementRotation(self, Increment=None):
        arguments = com_arguments([Increment])
        self.shaperange.IncrementRotation(*arguments)

    def IncrementTop(self, Increment=None):
        arguments = com_arguments([Increment])
        self.shaperange.IncrementTop(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.shaperange.Item(*arguments)

    def PickUp(self):
        self.shaperange.PickUp()

    def PickupAnimation(self):
        self.shaperange.PickupAnimation()

    def Regroup(self):
        return self.shaperange.Regroup()

    def RerouteConnections(self):
        self.shaperange.RerouteConnections()

    def ScaleHeight(self, Factor=None, RelativeToOriginalSize=None, fScale=None):
        arguments = com_arguments([Factor, RelativeToOriginalSize, fScale])
        return self.shaperange.ScaleHeight(*arguments)

    def ScaleWidth(self, Factor=None, RelativeToOriginalSize=None, fScale=None):
        arguments = com_arguments([Factor, RelativeToOriginalSize, fScale])
        self.shaperange.ScaleWidth(*arguments)

    def Select(self, Replace=None):
        arguments = com_arguments([Replace])
        self.shaperange.Select(*arguments)

    def SetShapesDefaultProperties(self):
        self.shaperange.SetShapesDefaultProperties()

    def Ungroup(self):
        return self.shaperange.Ungroup()

    def UpgradeMedia(self):
        self.shaperange.UpgradeMedia()

    def ZOrder(self, ZOrderCmd=None):
        arguments = com_arguments([ZOrderCmd])
        self.shaperange.ZOrder(*arguments)


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.shapes.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def HasTitle(self):
        return self.shapes.HasTitle

    # Lower case alias for HasTitle
    @property
    def hastitle(self):
        return self.HasTitle

    @property
    def Parent(self):
        return self.shapes.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Placeholders(self):
        return Placeholders(self.shapes.Placeholders)

    # Lower case alias for Placeholders
    @property
    def placeholders(self):
        return self.Placeholders

    @property
    def Title(self):
        return Shape(self.shapes.Title)

    # Lower case alias for Title
    @property
    def title(self):
        return self.Title

    def Add3DModel(self, FileName=None, LinkToFile=None, SaveWithDocument=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height])
        return self.shapes.Add3DModel(*arguments)

    def AddCallout(self, Type=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([Type, Left, Top, Width, Height])
        return self.shapes.AddCallout(*arguments)

    def AddConnector(self, Type=None, BeginX=None, BeginY=None, EndX=None, EndY=None):
        arguments = com_arguments([Type, BeginX, BeginY, EndX, EndY])
        return self.shapes.AddConnector(*arguments)

    def AddCurve(self, SafeArrayOfPoints=None):
        arguments = com_arguments([SafeArrayOfPoints])
        return self.shapes.AddCurve(*arguments)

    def AddLabel(self, Orientation=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([Orientation, Left, Top, Width, Height])
        return self.shapes.AddLabel(*arguments)

    def AddLine(self, BeginX=None, BeginY=None, EndX=None, EndY=None):
        arguments = com_arguments([BeginX, BeginY, EndX, EndY])
        return self.shapes.AddLine(*arguments)

    def AddMediaObject(self, FileName=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([FileName, Left, Top, Width, Height])
        return self.shapes.AddMediaObject(*arguments)

    def AddMediaObject2(self, FileName=None, LinkToFile=None, SaveWithDocument=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height])
        return self.shapes.AddMediaObject2(*arguments)

    def AddMediaObjectFromEmbedTag(self, EmbedTag=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([EmbedTag, Left, Top, Width, Height])
        return self.shapes.AddMediaObjectFromEmbedTag(*arguments)

    def AddOLEObject(self, Left=None, Top=None, Width=None, Height=None, ClassName=None, FileName=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Link=None):
        arguments = com_arguments([Left, Top, Width, Height, ClassName, FileName, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Link])
        return self.shapes.AddOLEObject(*arguments)

    def AddPicture(self, FileName=None, LinkToFile=None, SaveWithDocument=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height])
        return self.shapes.AddPicture(*arguments)

    def AddPlaceholder(self, Type=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([Type, Left, Top, Width, Height])
        return self.shapes.AddPlaceholder(*arguments)

    def AddPolyline(self, SafeArrayOfPoints=None):
        arguments = com_arguments([SafeArrayOfPoints])
        return self.shapes.AddPolyline(*arguments)

    def AddShape(self, Type=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([Type, Left, Top, Width, Height])
        return self.shapes.AddShape(*arguments)

    def AddSmartArt(self, Layout=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([Layout, Left, Top, Width, Height])
        return self.shapes.AddSmartArt(*arguments)

    def AddTable(self, NumRows=None, NumColumns=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([NumRows, NumColumns, Left, Top, Width, Height])
        return self.shapes.AddTable(*arguments)

    def AddTextbox(self, Orientation=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([Orientation, Left, Top, Width, Height])
        return self.shapes.AddTextbox(*arguments)

    def AddTextEffect(self, PresetTextEffect=None, Text=None, FontName=None, FontSize=None, FontBold=None, FontItalic=None, Left=None, Top=None):
        arguments = com_arguments([PresetTextEffect, Text, FontName, FontSize, FontBold, FontItalic, Left, Top])
        return self.shapes.AddTextEffect(*arguments)

    def AddTitle(self):
        return self.shapes.AddTitle()

    def BuildFreeform(self, EditingType=None, X1=None, Y1=None):
        arguments = com_arguments([EditingType, X1, Y1])
        return self.shapes.BuildFreeform(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.shapes.Item(*arguments)

    def Paste(self):
        return self.shapes.Paste()

    def PasteSpecial(self, DataType=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Link=None):
        arguments = com_arguments([DataType, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Link])
        return self.shapes.PasteSpecial(*arguments)

    def Range(self, Index=None):
        arguments = com_arguments([Index])
        return self.shapes.Range(*arguments)

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

    # Lower case alias for Background
    @property
    def background(self):
        return self.Background

    @property
    def BackgroundStyle(self):
        return self.slide.BackgroundStyle

    # Lower case alias for BackgroundStyle
    @property
    def backgroundstyle(self):
        return self.BackgroundStyle

    @BackgroundStyle.setter
    def BackgroundStyle(self, value):
        self.slide.BackgroundStyle = value

    # Lower case alias for BackgroundStyle setter
    @backgroundstyle.setter
    def backgroundstyle(self, value):
        self.BackgroundStyle = value

    @property
    def ColorScheme(self):
        return ColorScheme(self.slide.ColorScheme)

    # Lower case alias for ColorScheme
    @property
    def colorscheme(self):
        return self.ColorScheme

    @ColorScheme.setter
    def ColorScheme(self, value):
        self.slide.ColorScheme = value

    # Lower case alias for ColorScheme setter
    @colorscheme.setter
    def colorscheme(self, value):
        self.ColorScheme = value

    @property
    def Comments(self):
        return Comments(self.slide.Comments)

    # Lower case alias for Comments
    @property
    def comments(self):
        return self.Comments

    @property
    def CustomerData(self):
        return CustomerData(self.slide.CustomerData)

    # Lower case alias for CustomerData
    @property
    def customerdata(self):
        return self.CustomerData

    @property
    def CustomLayout(self):
        return CustomLayout(self.slide.CustomLayout)

    # Lower case alias for CustomLayout
    @property
    def customlayout(self):
        return self.CustomLayout

    @property
    def Design(self):
        return Design(self.slide.Design)

    # Lower case alias for Design
    @property
    def design(self):
        return self.Design

    @property
    def DisplayMasterShapes(self):
        return self.slide.DisplayMasterShapes

    # Lower case alias for DisplayMasterShapes
    @property
    def displaymastershapes(self):
        return self.DisplayMasterShapes

    @DisplayMasterShapes.setter
    def DisplayMasterShapes(self, value):
        self.slide.DisplayMasterShapes = value

    # Lower case alias for DisplayMasterShapes setter
    @displaymastershapes.setter
    def displaymastershapes(self, value):
        self.DisplayMasterShapes = value

    @property
    def FollowMasterBackground(self):
        return self.slide.FollowMasterBackground

    # Lower case alias for FollowMasterBackground
    @property
    def followmasterbackground(self):
        return self.FollowMasterBackground

    @FollowMasterBackground.setter
    def FollowMasterBackground(self, value):
        self.slide.FollowMasterBackground = value

    # Lower case alias for FollowMasterBackground setter
    @followmasterbackground.setter
    def followmasterbackground(self, value):
        self.FollowMasterBackground = value

    @property
    def HasNotesPage(self):
        return self.slide.HasNotesPage

    # Lower case alias for HasNotesPage
    @property
    def hasnotespage(self):
        return self.HasNotesPage

    @property
    def HeadersFooters(self):
        return HeadersFooters(self.slide.HeadersFooters)

    # Lower case alias for HeadersFooters
    @property
    def headersfooters(self):
        return self.HeadersFooters

    @property
    def Hyperlinks(self):
        return Hyperlinks(self.slide.Hyperlinks)

    # Lower case alias for Hyperlinks
    @property
    def hyperlinks(self):
        return self.Hyperlinks

    @property
    def Layout(self):
        return PpSlideLayout(self.slide.Layout)

    # Lower case alias for Layout
    @property
    def layout(self):
        return self.Layout

    @Layout.setter
    def Layout(self, value):
        self.slide.Layout = value

    # Lower case alias for Layout setter
    @layout.setter
    def layout(self, value):
        self.Layout = value

    @property
    def Master(self):
        return Master(self.slide.Master)

    # Lower case alias for Master
    @property
    def master(self):
        return self.Master

    @property
    def Name(self):
        return self.slide.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def NotesPage(self):
        return SlideRange(self.slide.NotesPage)

    # Lower case alias for NotesPage
    @property
    def notespage(self):
        return self.NotesPage

    @property
    def Parent(self):
        return self.slide.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PrintSteps(self):
        return self.slide.PrintSteps

    # Lower case alias for PrintSteps
    @property
    def printsteps(self):
        return self.PrintSteps

    @property
    def sectionIndex(self):
        return Slide(self.slide.sectionIndex)

    # Lower case alias for sectionIndex
    @property
    def sectionindex(self):
        return self.sectionIndex

    @property
    def Shapes(self):
        return Shapes(self.slide.Shapes)

    # Lower case alias for Shapes
    @property
    def shapes(self):
        return self.Shapes

    @property
    def SlideID(self):
        return self.slide.SlideID

    # Lower case alias for SlideID
    @property
    def slideid(self):
        return self.SlideID

    @property
    def SlideIndex(self):
        return Slides(self.slide.SlideIndex)

    # Lower case alias for SlideIndex
    @property
    def slideindex(self):
        return self.SlideIndex

    @property
    def SlideNumber(self):
        return self.slide.SlideNumber

    # Lower case alias for SlideNumber
    @property
    def slidenumber(self):
        return self.SlideNumber

    @property
    def SlideShowTransition(self):
        return SlideShowTransition(self.slide.SlideShowTransition)

    # Lower case alias for SlideShowTransition
    @property
    def slideshowtransition(self):
        return self.SlideShowTransition

    @property
    def Tags(self):
        return Tags(self.slide.Tags)

    # Lower case alias for Tags
    @property
    def tags(self):
        return self.Tags

    @property
    def ThemeColorScheme(self):
        return self.slide.ThemeColorScheme

    # Lower case alias for ThemeColorScheme
    @property
    def themecolorscheme(self):
        return self.ThemeColorScheme

    @property
    def TimeLine(self):
        return TimeLine(self.slide.TimeLine)

    # Lower case alias for TimeLine
    @property
    def timeline(self):
        return self.TimeLine

    def ApplyTemplate(self, FileName=None):
        arguments = com_arguments([FileName])
        self.slide.ApplyTemplate(*arguments)

    def ApplyTheme(self, themeName=None):
        arguments = com_arguments([themeName])
        self.slide.ApplyTheme(*arguments)

    def ApplyThemeColorScheme(self, themeColorSchemeName=None):
        arguments = com_arguments([themeColorSchemeName])
        self.slide.ApplyThemeColorScheme(*arguments)

    def Copy(self):
        self.slide.Copy()

    def Cut(self):
        self.slide.Cut()

    def Delete(self):
        self.slide.Delete()

    def Duplicate(self):
        return self.slide.Duplicate()

    def Export(self, FileName=None, FilterName=None, ScaleWidth=None, ScaleHeight=None):
        arguments = com_arguments([FileName, FilterName, ScaleWidth, ScaleHeight])
        self.slide.Export(*arguments)

    def MoveTo(self, toPos=None):
        arguments = com_arguments([toPos])
        self.slide.MoveTo(*arguments)

    def MoveToSectionStart(self, toSection=None):
        arguments = com_arguments([toSection])
        self.slide.MoveToSectionStart(*arguments)

    def PublishSlides(self, SlideLibraryUrl=None, Overwrite=None, UseSlideOrder=None):
        arguments = com_arguments([SlideLibraryUrl, Overwrite, UseSlideOrder])
        return self.slide.PublishSlides(*arguments)

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

    # Lower case alias for Background
    @property
    def background(self):
        return self.Background

    @property
    def BackgroundStyle(self):
        return self.sliderange.BackgroundStyle

    # Lower case alias for BackgroundStyle
    @property
    def backgroundstyle(self):
        return self.BackgroundStyle

    @BackgroundStyle.setter
    def BackgroundStyle(self, value):
        self.sliderange.BackgroundStyle = value

    # Lower case alias for BackgroundStyle setter
    @backgroundstyle.setter
    def backgroundstyle(self, value):
        self.BackgroundStyle = value

    @property
    def ColorScheme(self):
        return ColorScheme(self.sliderange.ColorScheme)

    # Lower case alias for ColorScheme
    @property
    def colorscheme(self):
        return self.ColorScheme

    @ColorScheme.setter
    def ColorScheme(self, value):
        self.sliderange.ColorScheme = value

    # Lower case alias for ColorScheme setter
    @colorscheme.setter
    def colorscheme(self, value):
        self.ColorScheme = value

    @property
    def Comments(self):
        return Comments(self.sliderange.Comments)

    # Lower case alias for Comments
    @property
    def comments(self):
        return self.Comments

    @property
    def Count(self):
        return self.sliderange.Count

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def CustomerData(self):
        return CustomerData(self.sliderange.CustomerData)

    # Lower case alias for CustomerData
    @property
    def customerdata(self):
        return self.CustomerData

    @property
    def CustomLayout(self):
        return CustomLayout(self.sliderange.CustomLayout)

    # Lower case alias for CustomLayout
    @property
    def customlayout(self):
        return self.CustomLayout

    @property
    def Design(self):
        return Design(self.sliderange.Design)

    # Lower case alias for Design
    @property
    def design(self):
        return self.Design

    @property
    def DisplayMasterShapes(self):
        return self.sliderange.DisplayMasterShapes

    # Lower case alias for DisplayMasterShapes
    @property
    def displaymastershapes(self):
        return self.DisplayMasterShapes

    @DisplayMasterShapes.setter
    def DisplayMasterShapes(self, value):
        self.sliderange.DisplayMasterShapes = value

    # Lower case alias for DisplayMasterShapes setter
    @displaymastershapes.setter
    def displaymastershapes(self, value):
        self.DisplayMasterShapes = value

    @property
    def FollowMasterBackground(self):
        return self.sliderange.FollowMasterBackground

    # Lower case alias for FollowMasterBackground
    @property
    def followmasterbackground(self):
        return self.FollowMasterBackground

    @FollowMasterBackground.setter
    def FollowMasterBackground(self, value):
        self.sliderange.FollowMasterBackground = value

    # Lower case alias for FollowMasterBackground setter
    @followmasterbackground.setter
    def followmasterbackground(self, value):
        self.FollowMasterBackground = value

    @property
    def HasNotesPage(self):
        return self.sliderange.HasNotesPage

    # Lower case alias for HasNotesPage
    @property
    def hasnotespage(self):
        return self.HasNotesPage

    @property
    def HeadersFooters(self):
        return HeadersFooters(self.sliderange.HeadersFooters)

    # Lower case alias for HeadersFooters
    @property
    def headersfooters(self):
        return self.HeadersFooters

    @property
    def Hyperlinks(self):
        return Hyperlinks(self.sliderange.Hyperlinks)

    # Lower case alias for Hyperlinks
    @property
    def hyperlinks(self):
        return self.Hyperlinks

    @property
    def Layout(self):
        return PpSlideLayout(self.sliderange.Layout)

    # Lower case alias for Layout
    @property
    def layout(self):
        return self.Layout

    @Layout.setter
    def Layout(self, value):
        self.sliderange.Layout = value

    # Lower case alias for Layout setter
    @layout.setter
    def layout(self, value):
        self.Layout = value

    @property
    def Master(self):
        return Master(self.sliderange.Master)

    # Lower case alias for Master
    @property
    def master(self):
        return self.Master

    @property
    def Name(self):
        return self.sliderange.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.sliderange.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def NotesPage(self):
        return SlideRange(self.sliderange.NotesPage)

    # Lower case alias for NotesPage
    @property
    def notespage(self):
        return self.NotesPage

    @property
    def Parent(self):
        return self.sliderange.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PrintSteps(self):
        return self.sliderange.PrintSteps

    # Lower case alias for PrintSteps
    @property
    def printsteps(self):
        return self.PrintSteps

    @property
    def sectionIndex(self):
        return SlideRange(self.sliderange.sectionIndex)

    # Lower case alias for sectionIndex
    @property
    def sectionindex(self):
        return self.sectionIndex

    @property
    def Shapes(self):
        return Shapes(self.sliderange.Shapes)

    # Lower case alias for Shapes
    @property
    def shapes(self):
        return self.Shapes

    @property
    def SlideID(self):
        return self.sliderange.SlideID

    # Lower case alias for SlideID
    @property
    def slideid(self):
        return self.SlideID

    @property
    def SlideIndex(self):
        return Slides(self.sliderange.SlideIndex)

    # Lower case alias for SlideIndex
    @property
    def slideindex(self):
        return self.SlideIndex

    @property
    def SlideNumber(self):
        return self.sliderange.SlideNumber

    # Lower case alias for SlideNumber
    @property
    def slidenumber(self):
        return self.SlideNumber

    @property
    def SlideShowTransition(self):
        return SlideShowTransition(self.sliderange.SlideShowTransition)

    # Lower case alias for SlideShowTransition
    @property
    def slideshowtransition(self):
        return self.SlideShowTransition

    @property
    def Tags(self):
        return Tags(self.sliderange.Tags)

    # Lower case alias for Tags
    @property
    def tags(self):
        return self.Tags

    @property
    def ThemeColorScheme(self):
        return self.sliderange.ThemeColorScheme

    # Lower case alias for ThemeColorScheme
    @property
    def themecolorscheme(self):
        return self.ThemeColorScheme

    @property
    def TimeLine(self):
        return TimeLine(self.sliderange.TimeLine)

    # Lower case alias for TimeLine
    @property
    def timeline(self):
        return self.TimeLine

    def ApplyTemplate(self, FileName=None):
        arguments = com_arguments([FileName])
        self.sliderange.ApplyTemplate(*arguments)

    def ApplyTheme(self, themeName=None):
        arguments = com_arguments([themeName])
        self.sliderange.ApplyTheme(*arguments)

    def ApplyThemeColorScheme(self, themeColorSchemeName=None):
        arguments = com_arguments([themeColorSchemeName])
        self.sliderange.ApplyThemeColorScheme(*arguments)

    def Copy(self):
        self.sliderange.Copy()

    def Cut(self):
        self.sliderange.Cut()

    def Delete(self):
        self.sliderange.Delete()

    def Duplicate(self):
        return self.sliderange.Duplicate()

    def Export(self, FileName=None, FilterName=None, ScaleWidth=None, ScaleHeight=None):
        arguments = com_arguments([FileName, FilterName, ScaleWidth, ScaleHeight])
        self.sliderange.Export(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.sliderange.Item(*arguments)

    def MoveTo(self, toPos=None):
        arguments = com_arguments([toPos])
        self.sliderange.MoveTo(*arguments)

    def MoveToSectionStart(self, toSection=None):
        arguments = com_arguments([toSection])
        self.sliderange.MoveToSectionStart(*arguments)

    def PublishSlides(self, SlideLibraryUrl=None, Overwrite=None):
        arguments = com_arguments([SlideLibraryUrl, Overwrite])
        self.sliderange.PublishSlides(*arguments)

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.slides.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def AddSlide(self, Index=None, pCustomLayout=None):
        arguments = com_arguments([Index, pCustomLayout])
        return self.slides.AddSlide(*arguments)

    def FindBySlideID(self, SlideID=None):
        arguments = com_arguments([SlideID])
        return self.slides.FindBySlideID(*arguments)

    def InsertFromFile(self, FileName=None, Index=None, SlideStart=None, SlideEnd=None):
        arguments = com_arguments([FileName, Index, SlideStart, SlideEnd])
        return self.slides.InsertFromFile(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.slides.Item(*arguments)

    def Paste(self, Index=None):
        arguments = com_arguments([Index])
        return self.slides.Paste(*arguments)

    def Range(self, Index=None):
        arguments = com_arguments([Index])
        return self.slides.Range(*arguments)


class SlideShowSettings:

    def __init__(self, slideshowsettings=None):
        self.slideshowsettings = slideshowsettings

    @property
    def AdvanceMode(self):
        return self.slideshowsettings.AdvanceMode

    # Lower case alias for AdvanceMode
    @property
    def advancemode(self):
        return self.AdvanceMode

    @AdvanceMode.setter
    def AdvanceMode(self, value):
        self.slideshowsettings.AdvanceMode = value

    # Lower case alias for AdvanceMode setter
    @advancemode.setter
    def advancemode(self, value):
        self.AdvanceMode = value

    @property
    def Application(self):
        return Application(self.slideshowsettings.Application)

    @property
    def EndingSlide(self):
        return self.slideshowsettings.EndingSlide

    # Lower case alias for EndingSlide
    @property
    def endingslide(self):
        return self.EndingSlide

    @EndingSlide.setter
    def EndingSlide(self, value):
        self.slideshowsettings.EndingSlide = value

    # Lower case alias for EndingSlide setter
    @endingslide.setter
    def endingslide(self, value):
        self.EndingSlide = value

    @property
    def LoopUntilStopped(self):
        return self.slideshowsettings.LoopUntilStopped

    # Lower case alias for LoopUntilStopped
    @property
    def loopuntilstopped(self):
        return self.LoopUntilStopped

    @LoopUntilStopped.setter
    def LoopUntilStopped(self, value):
        self.slideshowsettings.LoopUntilStopped = value

    # Lower case alias for LoopUntilStopped setter
    @loopuntilstopped.setter
    def loopuntilstopped(self, value):
        self.LoopUntilStopped = value

    @property
    def NamedSlideShows(self):
        return NamedSlideShows(self.slideshowsettings.NamedSlideShows)

    # Lower case alias for NamedSlideShows
    @property
    def namedslideshows(self):
        return self.NamedSlideShows

    @property
    def Parent(self):
        return self.slideshowsettings.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PointerColor(self):
        return ColorFormat(self.slideshowsettings.PointerColor)

    # Lower case alias for PointerColor
    @property
    def pointercolor(self):
        return self.PointerColor

    @property
    def RangeType(self):
        return self.slideshowsettings.RangeType

    # Lower case alias for RangeType
    @property
    def rangetype(self):
        return self.RangeType

    @RangeType.setter
    def RangeType(self, value):
        self.slideshowsettings.RangeType = value

    # Lower case alias for RangeType setter
    @rangetype.setter
    def rangetype(self, value):
        self.RangeType = value

    @property
    def ShowMediaControls(self):
        return self.slideshowsettings.ShowMediaControls

    # Lower case alias for ShowMediaControls
    @property
    def showmediacontrols(self):
        return self.ShowMediaControls

    @ShowMediaControls.setter
    def ShowMediaControls(self, value):
        self.slideshowsettings.ShowMediaControls = value

    # Lower case alias for ShowMediaControls setter
    @showmediacontrols.setter
    def showmediacontrols(self, value):
        self.ShowMediaControls = value

    @property
    def ShowPresenterView(self):
        return SlideShowSettings(self.slideshowsettings.ShowPresenterView)

    # Lower case alias for ShowPresenterView
    @property
    def showpresenterview(self):
        return self.ShowPresenterView

    @ShowPresenterView.setter
    def ShowPresenterView(self, value):
        self.slideshowsettings.ShowPresenterView = value

    # Lower case alias for ShowPresenterView setter
    @showpresenterview.setter
    def showpresenterview(self, value):
        self.ShowPresenterView = value

    @property
    def ShowScrollbar(self):
        return self.slideshowsettings.ShowScrollbar

    # Lower case alias for ShowScrollbar
    @property
    def showscrollbar(self):
        return self.ShowScrollbar

    @ShowScrollbar.setter
    def ShowScrollbar(self, value):
        self.slideshowsettings.ShowScrollbar = value

    # Lower case alias for ShowScrollbar setter
    @showscrollbar.setter
    def showscrollbar(self, value):
        self.ShowScrollbar = value

    @property
    def ShowType(self):
        return self.slideshowsettings.ShowType

    # Lower case alias for ShowType
    @property
    def showtype(self):
        return self.ShowType

    @ShowType.setter
    def ShowType(self, value):
        self.slideshowsettings.ShowType = value

    # Lower case alias for ShowType setter
    @showtype.setter
    def showtype(self, value):
        self.ShowType = value

    @property
    def ShowWithAnimation(self):
        return self.slideshowsettings.ShowWithAnimation

    # Lower case alias for ShowWithAnimation
    @property
    def showwithanimation(self):
        return self.ShowWithAnimation

    @ShowWithAnimation.setter
    def ShowWithAnimation(self, value):
        self.slideshowsettings.ShowWithAnimation = value

    # Lower case alias for ShowWithAnimation setter
    @showwithanimation.setter
    def showwithanimation(self, value):
        self.ShowWithAnimation = value

    @property
    def ShowWithNarration(self):
        return self.slideshowsettings.ShowWithNarration

    # Lower case alias for ShowWithNarration
    @property
    def showwithnarration(self):
        return self.ShowWithNarration

    @ShowWithNarration.setter
    def ShowWithNarration(self, value):
        self.slideshowsettings.ShowWithNarration = value

    # Lower case alias for ShowWithNarration setter
    @showwithnarration.setter
    def showwithnarration(self, value):
        self.ShowWithNarration = value

    @property
    def SlideShowName(self):
        return self.slideshowsettings.SlideShowName

    # Lower case alias for SlideShowName
    @property
    def slideshowname(self):
        return self.SlideShowName

    @SlideShowName.setter
    def SlideShowName(self, value):
        self.slideshowsettings.SlideShowName = value

    # Lower case alias for SlideShowName setter
    @slideshowname.setter
    def slideshowname(self, value):
        self.SlideShowName = value

    @property
    def StartingSlide(self):
        return self.slideshowsettings.StartingSlide

    # Lower case alias for StartingSlide
    @property
    def startingslide(self):
        return self.StartingSlide

    @StartingSlide.setter
    def StartingSlide(self, value):
        self.slideshowsettings.StartingSlide = value

    # Lower case alias for StartingSlide setter
    @startingslide.setter
    def startingslide(self, value):
        self.StartingSlide = value

    def Run(self):
        return self.slideshowsettings.Run()


class SlideShowTransition:

    def __init__(self, slideshowtransition=None):
        self.slideshowtransition = slideshowtransition

    @property
    def AdvanceOnClick(self):
        return self.slideshowtransition.AdvanceOnClick

    # Lower case alias for AdvanceOnClick
    @property
    def advanceonclick(self):
        return self.AdvanceOnClick

    @AdvanceOnClick.setter
    def AdvanceOnClick(self, value):
        self.slideshowtransition.AdvanceOnClick = value

    # Lower case alias for AdvanceOnClick setter
    @advanceonclick.setter
    def advanceonclick(self, value):
        self.AdvanceOnClick = value

    @property
    def AdvanceOnTime(self):
        return self.slideshowtransition.AdvanceOnTime

    # Lower case alias for AdvanceOnTime
    @property
    def advanceontime(self):
        return self.AdvanceOnTime

    @AdvanceOnTime.setter
    def AdvanceOnTime(self, value):
        self.slideshowtransition.AdvanceOnTime = value

    # Lower case alias for AdvanceOnTime setter
    @advanceontime.setter
    def advanceontime(self, value):
        self.AdvanceOnTime = value

    @property
    def AdvanceTime(self):
        return self.slideshowtransition.AdvanceTime

    # Lower case alias for AdvanceTime
    @property
    def advancetime(self):
        return self.AdvanceTime

    @AdvanceTime.setter
    def AdvanceTime(self, value):
        self.slideshowtransition.AdvanceTime = value

    # Lower case alias for AdvanceTime setter
    @advancetime.setter
    def advancetime(self, value):
        self.AdvanceTime = value

    @property
    def Application(self):
        return Application(self.slideshowtransition.Application)

    @property
    def Duration(self):
        return self.slideshowtransition.Duration

    # Lower case alias for Duration
    @property
    def duration(self):
        return self.Duration

    @Duration.setter
    def Duration(self, value):
        self.slideshowtransition.Duration = value

    # Lower case alias for Duration setter
    @duration.setter
    def duration(self, value):
        self.Duration = value

    @property
    def EntryEffect(self):
        return self.slideshowtransition.EntryEffect

    # Lower case alias for EntryEffect
    @property
    def entryeffect(self):
        return self.EntryEffect

    @EntryEffect.setter
    def EntryEffect(self, value):
        self.slideshowtransition.EntryEffect = value

    # Lower case alias for EntryEffect setter
    @entryeffect.setter
    def entryeffect(self, value):
        self.EntryEffect = value

    @property
    def Hidden(self):
        return self.slideshowtransition.Hidden

    # Lower case alias for Hidden
    @property
    def hidden(self):
        return self.Hidden

    @Hidden.setter
    def Hidden(self, value):
        self.slideshowtransition.Hidden = value

    # Lower case alias for Hidden setter
    @hidden.setter
    def hidden(self, value):
        self.Hidden = value

    @property
    def LoopSoundUntilNext(self):
        return self.slideshowtransition.LoopSoundUntilNext

    # Lower case alias for LoopSoundUntilNext
    @property
    def loopsounduntilnext(self):
        return self.LoopSoundUntilNext

    @LoopSoundUntilNext.setter
    def LoopSoundUntilNext(self, value):
        self.slideshowtransition.LoopSoundUntilNext = value

    # Lower case alias for LoopSoundUntilNext setter
    @loopsounduntilnext.setter
    def loopsounduntilnext(self, value):
        self.LoopSoundUntilNext = value

    @property
    def Parent(self):
        return self.slideshowtransition.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SoundEffect(self):
        return SoundEffect(self.slideshowtransition.SoundEffect)

    # Lower case alias for SoundEffect
    @property
    def soundeffect(self):
        return self.SoundEffect

    @property
    def Speed(self):
        return self.slideshowtransition.Speed

    # Lower case alias for Speed
    @property
    def speed(self):
        return self.Speed

    @Speed.setter
    def Speed(self, value):
        self.slideshowtransition.Speed = value

    # Lower case alias for Speed setter
    @speed.setter
    def speed(self, value):
        self.Speed = value


class SlideShowView:

    def __init__(self, slideshowview=None):
        self.slideshowview = slideshowview

    @property
    def AcceleratorsEnabled(self):
        return self.slideshowview.AcceleratorsEnabled

    # Lower case alias for AcceleratorsEnabled
    @property
    def acceleratorsenabled(self):
        return self.AcceleratorsEnabled

    @AcceleratorsEnabled.setter
    def AcceleratorsEnabled(self, value):
        self.slideshowview.AcceleratorsEnabled = value

    # Lower case alias for AcceleratorsEnabled setter
    @acceleratorsenabled.setter
    def acceleratorsenabled(self, value):
        self.AcceleratorsEnabled = value

    @property
    def AdvanceMode(self):
        return self.slideshowview.AdvanceMode

    # Lower case alias for AdvanceMode
    @property
    def advancemode(self):
        return self.AdvanceMode

    @property
    def Application(self):
        return Application(self.slideshowview.Application)

    @property
    def CurrentShowPosition(self):
        return self.slideshowview.CurrentShowPosition

    # Lower case alias for CurrentShowPosition
    @property
    def currentshowposition(self):
        return self.CurrentShowPosition

    @property
    def IsNamedShow(self):
        return self.slideshowview.IsNamedShow

    # Lower case alias for IsNamedShow
    @property
    def isnamedshow(self):
        return self.IsNamedShow

    @property
    def LastSlideViewed(self):
        return Slide(self.slideshowview.LastSlideViewed)

    # Lower case alias for LastSlideViewed
    @property
    def lastslideviewed(self):
        return self.LastSlideViewed

    @property
    def MediaControlsHeight(self):
        return self.slideshowview.MediaControlsHeight

    # Lower case alias for MediaControlsHeight
    @property
    def mediacontrolsheight(self):
        return self.MediaControlsHeight

    @property
    def MediaControlsLeft(self):
        return Slide(self.slideshowview.MediaControlsLeft)

    # Lower case alias for MediaControlsLeft
    @property
    def mediacontrolsleft(self):
        return self.MediaControlsLeft

    @property
    def MediaControlsTop(self):
        return Slide(self.slideshowview.MediaControlsTop)

    # Lower case alias for MediaControlsTop
    @property
    def mediacontrolstop(self):
        return self.MediaControlsTop

    @property
    def MediaControlsVisible(self):
        return self.slideshowview.MediaControlsVisible

    # Lower case alias for MediaControlsVisible
    @property
    def mediacontrolsvisible(self):
        return self.MediaControlsVisible

    @property
    def MediaControlsWidth(self):
        return self.slideshowview.MediaControlsWidth

    # Lower case alias for MediaControlsWidth
    @property
    def mediacontrolswidth(self):
        return self.MediaControlsWidth

    @property
    def Parent(self):
        return self.slideshowview.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PointerColor(self):
        return ColorFormat(self.slideshowview.PointerColor)

    # Lower case alias for PointerColor
    @property
    def pointercolor(self):
        return self.PointerColor

    @property
    def PointerType(self):
        return self.slideshowview.PointerType

    # Lower case alias for PointerType
    @property
    def pointertype(self):
        return self.PointerType

    @PointerType.setter
    def PointerType(self, value):
        self.slideshowview.PointerType = value

    # Lower case alias for PointerType setter
    @pointertype.setter
    def pointertype(self, value):
        self.PointerType = value

    @property
    def PresentationElapsedTime(self):
        return self.slideshowview.PresentationElapsedTime

    # Lower case alias for PresentationElapsedTime
    @property
    def presentationelapsedtime(self):
        return self.PresentationElapsedTime

    @property
    def Slide(self):
        return Slide(self.slideshowview.Slide)

    # Lower case alias for Slide
    @property
    def slide(self):
        return self.Slide

    @property
    def SlideElapsedTime(self):
        return self.slideshowview.SlideElapsedTime

    # Lower case alias for SlideElapsedTime
    @property
    def slideelapsedtime(self):
        return self.SlideElapsedTime

    @SlideElapsedTime.setter
    def SlideElapsedTime(self, value):
        self.slideshowview.SlideElapsedTime = value

    # Lower case alias for SlideElapsedTime setter
    @slideelapsedtime.setter
    def slideelapsedtime(self, value):
        self.SlideElapsedTime = value

    @property
    def SlideShowName(self):
        return self.slideshowview.SlideShowName

    # Lower case alias for SlideShowName
    @property
    def slideshowname(self):
        return self.SlideShowName

    @property
    def State(self):
        return self.slideshowview.State

    # Lower case alias for State
    @property
    def state(self):
        return self.State

    @State.setter
    def State(self, value):
        self.slideshowview.State = value

    # Lower case alias for State setter
    @state.setter
    def state(self, value):
        self.State = value

    @property
    def Zoom(self):
        return self.slideshowview.Zoom

    # Lower case alias for Zoom
    @property
    def zoom(self):
        return self.Zoom

    def DrawLine(self, BeginX=None, BeginY=None, EndX=None, EndY=None):
        arguments = com_arguments([BeginX, BeginY, EndX, EndY])
        self.slideshowview.DrawLine(*arguments)

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
        arguments = com_arguments([Index])
        self.slideshowview.GotoClick(*arguments)

    def GotoNamedShow(self, SlideShowName=None):
        arguments = com_arguments([SlideShowName])
        self.slideshowview.GotoNamedShow(*arguments)

    def GotoSlide(self, Index=None, ResetSlide=None):
        arguments = com_arguments([Index, ResetSlide])
        self.slideshowview.GotoSlide(*arguments)

    def Last(self):
        self.slideshowview.Last()

    def Next(self):
        self.slideshowview.Next()

    def Player(self, ShapeId=None):
        arguments = com_arguments([ShapeId])
        return self.slideshowview.Player(*arguments)

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

    # Lower case alias for Active
    @property
    def active(self):
        return self.Active

    @property
    def Application(self):
        return Application(self.slideshowwindow.Application)

    @property
    def Height(self):
        return self.slideshowwindow.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @Height.setter
    def Height(self, value):
        self.slideshowwindow.Height = value

    # Lower case alias for Height setter
    @height.setter
    def height(self, value):
        self.Height = value

    @property
    def IsFullScreen(self):
        return self.slideshowwindow.IsFullScreen

    # Lower case alias for IsFullScreen
    @property
    def isfullscreen(self):
        return self.IsFullScreen

    @property
    def Left(self):
        return self.slideshowwindow.Left

    # Lower case alias for Left
    @property
    def left(self):
        return self.Left

    @Left.setter
    def Left(self, value):
        self.slideshowwindow.Left = value

    # Lower case alias for Left setter
    @left.setter
    def left(self, value):
        self.Left = value

    @property
    def Parent(self):
        return self.slideshowwindow.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Presentation(self):
        return Presentation(self.slideshowwindow.Presentation)

    # Lower case alias for Presentation
    @property
    def presentation(self):
        return self.Presentation

    @property
    def Top(self):
        return self.slideshowwindow.Top

    # Lower case alias for Top
    @property
    def top(self):
        return self.Top

    @Top.setter
    def Top(self, value):
        self.slideshowwindow.Top = value

    # Lower case alias for Top setter
    @top.setter
    def top(self, value):
        self.Top = value

    @property
    def View(self):
        return SlideShowView(self.slideshowwindow.View)

    # Lower case alias for View
    @property
    def view(self):
        return self.View

    @property
    def Width(self):
        return self.slideshowwindow.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

    @Width.setter
    def Width(self, value):
        self.slideshowwindow.Width = value

    # Lower case alias for Width setter
    @width.setter
    def width(self, value):
        self.Width = value

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.slideshowwindows.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.slideshowwindows.Item(*arguments)


class SoundEffect:

    def __init__(self, soundeffect=None):
        self.soundeffect = soundeffect

    @property
    def Application(self):
        return Application(self.soundeffect.Application)

    @property
    def Name(self):
        return self.soundeffect.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.soundeffect.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.soundeffect.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Type(self):
        return self.soundeffect.Type

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @Type.setter
    def Type(self, value):
        self.soundeffect.Type = value

    # Lower case alias for Type setter
    @type.setter
    def type(self, value):
        self.Type = value

    def ImportFromFile(self, FullName=None):
        arguments = com_arguments([FullName])
        self.soundeffect.ImportFromFile(*arguments)

    def Play(self):
        self.soundeffect.Play()


class Table:

    def __init__(self, table=None):
        self.table = table

    @property
    def AlternativeText(self):
        return self.table.AlternativeText

    # Lower case alias for AlternativeText
    @property
    def alternativetext(self):
        return self.AlternativeText

    @AlternativeText.setter
    def AlternativeText(self, value):
        self.table.AlternativeText = value

    # Lower case alias for AlternativeText setter
    @alternativetext.setter
    def alternativetext(self, value):
        self.AlternativeText = value

    @property
    def Application(self):
        return Application(self.table.Application)

    @property
    def Background(self):
        return TableBackground(self.table.Background)

    # Lower case alias for Background
    @property
    def background(self):
        return self.Background

    @property
    def Columns(self):
        return Columns(self.table.Columns)

    # Lower case alias for Columns
    @property
    def columns(self):
        return self.Columns

    @property
    def FirstCol(self):
        return self.table.FirstCol

    # Lower case alias for FirstCol
    @property
    def firstcol(self):
        return self.FirstCol

    @FirstCol.setter
    def FirstCol(self, value):
        self.table.FirstCol = value

    # Lower case alias for FirstCol setter
    @firstcol.setter
    def firstcol(self, value):
        self.FirstCol = value

    @property
    def FirstRow(self):
        return self.table.FirstRow

    # Lower case alias for FirstRow
    @property
    def firstrow(self):
        return self.FirstRow

    @FirstRow.setter
    def FirstRow(self, value):
        self.table.FirstRow = value

    # Lower case alias for FirstRow setter
    @firstrow.setter
    def firstrow(self, value):
        self.FirstRow = value

    @property
    def HorizBanding(self):
        return self.table.HorizBanding

    # Lower case alias for HorizBanding
    @property
    def horizbanding(self):
        return self.HorizBanding

    @HorizBanding.setter
    def HorizBanding(self, value):
        self.table.HorizBanding = value

    # Lower case alias for HorizBanding setter
    @horizbanding.setter
    def horizbanding(self, value):
        self.HorizBanding = value

    @property
    def LastCol(self):
        return self.table.LastCol

    # Lower case alias for LastCol
    @property
    def lastcol(self):
        return self.LastCol

    @LastCol.setter
    def LastCol(self, value):
        self.table.LastCol = value

    # Lower case alias for LastCol setter
    @lastcol.setter
    def lastcol(self, value):
        self.LastCol = value

    @property
    def LastRow(self):
        return self.table.LastRow

    # Lower case alias for LastRow
    @property
    def lastrow(self):
        return self.LastRow

    @LastRow.setter
    def LastRow(self, value):
        self.table.LastRow = value

    # Lower case alias for LastRow setter
    @lastrow.setter
    def lastrow(self, value):
        self.LastRow = value

    @property
    def Parent(self):
        return self.table.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Rows(self):
        return Rows(self.table.Rows)

    # Lower case alias for Rows
    @property
    def rows(self):
        return self.Rows

    @property
    def Style(self):
        return TableStyle(self.table.Style)

    # Lower case alias for Style
    @property
    def style(self):
        return self.Style

    @property
    def TableDirection(self):
        return self.table.TableDirection

    # Lower case alias for TableDirection
    @property
    def tabledirection(self):
        return self.TableDirection

    @TableDirection.setter
    def TableDirection(self, value):
        self.table.TableDirection = value

    # Lower case alias for TableDirection setter
    @tabledirection.setter
    def tabledirection(self, value):
        self.TableDirection = value

    @property
    def Title(self):
        return Table(self.table.Title)

    # Lower case alias for Title
    @property
    def title(self):
        return self.Title

    @Title.setter
    def Title(self, value):
        self.table.Title = value

    # Lower case alias for Title setter
    @title.setter
    def title(self, value):
        self.Title = value

    @property
    def VertBanding(self):
        return self.table.VertBanding

    # Lower case alias for VertBanding
    @property
    def vertbanding(self):
        return self.VertBanding

    @VertBanding.setter
    def VertBanding(self, value):
        self.table.VertBanding = value

    # Lower case alias for VertBanding setter
    @vertbanding.setter
    def vertbanding(self, value):
        self.VertBanding = value

    def ApplyStyle(self, StyleID=None, SaveFormatting=None):
        arguments = com_arguments([StyleID, SaveFormatting])
        self.table.ApplyStyle(*arguments)

    def Cell(self, Row=None, Column=None):
        arguments = com_arguments([Row, Column])
        return self.table.Cell(*arguments)

    def ScaleProportionally(self, scale=None):
        arguments = com_arguments([scale])
        self.table.ScaleProportionally(*arguments)


class TableBackground:

    def __init__(self, tablebackground=None):
        self.tablebackground = tablebackground

    @property
    def Fill(self):
        return FillFormat(self.tablebackground.Fill)

    # Lower case alias for Fill
    @property
    def fill(self):
        return self.Fill

    @property
    def Picture(self):
        return PictureFormat(self.tablebackground.Picture)

    # Lower case alias for Picture
    @property
    def picture(self):
        return self.Picture

    @property
    def Reflection(self):
        return self.tablebackground.Reflection

    # Lower case alias for Reflection
    @property
    def reflection(self):
        return self.Reflection

    @property
    def Shadow(self):
        return ShadowFormat(self.tablebackground.Shadow)

    # Lower case alias for Shadow
    @property
    def shadow(self):
        return self.Shadow


class TableStyle:

    def __init__(self, tablestyle=None):
        self.tablestyle = tablestyle

    @property
    def Id(self):
        return self.tablestyle.Id

    # Lower case alias for Id
    @property
    def id(self):
        return self.Id

    @property
    def Name(self):
        return self.tablestyle.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name


class TabStop:

    def __init__(self, tabstop=None):
        self.tabstop = tabstop

    @property
    def Application(self):
        return Application(self.tabstop.Application)

    @property
    def Parent(self):
        return self.tabstop.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return self.tabstop.Position

    # Lower case alias for Position
    @property
    def position(self):
        return self.Position

    @Position.setter
    def Position(self, value):
        self.tabstop.Position = value

    # Lower case alias for Position setter
    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Type(self):
        return self.tabstop.Type

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @Type.setter
    def Type(self, value):
        self.tabstop.Type = value

    # Lower case alias for Type setter
    @type.setter
    def type(self, value):
        self.Type = value

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def DefaultSpacing(self):
        return self.tabstops.DefaultSpacing

    # Lower case alias for DefaultSpacing
    @property
    def defaultspacing(self):
        return self.DefaultSpacing

    @DefaultSpacing.setter
    def DefaultSpacing(self, value):
        self.tabstops.DefaultSpacing = value

    # Lower case alias for DefaultSpacing setter
    @defaultspacing.setter
    def defaultspacing(self, value):
        self.DefaultSpacing = value

    @property
    def Parent(self):
        return self.tabstops.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Add(self, Type=None, Position=None):
        arguments = com_arguments([Type, Position])
        return TabStop(self.tabstops.Add(*arguments))

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.tabstops.Item(*arguments)


class Tags:

    def __init__(self, tags=None):
        self.tags = tags

    @property
    def Application(self):
        return Application(self.tags.Application)

    @property
    def Count(self):
        return self.tags.Count

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.tags.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Add(self, Name=None, Value=None):
        arguments = com_arguments([Name, Value])
        self.tags.Add(*arguments)

    def Delete(self, Name=None):
        arguments = com_arguments([Name])
        self.tags.Delete(*arguments)

    def Item(self, Name=None):
        arguments = com_arguments([Name])
        return self.tags.Item(*arguments)

    def Name(self, Index=None):
        arguments = com_arguments([Index])
        return self.tags.Name(*arguments)

    def Value(self, Index=None):
        arguments = com_arguments([Index])
        return self.tags.Value(*arguments)


class TextEffectFormat:

    def __init__(self, texteffectformat=None):
        self.texteffectformat = texteffectformat

    @property
    def Alignment(self):
        return self.texteffectformat.Alignment

    # Lower case alias for Alignment
    @property
    def alignment(self):
        return self.Alignment

    @Alignment.setter
    def Alignment(self, value):
        self.texteffectformat.Alignment = value

    # Lower case alias for Alignment setter
    @alignment.setter
    def alignment(self, value):
        self.Alignment = value

    @property
    def Application(self):
        return Application(self.texteffectformat.Application)

    @property
    def Creator(self):
        return self.texteffectformat.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def FontBold(self):
        return self.texteffectformat.FontBold

    # Lower case alias for FontBold
    @property
    def fontbold(self):
        return self.FontBold

    @FontBold.setter
    def FontBold(self, value):
        self.texteffectformat.FontBold = value

    # Lower case alias for FontBold setter
    @fontbold.setter
    def fontbold(self, value):
        self.FontBold = value

    @property
    def FontItalic(self):
        return self.texteffectformat.FontItalic

    # Lower case alias for FontItalic
    @property
    def fontitalic(self):
        return self.FontItalic

    @FontItalic.setter
    def FontItalic(self, value):
        self.texteffectformat.FontItalic = value

    # Lower case alias for FontItalic setter
    @fontitalic.setter
    def fontitalic(self, value):
        self.FontItalic = value

    @property
    def FontName(self):
        return self.texteffectformat.FontName

    # Lower case alias for FontName
    @property
    def fontname(self):
        return self.FontName

    @FontName.setter
    def FontName(self, value):
        self.texteffectformat.FontName = value

    # Lower case alias for FontName setter
    @fontname.setter
    def fontname(self, value):
        self.FontName = value

    @property
    def FontSize(self):
        return self.texteffectformat.FontSize

    # Lower case alias for FontSize
    @property
    def fontsize(self):
        return self.FontSize

    @FontSize.setter
    def FontSize(self, value):
        self.texteffectformat.FontSize = value

    # Lower case alias for FontSize setter
    @fontsize.setter
    def fontsize(self, value):
        self.FontSize = value

    @property
    def KernedPairs(self):
        return self.texteffectformat.KernedPairs

    # Lower case alias for KernedPairs
    @property
    def kernedpairs(self):
        return self.KernedPairs

    @KernedPairs.setter
    def KernedPairs(self, value):
        self.texteffectformat.KernedPairs = value

    # Lower case alias for KernedPairs setter
    @kernedpairs.setter
    def kernedpairs(self, value):
        self.KernedPairs = value

    @property
    def NormalizedHeight(self):
        return self.texteffectformat.NormalizedHeight

    # Lower case alias for NormalizedHeight
    @property
    def normalizedheight(self):
        return self.NormalizedHeight

    @NormalizedHeight.setter
    def NormalizedHeight(self, value):
        self.texteffectformat.NormalizedHeight = value

    # Lower case alias for NormalizedHeight setter
    @normalizedheight.setter
    def normalizedheight(self, value):
        self.NormalizedHeight = value

    @property
    def Parent(self):
        return self.texteffectformat.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PresetShape(self):
        return self.texteffectformat.PresetShape

    # Lower case alias for PresetShape
    @property
    def presetshape(self):
        return self.PresetShape

    @PresetShape.setter
    def PresetShape(self, value):
        self.texteffectformat.PresetShape = value

    # Lower case alias for PresetShape setter
    @presetshape.setter
    def presetshape(self, value):
        self.PresetShape = value

    @property
    def PresetTextEffect(self):
        return self.texteffectformat.PresetTextEffect

    # Lower case alias for PresetTextEffect
    @property
    def presettexteffect(self):
        return self.PresetTextEffect

    @PresetTextEffect.setter
    def PresetTextEffect(self, value):
        self.texteffectformat.PresetTextEffect = value

    # Lower case alias for PresetTextEffect setter
    @presettexteffect.setter
    def presettexteffect(self, value):
        self.PresetTextEffect = value

    @property
    def RotatedChars(self):
        return self.texteffectformat.RotatedChars

    # Lower case alias for RotatedChars
    @property
    def rotatedchars(self):
        return self.RotatedChars

    @RotatedChars.setter
    def RotatedChars(self, value):
        self.texteffectformat.RotatedChars = value

    # Lower case alias for RotatedChars setter
    @rotatedchars.setter
    def rotatedchars(self, value):
        self.RotatedChars = value

    @property
    def Text(self):
        return self.texteffectformat.Text

    # Lower case alias for Text
    @property
    def text(self):
        return self.Text

    @Text.setter
    def Text(self, value):
        self.texteffectformat.Text = value

    # Lower case alias for Text setter
    @text.setter
    def text(self, value):
        self.Text = value

    @property
    def Tracking(self):
        return self.texteffectformat.Tracking

    # Lower case alias for Tracking
    @property
    def tracking(self):
        return self.Tracking

    @Tracking.setter
    def Tracking(self, value):
        self.texteffectformat.Tracking = value

    # Lower case alias for Tracking setter
    @tracking.setter
    def tracking(self, value):
        self.Tracking = value

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

    # Lower case alias for AutoSize
    @property
    def autosize(self):
        return self.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.textframe.AutoSize = value

    # Lower case alias for AutoSize setter
    @autosize.setter
    def autosize(self, value):
        self.AutoSize = value

    @property
    def Creator(self):
        return self.textframe.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def HasText(self):
        return self.textframe.HasText

    # Lower case alias for HasText
    @property
    def hastext(self):
        return self.HasText

    @property
    def HorizontalAnchor(self):
        return self.textframe.HorizontalAnchor

    # Lower case alias for HorizontalAnchor
    @property
    def horizontalanchor(self):
        return self.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.textframe.HorizontalAnchor = value

    # Lower case alias for HorizontalAnchor setter
    @horizontalanchor.setter
    def horizontalanchor(self, value):
        self.HorizontalAnchor = value

    @property
    def MarginBottom(self):
        return self.textframe.MarginBottom

    # Lower case alias for MarginBottom
    @property
    def marginbottom(self):
        return self.MarginBottom

    @MarginBottom.setter
    def MarginBottom(self, value):
        self.textframe.MarginBottom = value

    # Lower case alias for MarginBottom setter
    @marginbottom.setter
    def marginbottom(self, value):
        self.MarginBottom = value

    @property
    def MarginLeft(self):
        return self.textframe.MarginLeft

    # Lower case alias for MarginLeft
    @property
    def marginleft(self):
        return self.MarginLeft

    @MarginLeft.setter
    def MarginLeft(self, value):
        self.textframe.MarginLeft = value

    # Lower case alias for MarginLeft setter
    @marginleft.setter
    def marginleft(self, value):
        self.MarginLeft = value

    @property
    def MarginRight(self):
        return self.textframe.MarginRight

    # Lower case alias for MarginRight
    @property
    def marginright(self):
        return self.MarginRight

    @MarginRight.setter
    def MarginRight(self, value):
        self.textframe.MarginRight = value

    # Lower case alias for MarginRight setter
    @marginright.setter
    def marginright(self, value):
        self.MarginRight = value

    @property
    def MarginTop(self):
        return self.textframe.MarginTop

    # Lower case alias for MarginTop
    @property
    def margintop(self):
        return self.MarginTop

    @MarginTop.setter
    def MarginTop(self, value):
        self.textframe.MarginTop = value

    # Lower case alias for MarginTop setter
    @margintop.setter
    def margintop(self, value):
        self.MarginTop = value

    @property
    def Orientation(self):
        return self.textframe.Orientation

    # Lower case alias for Orientation
    @property
    def orientation(self):
        return self.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.textframe.Orientation = value

    # Lower case alias for Orientation setter
    @orientation.setter
    def orientation(self, value):
        self.Orientation = value

    @property
    def Parent(self):
        return self.textframe.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Ruler(self):
        return Ruler(self.textframe.Ruler)

    # Lower case alias for Ruler
    @property
    def ruler(self):
        return self.Ruler

    @property
    def TextRange(self):
        return TextRange(self.textframe.TextRange)

    # Lower case alias for TextRange
    @property
    def textrange(self):
        return self.TextRange

    @property
    def VerticalAnchor(self):
        return self.textframe.VerticalAnchor

    # Lower case alias for VerticalAnchor
    @property
    def verticalanchor(self):
        return self.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.textframe.VerticalAnchor = value

    # Lower case alias for VerticalAnchor setter
    @verticalanchor.setter
    def verticalanchor(self, value):
        self.VerticalAnchor = value

    @property
    def WordWrap(self):
        return self.textframe.WordWrap

    # Lower case alias for WordWrap
    @property
    def wordwrap(self):
        return self.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.textframe.WordWrap = value

    # Lower case alias for WordWrap setter
    @wordwrap.setter
    def wordwrap(self, value):
        self.WordWrap = value

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

    # Lower case alias for AutoSize
    @property
    def autosize(self):
        return self.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.textframe2.AutoSize = value

    # Lower case alias for AutoSize setter
    @autosize.setter
    def autosize(self, value):
        self.AutoSize = value

    @property
    def Column(self):
        return Column(self.textframe2.Column)

    # Lower case alias for Column
    @property
    def column(self):
        return self.Column

    @property
    def Creator(self):
        return self.textframe2.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def HasText(self):
        return self.textframe2.HasText

    # Lower case alias for HasText
    @property
    def hastext(self):
        return self.HasText

    @property
    def HorizontalAnchor(self):
        return self.textframe2.HorizontalAnchor

    # Lower case alias for HorizontalAnchor
    @property
    def horizontalanchor(self):
        return self.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.textframe2.HorizontalAnchor = value

    # Lower case alias for HorizontalAnchor setter
    @horizontalanchor.setter
    def horizontalanchor(self, value):
        self.HorizontalAnchor = value

    @property
    def MarginBottom(self):
        return self.textframe2.MarginBottom

    # Lower case alias for MarginBottom
    @property
    def marginbottom(self):
        return self.MarginBottom

    @MarginBottom.setter
    def MarginBottom(self, value):
        self.textframe2.MarginBottom = value

    # Lower case alias for MarginBottom setter
    @marginbottom.setter
    def marginbottom(self, value):
        self.MarginBottom = value

    @property
    def MarginLeft(self):
        return self.textframe2.MarginLeft

    # Lower case alias for MarginLeft
    @property
    def marginleft(self):
        return self.MarginLeft

    @MarginLeft.setter
    def MarginLeft(self, value):
        self.textframe2.MarginLeft = value

    # Lower case alias for MarginLeft setter
    @marginleft.setter
    def marginleft(self, value):
        self.MarginLeft = value

    @property
    def MarginRight(self):
        return self.textframe2.MarginRight

    # Lower case alias for MarginRight
    @property
    def marginright(self):
        return self.MarginRight

    @MarginRight.setter
    def MarginRight(self, value):
        self.textframe2.MarginRight = value

    # Lower case alias for MarginRight setter
    @marginright.setter
    def marginright(self, value):
        self.MarginRight = value

    @property
    def MarginTop(self):
        return self.textframe2.MarginTop

    # Lower case alias for MarginTop
    @property
    def margintop(self):
        return self.MarginTop

    @MarginTop.setter
    def MarginTop(self, value):
        self.textframe2.MarginTop = value

    # Lower case alias for MarginTop setter
    @margintop.setter
    def margintop(self, value):
        self.MarginTop = value

    @property
    def NoTextRotation(self):
        return self.textframe2.NoTextRotation

    # Lower case alias for NoTextRotation
    @property
    def notextrotation(self):
        return self.NoTextRotation

    @NoTextRotation.setter
    def NoTextRotation(self, value):
        self.textframe2.NoTextRotation = value

    # Lower case alias for NoTextRotation setter
    @notextrotation.setter
    def notextrotation(self, value):
        self.NoTextRotation = value

    @property
    def Orientation(self):
        return self.textframe2.Orientation

    # Lower case alias for Orientation
    @property
    def orientation(self):
        return self.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.textframe2.Orientation = value

    # Lower case alias for Orientation setter
    @orientation.setter
    def orientation(self, value):
        self.Orientation = value

    @property
    def Parent(self):
        return self.textframe2.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PathFormat(self):
        return self.textframe2.PathFormat

    # Lower case alias for PathFormat
    @property
    def pathformat(self):
        return self.PathFormat

    @PathFormat.setter
    def PathFormat(self, value):
        self.textframe2.PathFormat = value

    # Lower case alias for PathFormat setter
    @pathformat.setter
    def pathformat(self, value):
        self.PathFormat = value

    @property
    def Ruler(self):
        return self.textframe2.Ruler

    # Lower case alias for Ruler
    @property
    def ruler(self):
        return self.Ruler

    @property
    def TextRange(self):
        return self.textframe2.TextRange

    # Lower case alias for TextRange
    @property
    def textrange(self):
        return self.TextRange

    @property
    def ThreeD(self):
        return ThreeDFormat(self.textframe2.ThreeD)

    # Lower case alias for ThreeD
    @property
    def threed(self):
        return self.ThreeD

    @property
    def VerticalAnchor(self):
        return self.textframe2.VerticalAnchor

    # Lower case alias for VerticalAnchor
    @property
    def verticalanchor(self):
        return self.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.textframe2.VerticalAnchor = value

    # Lower case alias for VerticalAnchor setter
    @verticalanchor.setter
    def verticalanchor(self, value):
        self.VerticalAnchor = value

    @property
    def WarpFormat(self):
        return self.textframe2.WarpFormat

    # Lower case alias for WarpFormat
    @property
    def warpformat(self):
        return self.WarpFormat

    @WarpFormat.setter
    def WarpFormat(self, value):
        self.textframe2.WarpFormat = value

    # Lower case alias for WarpFormat setter
    @warpformat.setter
    def warpformat(self, value):
        self.WarpFormat = value

    @property
    def WordArtFormat(self):
        return self.textframe2.WordArtFormat

    # Lower case alias for WordArtFormat
    @property
    def wordartformat(self):
        return self.WordArtFormat

    @WordArtFormat.setter
    def WordArtFormat(self, value):
        self.textframe2.WordArtFormat = value

    # Lower case alias for WordArtFormat setter
    @wordartformat.setter
    def wordartformat(self, value):
        self.WordArtFormat = value

    @property
    def WordWrap(self):
        return self.textframe2.WordWrap

    # Lower case alias for WordWrap
    @property
    def wordwrap(self):
        return self.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.textframe2.WordWrap = value

    # Lower case alias for WordWrap setter
    @wordwrap.setter
    def wordwrap(self, value):
        self.WordWrap = value

    def DeleteText(self):
        return self.textframe2.DeleteText()


class TextRange:

    def __init__(self, textrange=None):
        self.textrange = textrange

    @property
    def ActionSettings(self):
        return ActionSettings(self.textrange.ActionSettings)

    # Lower case alias for ActionSettings
    @property
    def actionsettings(self):
        return self.ActionSettings

    @property
    def Application(self):
        return Application(self.textrange.Application)

    @property
    def BoundHeight(self):
        return self.textrange.BoundHeight

    # Lower case alias for BoundHeight
    @property
    def boundheight(self):
        return self.BoundHeight

    @property
    def BoundLeft(self):
        return self.textrange.BoundLeft

    # Lower case alias for BoundLeft
    @property
    def boundleft(self):
        return self.BoundLeft

    @property
    def BoundTop(self):
        return self.textrange.BoundTop

    # Lower case alias for BoundTop
    @property
    def boundtop(self):
        return self.BoundTop

    @property
    def BoundWidth(self):
        return self.textrange.BoundWidth

    # Lower case alias for BoundWidth
    @property
    def boundwidth(self):
        return self.BoundWidth

    @property
    def Count(self):
        return self.textrange.Count

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Font(self):
        return Font(self.textrange.Font)

    # Lower case alias for Font
    @property
    def font(self):
        return self.Font

    @property
    def IndentLevel(self):
        return self.textrange.IndentLevel

    # Lower case alias for IndentLevel
    @property
    def indentlevel(self):
        return self.IndentLevel

    @IndentLevel.setter
    def IndentLevel(self, value):
        self.textrange.IndentLevel = value

    # Lower case alias for IndentLevel setter
    @indentlevel.setter
    def indentlevel(self, value):
        self.IndentLevel = value

    @property
    def LanguageID(self):
        return self.textrange.LanguageID

    # Lower case alias for LanguageID
    @property
    def languageid(self):
        return self.LanguageID

    @LanguageID.setter
    def LanguageID(self, value):
        self.textrange.LanguageID = value

    # Lower case alias for LanguageID setter
    @languageid.setter
    def languageid(self, value):
        self.LanguageID = value

    @property
    def Length(self):
        return self.textrange.Length

    # Lower case alias for Length
    @property
    def length(self):
        return self.Length

    @property
    def ParagraphFormat(self):
        return ParagraphFormat(self.textrange.ParagraphFormat)

    # Lower case alias for ParagraphFormat
    @property
    def paragraphformat(self):
        return self.ParagraphFormat

    @property
    def Parent(self):
        return self.textrange.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Start(self):
        return self.textrange.Start

    # Lower case alias for Start
    @property
    def start(self):
        return self.Start

    @property
    def Text(self):
        return self.textrange.Text

    # Lower case alias for Text
    @property
    def text(self):
        return self.Text

    @Text.setter
    def Text(self, value):
        self.textrange.Text = value

    # Lower case alias for Text setter
    @text.setter
    def text(self, value):
        self.Text = value

    def AddPeriods(self):
        self.textrange.AddPeriods()

    def ChangeCase(self, Type=None):
        arguments = com_arguments([Type])
        self.textrange.ChangeCase(*arguments)

    def Characters(self, Start=None, Length=None):
        arguments = com_arguments([Start, Length])
        return self.textrange.Characters(*arguments)

    def Copy(self):
        self.textrange.Copy()

    def Cut(self):
        self.textrange.Cut()

    def Delete(self):
        self.textrange.Delete()

    def Find(self, FindWhat=None, After=None, MatchCase=None, WholeWords=None):
        arguments = com_arguments([FindWhat, After, MatchCase, WholeWords])
        return self.textrange.Find(*arguments)

    def InsertAfter(self, NewText=None):
        arguments = com_arguments([NewText])
        self.textrange.InsertAfter(*arguments)

    def InsertBefore(self, NewText=None):
        arguments = com_arguments([NewText])
        self.textrange.InsertBefore(*arguments)

    def InsertDateTime(self, DateTimeFormat=None, InsertAsField=None):
        arguments = com_arguments([DateTimeFormat, InsertAsField])
        return self.textrange.InsertDateTime(*arguments)

    def InsertSlideNumber(self):
        return self.textrange.InsertSlideNumber()

    def InsertSymbol(self, FontName=None, CharNumber=None, UniCode=None):
        arguments = com_arguments([FontName, CharNumber, UniCode])
        return self.textrange.InsertSymbol(*arguments)

    def Lines(self, Start=None, Length=None):
        arguments = com_arguments([Start, Length])
        return self.textrange.Lines(*arguments)

    def LtrRun(self):
        self.textrange.LtrRun()

    def Paragraphs(self, Start=None, Length=None):
        arguments = com_arguments([Start, Length])
        return self.textrange.Paragraphs(*arguments)

    def Paste(self):
        return self.textrange.Paste()

    def PasteSpecial(self, DataType=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Link=None):
        arguments = com_arguments([DataType, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Link])
        return self.textrange.PasteSpecial(*arguments)

    def RemovePeriods(self):
        self.textrange.RemovePeriods()

    def Replace(self, FindWhat=None, ReplaceWhat=None, After=None, MatchCase=None, WholeWords=None):
        arguments = com_arguments([FindWhat, ReplaceWhat, After, MatchCase, WholeWords])
        return self.textrange.Replace(*arguments)

    def RotatedBounds(self, X1=None, Y1=None, X2=None, Y2=None, X3=None, Y3=None, X4=None, Y4=None):
        arguments = com_arguments([X1, Y1, X2, Y2, X3, Y3, X4, Y4])
        self.textrange.RotatedBounds(*arguments)

    def RtlRun(self):
        self.textrange.RtlRun()

    def Runs(self, Start=None, Length=None):
        arguments = com_arguments([Start, Length])
        return self.textrange.Runs(*arguments)

    def Select(self):
        self.textrange.Select()

    def Sentences(self, Start=None, Length=None):
        arguments = com_arguments([Start, Length])
        return self.textrange.Sentences(*arguments)

    def TrimText(self):
        return self.textrange.TrimText()

    def Words(self, Start=None, Length=None):
        arguments = com_arguments([Start, Length])
        return self.textrange.Words(*arguments)


class TextStyle:

    def __init__(self, textstyle=None):
        self.textstyle = textstyle

    @property
    def Application(self):
        return Application(self.textstyle.Application)

    @property
    def Levels(self):
        return TextStyleLevels(self.textstyle.Levels)

    # Lower case alias for Levels
    @property
    def levels(self):
        return self.Levels

    @property
    def Parent(self):
        return self.textstyle.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Ruler(self):
        return Ruler(self.textstyle.Ruler)

    # Lower case alias for Ruler
    @property
    def ruler(self):
        return self.Ruler

    @property
    def TextFrame(self):
        return TextFrame(self.textstyle.TextFrame)

    # Lower case alias for TextFrame
    @property
    def textframe(self):
        return self.TextFrame


class TextStyleLevel:

    def __init__(self, textstylelevel=None):
        self.textstylelevel = textstylelevel

    @property
    def Application(self):
        return Application(self.textstylelevel.Application)

    @property
    def Font(self):
        return Font(self.textstylelevel.Font)

    # Lower case alias for Font
    @property
    def font(self):
        return self.Font

    @property
    def ParagraphFormat(self):
        return ParagraphFormat(self.textstylelevel.ParagraphFormat)

    # Lower case alias for ParagraphFormat
    @property
    def paragraphformat(self):
        return self.ParagraphFormat

    @property
    def Parent(self):
        return self.textstylelevel.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.textstylelevels.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.textstylelevels.Item(*arguments)


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.textstyles.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Item(self, Type=None):
        arguments = com_arguments([Type])
        return self.textstyles.Item(*arguments)


class ThreeDFormat:

    def __init__(self, threedformat=None):
        self.threedformat = threedformat

    @property
    def Application(self):
        return Application(self.threedformat.Application)

    @property
    def BevelBottomDepth(self):
        return ThreeDFormat(self.threedformat.BevelBottomDepth)

    # Lower case alias for BevelBottomDepth
    @property
    def bevelbottomdepth(self):
        return self.BevelBottomDepth

    @BevelBottomDepth.setter
    def BevelBottomDepth(self, value):
        self.threedformat.BevelBottomDepth = value

    # Lower case alias for BevelBottomDepth setter
    @bevelbottomdepth.setter
    def bevelbottomdepth(self, value):
        self.BevelBottomDepth = value

    @property
    def BevelBottomInset(self):
        return ThreeDFormat(self.threedformat.BevelBottomInset)

    # Lower case alias for BevelBottomInset
    @property
    def bevelbottominset(self):
        return self.BevelBottomInset

    @BevelBottomInset.setter
    def BevelBottomInset(self, value):
        self.threedformat.BevelBottomInset = value

    # Lower case alias for BevelBottomInset setter
    @bevelbottominset.setter
    def bevelbottominset(self, value):
        self.BevelBottomInset = value

    @property
    def BevelBottomType(self):
        return self.threedformat.BevelBottomType

    # Lower case alias for BevelBottomType
    @property
    def bevelbottomtype(self):
        return self.BevelBottomType

    @BevelBottomType.setter
    def BevelBottomType(self, value):
        self.threedformat.BevelBottomType = value

    # Lower case alias for BevelBottomType setter
    @bevelbottomtype.setter
    def bevelbottomtype(self, value):
        self.BevelBottomType = value

    @property
    def BevelTopDepth(self):
        return ThreeDFormat(self.threedformat.BevelTopDepth)

    # Lower case alias for BevelTopDepth
    @property
    def beveltopdepth(self):
        return self.BevelTopDepth

    @BevelTopDepth.setter
    def BevelTopDepth(self, value):
        self.threedformat.BevelTopDepth = value

    # Lower case alias for BevelTopDepth setter
    @beveltopdepth.setter
    def beveltopdepth(self, value):
        self.BevelTopDepth = value

    @property
    def BevelTopInset(self):
        return ThreeDFormat(self.threedformat.BevelTopInset)

    # Lower case alias for BevelTopInset
    @property
    def beveltopinset(self):
        return self.BevelTopInset

    @BevelTopInset.setter
    def BevelTopInset(self, value):
        self.threedformat.BevelTopInset = value

    # Lower case alias for BevelTopInset setter
    @beveltopinset.setter
    def beveltopinset(self, value):
        self.BevelTopInset = value

    @property
    def BevelTopType(self):
        return self.threedformat.BevelTopType

    # Lower case alias for BevelTopType
    @property
    def beveltoptype(self):
        return self.BevelTopType

    @BevelTopType.setter
    def BevelTopType(self, value):
        self.threedformat.BevelTopType = value

    # Lower case alias for BevelTopType setter
    @beveltoptype.setter
    def beveltoptype(self, value):
        self.BevelTopType = value

    @property
    def ContourColor(self):
        return ColorFormat(self.threedformat.ContourColor)

    # Lower case alias for ContourColor
    @property
    def contourcolor(self):
        return self.ContourColor

    @property
    def ContourWidth(self):
        return ThreeDFormat(self.threedformat.ContourWidth)

    # Lower case alias for ContourWidth
    @property
    def contourwidth(self):
        return self.ContourWidth

    @ContourWidth.setter
    def ContourWidth(self, value):
        self.threedformat.ContourWidth = value

    # Lower case alias for ContourWidth setter
    @contourwidth.setter
    def contourwidth(self, value):
        self.ContourWidth = value

    @property
    def Creator(self):
        return self.threedformat.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Depth(self):
        return self.threedformat.Depth

    # Lower case alias for Depth
    @property
    def depth(self):
        return self.Depth

    @Depth.setter
    def Depth(self, value):
        self.threedformat.Depth = value

    # Lower case alias for Depth setter
    @depth.setter
    def depth(self, value):
        self.Depth = value

    @property
    def ExtrusionColor(self):
        return ColorFormat(self.threedformat.ExtrusionColor)

    # Lower case alias for ExtrusionColor
    @property
    def extrusioncolor(self):
        return self.ExtrusionColor

    @property
    def ExtrusionColorType(self):
        return self.threedformat.ExtrusionColorType

    # Lower case alias for ExtrusionColorType
    @property
    def extrusioncolortype(self):
        return self.ExtrusionColorType

    @ExtrusionColorType.setter
    def ExtrusionColorType(self, value):
        self.threedformat.ExtrusionColorType = value

    # Lower case alias for ExtrusionColorType setter
    @extrusioncolortype.setter
    def extrusioncolortype(self, value):
        self.ExtrusionColorType = value

    @property
    def FieldOfView(self):
        return ThreeDFormat(self.threedformat.FieldOfView)

    # Lower case alias for FieldOfView
    @property
    def fieldofview(self):
        return self.FieldOfView

    @FieldOfView.setter
    def FieldOfView(self, value):
        self.threedformat.FieldOfView = value

    # Lower case alias for FieldOfView setter
    @fieldofview.setter
    def fieldofview(self, value):
        self.FieldOfView = value

    @property
    def LightAngle(self):
        return self.threedformat.LightAngle

    # Lower case alias for LightAngle
    @property
    def lightangle(self):
        return self.LightAngle

    @LightAngle.setter
    def LightAngle(self, value):
        self.threedformat.LightAngle = value

    # Lower case alias for LightAngle setter
    @lightangle.setter
    def lightangle(self, value):
        self.LightAngle = value

    @property
    def Parent(self):
        return self.threedformat.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Perspective(self):
        return self.threedformat.Perspective

    # Lower case alias for Perspective
    @property
    def perspective(self):
        return self.Perspective

    @Perspective.setter
    def Perspective(self, value):
        self.threedformat.Perspective = value

    # Lower case alias for Perspective setter
    @perspective.setter
    def perspective(self, value):
        self.Perspective = value

    @property
    def PresetCamera(self):
        return ThreeDFormat(self.threedformat.PresetCamera)

    # Lower case alias for PresetCamera
    @property
    def presetcamera(self):
        return self.PresetCamera

    @property
    def PresetExtrusionDirection(self):
        return self.threedformat.PresetExtrusionDirection

    # Lower case alias for PresetExtrusionDirection
    @property
    def presetextrusiondirection(self):
        return self.PresetExtrusionDirection

    @property
    def PresetLighting(self):
        return ThreeDFormat(self.threedformat.PresetLighting)

    # Lower case alias for PresetLighting
    @property
    def presetlighting(self):
        return self.PresetLighting

    @PresetLighting.setter
    def PresetLighting(self, value):
        self.threedformat.PresetLighting = value

    # Lower case alias for PresetLighting setter
    @presetlighting.setter
    def presetlighting(self, value):
        self.PresetLighting = value

    @property
    def PresetLightingDirection(self):
        return self.threedformat.PresetLightingDirection

    # Lower case alias for PresetLightingDirection
    @property
    def presetlightingdirection(self):
        return self.PresetLightingDirection

    @PresetLightingDirection.setter
    def PresetLightingDirection(self, value):
        self.threedformat.PresetLightingDirection = value

    # Lower case alias for PresetLightingDirection setter
    @presetlightingdirection.setter
    def presetlightingdirection(self, value):
        self.PresetLightingDirection = value

    @property
    def PresetLightingSoftness(self):
        return self.threedformat.PresetLightingSoftness

    # Lower case alias for PresetLightingSoftness
    @property
    def presetlightingsoftness(self):
        return self.PresetLightingSoftness

    @PresetLightingSoftness.setter
    def PresetLightingSoftness(self, value):
        self.threedformat.PresetLightingSoftness = value

    # Lower case alias for PresetLightingSoftness setter
    @presetlightingsoftness.setter
    def presetlightingsoftness(self, value):
        self.PresetLightingSoftness = value

    @property
    def PresetMaterial(self):
        return self.threedformat.PresetMaterial

    # Lower case alias for PresetMaterial
    @property
    def presetmaterial(self):
        return self.PresetMaterial

    @PresetMaterial.setter
    def PresetMaterial(self, value):
        self.threedformat.PresetMaterial = value

    # Lower case alias for PresetMaterial setter
    @presetmaterial.setter
    def presetmaterial(self, value):
        self.PresetMaterial = value

    @property
    def PresetThreeDFormat(self):
        return self.threedformat.PresetThreeDFormat

    # Lower case alias for PresetThreeDFormat
    @property
    def presetthreedformat(self):
        return self.PresetThreeDFormat

    @property
    def ProjectText(self):
        return self.threedformat.ProjectText

    # Lower case alias for ProjectText
    @property
    def projecttext(self):
        return self.ProjectText

    @ProjectText.setter
    def ProjectText(self, value):
        self.threedformat.ProjectText = value

    # Lower case alias for ProjectText setter
    @projecttext.setter
    def projecttext(self, value):
        self.ProjectText = value

    @property
    def RotationX(self):
        return self.threedformat.RotationX

    # Lower case alias for RotationX
    @property
    def rotationx(self):
        return self.RotationX

    @RotationX.setter
    def RotationX(self, value):
        self.threedformat.RotationX = value

    # Lower case alias for RotationX setter
    @rotationx.setter
    def rotationx(self, value):
        self.RotationX = value

    @property
    def RotationY(self):
        return self.threedformat.RotationY

    # Lower case alias for RotationY
    @property
    def rotationy(self):
        return self.RotationY

    @RotationY.setter
    def RotationY(self, value):
        self.threedformat.RotationY = value

    # Lower case alias for RotationY setter
    @rotationy.setter
    def rotationy(self, value):
        self.RotationY = value

    @property
    def RotationZ(self):
        return ThreeDFormat(self.threedformat.RotationZ)

    # Lower case alias for RotationZ
    @property
    def rotationz(self):
        return self.RotationZ

    @RotationZ.setter
    def RotationZ(self, value):
        self.threedformat.RotationZ = value

    # Lower case alias for RotationZ setter
    @rotationz.setter
    def rotationz(self, value):
        self.RotationZ = value

    @property
    def Visible(self):
        return self.threedformat.Visible

    # Lower case alias for Visible
    @property
    def visible(self):
        return self.Visible

    @Visible.setter
    def Visible(self, value):
        self.threedformat.Visible = value

    # Lower case alias for Visible setter
    @visible.setter
    def visible(self, value):
        self.Visible = value

    @property
    def Z(self):
        return ThreeDFormat(self.threedformat.Z)

    # Lower case alias for Z
    @property
    def z(self):
        return self.Z

    @Z.setter
    def Z(self, value):
        self.threedformat.Z = value

    # Lower case alias for Z setter
    @z.setter
    def z(self, value):
        self.Z = value

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

    # Lower case alias for Alignment
    @property
    def alignment(self):
        return self.Alignment

    @Alignment.setter
    def Alignment(self, value):
        self.ticklabels.Alignment = value

    # Lower case alias for Alignment setter
    @alignment.setter
    def alignment(self, value):
        self.Alignment = value

    @property
    def Application(self):
        return self.ticklabels.Application

    @property
    def Creator(self):
        return self.ticklabels.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Depth(self):
        return self.ticklabels.Depth

    # Lower case alias for Depth
    @property
    def depth(self):
        return self.Depth

    @property
    def Font(self):
        return ChartFont(self.ticklabels.Font)

    # Lower case alias for Font
    @property
    def font(self):
        return self.Font

    @property
    def Format(self):
        return ChartFormat(self.ticklabels.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def MultiLevel(self):
        return self.ticklabels.MultiLevel

    # Lower case alias for MultiLevel
    @property
    def multilevel(self):
        return self.MultiLevel

    @MultiLevel.setter
    def MultiLevel(self, value):
        self.ticklabels.MultiLevel = value

    # Lower case alias for MultiLevel setter
    @multilevel.setter
    def multilevel(self, value):
        self.MultiLevel = value

    @property
    def Name(self):
        return self.ticklabels.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def NumberFormat(self):
        return self.ticklabels.NumberFormat

    # Lower case alias for NumberFormat
    @property
    def numberformat(self):
        return self.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value):
        self.ticklabels.NumberFormat = value

    # Lower case alias for NumberFormat setter
    @numberformat.setter
    def numberformat(self, value):
        self.NumberFormat = value

    @property
    def NumberFormatLinked(self):
        return self.ticklabels.NumberFormatLinked

    # Lower case alias for NumberFormatLinked
    @property
    def numberformatlinked(self):
        return self.NumberFormatLinked

    @NumberFormatLinked.setter
    def NumberFormatLinked(self, value):
        self.ticklabels.NumberFormatLinked = value

    # Lower case alias for NumberFormatLinked setter
    @numberformatlinked.setter
    def numberformatlinked(self, value):
        self.NumberFormatLinked = value

    @property
    def NumberFormatLocal(self):
        return self.ticklabels.NumberFormatLocal

    # Lower case alias for NumberFormatLocal
    @property
    def numberformatlocal(self):
        return self.NumberFormatLocal

    @NumberFormatLocal.setter
    def NumberFormatLocal(self, value):
        self.ticklabels.NumberFormatLocal = value

    # Lower case alias for NumberFormatLocal setter
    @numberformatlocal.setter
    def numberformatlocal(self, value):
        self.NumberFormatLocal = value

    @property
    def Offset(self):
        return self.ticklabels.Offset

    # Lower case alias for Offset
    @property
    def offset(self):
        return self.Offset

    @Offset.setter
    def Offset(self, value):
        self.ticklabels.Offset = value

    # Lower case alias for Offset setter
    @offset.setter
    def offset(self, value):
        self.Offset = value

    @property
    def Orientation(self):
        return self.ticklabels.Orientation

    # Lower case alias for Orientation
    @property
    def orientation(self):
        return self.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.ticklabels.Orientation = value

    # Lower case alias for Orientation setter
    @orientation.setter
    def orientation(self, value):
        self.Orientation = value

    @property
    def Parent(self):
        return self.ticklabels.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ReadingOrder(self):
        return XlReadingOrder(self.ticklabels.ReadingOrder)

    # Lower case alias for ReadingOrder
    @property
    def readingorder(self):
        return self.ReadingOrder

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.ticklabels.ReadingOrder = value

    # Lower case alias for ReadingOrder setter
    @readingorder.setter
    def readingorder(self, value):
        self.ReadingOrder = value

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

    # Lower case alias for InteractiveSequences
    @property
    def interactivesequences(self):
        return self.InteractiveSequences

    @property
    def MainSequence(self):
        return Sequence(self.timeline.MainSequence)

    # Lower case alias for MainSequence
    @property
    def mainsequence(self):
        return self.MainSequence

    @property
    def Parent(self):
        return self.timeline.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent


class Timing:

    def __init__(self, timing=None):
        self.timing = timing

    @property
    def Accelerate(self):
        return self.timing.Accelerate

    # Lower case alias for Accelerate
    @property
    def accelerate(self):
        return self.Accelerate

    @Accelerate.setter
    def Accelerate(self, value):
        self.timing.Accelerate = value

    # Lower case alias for Accelerate setter
    @accelerate.setter
    def accelerate(self, value):
        self.Accelerate = value

    @property
    def Application(self):
        return Application(self.timing.Application)

    @property
    def AutoReverse(self):
        return self.timing.AutoReverse

    # Lower case alias for AutoReverse
    @property
    def autoreverse(self):
        return self.AutoReverse

    @AutoReverse.setter
    def AutoReverse(self, value):
        self.timing.AutoReverse = value

    # Lower case alias for AutoReverse setter
    @autoreverse.setter
    def autoreverse(self, value):
        self.AutoReverse = value

    @property
    def BounceEnd(self):
        return self.timing.BounceEnd

    # Lower case alias for BounceEnd
    @property
    def bounceend(self):
        return self.BounceEnd

    @BounceEnd.setter
    def BounceEnd(self, value):
        self.timing.BounceEnd = value

    # Lower case alias for BounceEnd setter
    @bounceend.setter
    def bounceend(self, value):
        self.BounceEnd = value

    @property
    def BounceEndIntensity(self):
        return self.timing.BounceEndIntensity

    # Lower case alias for BounceEndIntensity
    @property
    def bounceendintensity(self):
        return self.BounceEndIntensity

    @BounceEndIntensity.setter
    def BounceEndIntensity(self, value):
        self.timing.BounceEndIntensity = value

    # Lower case alias for BounceEndIntensity setter
    @bounceendintensity.setter
    def bounceendintensity(self, value):
        self.BounceEndIntensity = value

    @property
    def Decelerate(self):
        return self.timing.Decelerate

    # Lower case alias for Decelerate
    @property
    def decelerate(self):
        return self.Decelerate

    @Decelerate.setter
    def Decelerate(self, value):
        self.timing.Decelerate = value

    # Lower case alias for Decelerate setter
    @decelerate.setter
    def decelerate(self, value):
        self.Decelerate = value

    @property
    def Parent(self):
        return self.timing.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RepeatCount(self):
        return self.timing.RepeatCount

    # Lower case alias for RepeatCount
    @property
    def repeatcount(self):
        return self.RepeatCount

    @RepeatCount.setter
    def RepeatCount(self, value):
        self.timing.RepeatCount = value

    # Lower case alias for RepeatCount setter
    @repeatcount.setter
    def repeatcount(self, value):
        self.RepeatCount = value

    @property
    def RepeatDuration(self):
        return self.timing.RepeatDuration

    # Lower case alias for RepeatDuration
    @property
    def repeatduration(self):
        return self.RepeatDuration

    @RepeatDuration.setter
    def RepeatDuration(self, value):
        self.timing.RepeatDuration = value

    # Lower case alias for RepeatDuration setter
    @repeatduration.setter
    def repeatduration(self, value):
        self.RepeatDuration = value

    @property
    def Restart(self):
        return self.timing.Restart

    # Lower case alias for Restart
    @property
    def restart(self):
        return self.Restart

    @Restart.setter
    def Restart(self, value):
        self.timing.Restart = value

    # Lower case alias for Restart setter
    @restart.setter
    def restart(self, value):
        self.Restart = value

    @property
    def RewindAtEnd(self):
        return self.timing.RewindAtEnd

    # Lower case alias for RewindAtEnd
    @property
    def rewindatend(self):
        return self.RewindAtEnd

    @RewindAtEnd.setter
    def RewindAtEnd(self, value):
        self.timing.RewindAtEnd = value

    # Lower case alias for RewindAtEnd setter
    @rewindatend.setter
    def rewindatend(self, value):
        self.RewindAtEnd = value

    @property
    def SmoothEnd(self):
        return self.timing.SmoothEnd

    # Lower case alias for SmoothEnd
    @property
    def smoothend(self):
        return self.SmoothEnd

    @SmoothEnd.setter
    def SmoothEnd(self, value):
        self.timing.SmoothEnd = value

    # Lower case alias for SmoothEnd setter
    @smoothend.setter
    def smoothend(self, value):
        self.SmoothEnd = value

    @property
    def SmoothStart(self):
        return self.timing.SmoothStart

    # Lower case alias for SmoothStart
    @property
    def smoothstart(self):
        return self.SmoothStart

    @SmoothStart.setter
    def SmoothStart(self, value):
        self.timing.SmoothStart = value

    # Lower case alias for SmoothStart setter
    @smoothstart.setter
    def smoothstart(self, value):
        self.SmoothStart = value

    @property
    def Speed(self):
        return self.timing.Speed

    # Lower case alias for Speed
    @property
    def speed(self):
        return self.Speed

    @Speed.setter
    def Speed(self, value):
        self.timing.Speed = value

    # Lower case alias for Speed setter
    @speed.setter
    def speed(self, value):
        self.Speed = value

    @property
    def triggerBookmark(self):
        return self.timing.triggerBookmark

    # Lower case alias for triggerBookmark
    @property
    def triggerbookmark(self):
        return self.triggerBookmark

    @triggerBookmark.setter
    def triggerBookmark(self, value):
        self.timing.triggerBookmark = value

    # Lower case alias for triggerBookmark setter
    @triggerbookmark.setter
    def triggerbookmark(self, value):
        self.triggerBookmark = value

    @property
    def TriggerDelayTime(self):
        return self.timing.TriggerDelayTime

    # Lower case alias for TriggerDelayTime
    @property
    def triggerdelaytime(self):
        return self.TriggerDelayTime

    @TriggerDelayTime.setter
    def TriggerDelayTime(self, value):
        self.timing.TriggerDelayTime = value

    # Lower case alias for TriggerDelayTime setter
    @triggerdelaytime.setter
    def triggerdelaytime(self, value):
        self.TriggerDelayTime = value

    @property
    def TriggerShape(self):
        return self.timing.TriggerShape

    # Lower case alias for TriggerShape
    @property
    def triggershape(self):
        return self.TriggerShape

    @TriggerShape.setter
    def TriggerShape(self, value):
        self.timing.TriggerShape = value

    # Lower case alias for TriggerShape setter
    @triggershape.setter
    def triggershape(self, value):
        self.TriggerShape = value

    @property
    def TriggerType(self):
        return self.timing.TriggerType

    # Lower case alias for TriggerType
    @property
    def triggertype(self):
        return self.TriggerType

    @TriggerType.setter
    def TriggerType(self, value):
        self.timing.TriggerType = value

    # Lower case alias for TriggerType setter
    @triggertype.setter
    def triggertype(self, value):
        self.TriggerType = value


class Trendline:

    def __init__(self, trendline=None):
        self.trendline = trendline

    @property
    def Application(self):
        return self.trendline.Application

    @property
    def Backward2(self):
        return self.trendline.Backward2

    # Lower case alias for Backward2
    @property
    def backward2(self):
        return self.Backward2

    @Backward2.setter
    def Backward2(self, value):
        self.trendline.Backward2 = value

    # Lower case alias for Backward2 setter
    @backward2.setter
    def backward2(self, value):
        self.Backward2 = value

    @property
    def Border(self):
        return ChartBorder(self.trendline.Border)

    # Lower case alias for Border
    @property
    def border(self):
        return self.Border

    @property
    def Creator(self):
        return self.trendline.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def DataLabel(self):
        return DataLabel(self.trendline.DataLabel)

    # Lower case alias for DataLabel
    @property
    def datalabel(self):
        return self.DataLabel

    @property
    def DisplayEquation(self):
        return self.trendline.DisplayEquation

    # Lower case alias for DisplayEquation
    @property
    def displayequation(self):
        return self.DisplayEquation

    @DisplayEquation.setter
    def DisplayEquation(self, value):
        self.trendline.DisplayEquation = value

    # Lower case alias for DisplayEquation setter
    @displayequation.setter
    def displayequation(self, value):
        self.DisplayEquation = value

    @property
    def DisplayRSquared(self):
        return self.trendline.DisplayRSquared

    # Lower case alias for DisplayRSquared
    @property
    def displayrsquared(self):
        return self.DisplayRSquared

    @DisplayRSquared.setter
    def DisplayRSquared(self, value):
        self.trendline.DisplayRSquared = value

    # Lower case alias for DisplayRSquared setter
    @displayrsquared.setter
    def displayrsquared(self, value):
        self.DisplayRSquared = value

    @property
    def Format(self):
        return ChartFormat(self.trendline.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def Forward2(self):
        return self.trendline.Forward2

    # Lower case alias for Forward2
    @property
    def forward2(self):
        return self.Forward2

    @Forward2.setter
    def Forward2(self, value):
        self.trendline.Forward2 = value

    # Lower case alias for Forward2 setter
    @forward2.setter
    def forward2(self, value):
        self.Forward2 = value

    @property
    def Index(self):
        return self.trendline.Index

    # Lower case alias for Index
    @property
    def index(self):
        return self.Index

    @property
    def Intercept(self):
        return self.trendline.Intercept

    # Lower case alias for Intercept
    @property
    def intercept(self):
        return self.Intercept

    @Intercept.setter
    def Intercept(self, value):
        self.trendline.Intercept = value

    # Lower case alias for Intercept setter
    @intercept.setter
    def intercept(self, value):
        self.Intercept = value

    @property
    def InterceptIsAuto(self):
        return self.trendline.InterceptIsAuto

    # Lower case alias for InterceptIsAuto
    @property
    def interceptisauto(self):
        return self.InterceptIsAuto

    @InterceptIsAuto.setter
    def InterceptIsAuto(self, value):
        self.trendline.InterceptIsAuto = value

    # Lower case alias for InterceptIsAuto setter
    @interceptisauto.setter
    def interceptisauto(self, value):
        self.InterceptIsAuto = value

    @property
    def Name(self):
        return self.trendline.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.trendline.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def NameIsAuto(self):
        return self.trendline.NameIsAuto

    # Lower case alias for NameIsAuto
    @property
    def nameisauto(self):
        return self.NameIsAuto

    @NameIsAuto.setter
    def NameIsAuto(self, value):
        self.trendline.NameIsAuto = value

    # Lower case alias for NameIsAuto setter
    @nameisauto.setter
    def nameisauto(self, value):
        self.NameIsAuto = value

    @property
    def Order(self):
        return self.trendline.Order

    # Lower case alias for Order
    @property
    def order(self):
        return self.Order

    @Order.setter
    def Order(self, value):
        self.trendline.Order = value

    # Lower case alias for Order setter
    @order.setter
    def order(self, value):
        self.Order = value

    @property
    def Parent(self):
        return self.trendline.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Period(self):
        return self.trendline.Period

    # Lower case alias for Period
    @property
    def period(self):
        return self.Period

    @Period.setter
    def Period(self, value):
        self.trendline.Period = value

    # Lower case alias for Period setter
    @period.setter
    def period(self, value):
        self.Period = value

    @property
    def Type(self):
        return XlTrendlineType(self.trendline.Type)

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @Type.setter
    def Type(self, value):
        self.trendline.Type = value

    # Lower case alias for Type setter
    @type.setter
    def type(self, value):
        self.Type = value

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.trendlines.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Parent(self):
        return self.trendlines.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    def Add(self, Type=None, Order=None, Period=None, Forward=None, Backward=None, Intercept=None, DisplayEquation=None, DisplayRSquared=None, Name=None):
        arguments = com_arguments([Type, Order, Period, Forward, Backward, Intercept, DisplayEquation, DisplayRSquared, Name])
        return Trendline(self.trendlines.Add(*arguments))

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Trendline(self.trendlines.Item(*arguments))


class UpBars:

    def __init__(self, upbars=None):
        self.upbars = upbars

    @property
    def Application(self):
        return self.upbars.Application

    @property
    def Creator(self):
        return self.upbars.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Fill(self):
        return FillFormat(self.upbars.Fill)

    # Lower case alias for Fill
    @property
    def fill(self):
        return self.Fill

    @property
    def Format(self):
        return ChartFormat(self.upbars.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def Name(self):
        return self.upbars.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.upbars.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

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

    # Lower case alias for DisplaySlideMiniature
    @property
    def displayslideminiature(self):
        return self.DisplaySlideMiniature

    @DisplaySlideMiniature.setter
    def DisplaySlideMiniature(self, value):
        self.view.DisplaySlideMiniature = value

    # Lower case alias for DisplaySlideMiniature setter
    @displayslideminiature.setter
    def displayslideminiature(self, value):
        self.DisplaySlideMiniature = value

    @property
    def MediaControlsHeight(self):
        return self.view.MediaControlsHeight

    # Lower case alias for MediaControlsHeight
    @property
    def mediacontrolsheight(self):
        return self.MediaControlsHeight

    @property
    def MediaControlsLeft(self):
        return self.view.MediaControlsLeft

    # Lower case alias for MediaControlsLeft
    @property
    def mediacontrolsleft(self):
        return self.MediaControlsLeft

    @property
    def MediaControlsTop(self):
        return self.view.MediaControlsTop

    # Lower case alias for MediaControlsTop
    @property
    def mediacontrolstop(self):
        return self.MediaControlsTop

    @property
    def MediaControlsVisible(self):
        return self.view.MediaControlsVisible

    # Lower case alias for MediaControlsVisible
    @property
    def mediacontrolsvisible(self):
        return self.MediaControlsVisible

    @property
    def MediaControlsWidth(self):
        return self.view.MediaControlsWidth

    # Lower case alias for MediaControlsWidth
    @property
    def mediacontrolswidth(self):
        return self.MediaControlsWidth

    @property
    def Parent(self):
        return self.view.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PrintOptions(self):
        return PrintOptions(self.view.PrintOptions)

    # Lower case alias for PrintOptions
    @property
    def printoptions(self):
        return self.PrintOptions

    @property
    def Slide(self):
        return Slide(self.view.Slide)

    # Lower case alias for Slide
    @property
    def slide(self):
        return self.Slide

    @Slide.setter
    def Slide(self, value):
        self.view.Slide = value

    # Lower case alias for Slide setter
    @slide.setter
    def slide(self, value):
        self.Slide = value

    @property
    def Type(self):
        return self.view.Type

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @property
    def Zoom(self):
        return self.view.Zoom

    # Lower case alias for Zoom
    @property
    def zoom(self):
        return self.Zoom

    @Zoom.setter
    def Zoom(self, value):
        self.view.Zoom = value

    # Lower case alias for Zoom setter
    @zoom.setter
    def zoom(self, value):
        self.Zoom = value

    @property
    def ZoomToFit(self):
        return self.view.ZoomToFit

    # Lower case alias for ZoomToFit
    @property
    def zoomtofit(self):
        return self.ZoomToFit

    @ZoomToFit.setter
    def ZoomToFit(self, value):
        self.view.ZoomToFit = value

    # Lower case alias for ZoomToFit setter
    @zoomtofit.setter
    def zoomtofit(self, value):
        self.ZoomToFit = value

    def GotoSlide(self, Index=None):
        arguments = com_arguments([Index])
        self.view.GotoSlide(*arguments)

    def Paste(self):
        self.view.Paste()

    def PasteSpecial(self, DataType=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Link=None):
        arguments = com_arguments([DataType, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Link])
        self.view.PasteSpecial(*arguments)

    def Player(self, ShapeId=None):
        arguments = com_arguments([ShapeId])
        return self.view.Player(*arguments)

    def PrintOut(self, From=None, To=None, PrintToFile=None, Copies=None, Collate=None):
        arguments = com_arguments([From, To, PrintToFile, Copies, Collate])
        self.view.PrintOut(*arguments)


class Walls:

    def __init__(self, walls=None):
        self.walls = walls

    @property
    def Application(self):
        return self.walls.Application

    @property
    def Creator(self):
        return self.walls.Creator

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.walls.Format)

    # Lower case alias for Format
    @property
    def format(self):
        return self.Format

    @property
    def Name(self):
        return self.walls.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.walls.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PictureType(self):
        return self.walls.PictureType

    # Lower case alias for PictureType
    @property
    def picturetype(self):
        return self.PictureType

    @PictureType.setter
    def PictureType(self, value):
        self.walls.PictureType = value

    # Lower case alias for PictureType setter
    @picturetype.setter
    def picturetype(self, value):
        self.PictureType = value

    @property
    def PictureUnit(self):
        return self.walls.PictureUnit

    # Lower case alias for PictureUnit
    @property
    def pictureunit(self):
        return self.PictureUnit

    @PictureUnit.setter
    def PictureUnit(self, value):
        self.walls.PictureUnit = value

    # Lower case alias for PictureUnit setter
    @pictureunit.setter
    def pictureunit(self, value):
        self.PictureUnit = value

    @property
    def Thickness(self):
        return self.walls.Thickness

    # Lower case alias for Thickness
    @property
    def thickness(self):
        return self.Thickness

    @Thickness.setter
    def Thickness(self, value):
        self.walls.Thickness = value

    # Lower case alias for Thickness setter
    @thickness.setter
    def thickness(self, value):
        self.Thickness = value

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
