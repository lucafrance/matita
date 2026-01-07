from . import com_arguments, unwrap
from .office import *

import win32com.client

class ActionSetting:

    def __init__(self, actionsetting=None):
        self.com_object= actionsetting

    @property
    def Action(self):
        return self.com_object.Action

    @Action.setter
    def Action(self, value):
        self.com_object.Action = value

    # Lower case aliases for Action
    @property
    def action(self):
        return self.Action

    @action.setter
    def action(self, value):
        self.Action = value

    @property
    def ActionVerb(self):
        return self.com_object.ActionVerb

    @ActionVerb.setter
    def ActionVerb(self, value):
        self.com_object.ActionVerb = value

    # Lower case aliases for ActionVerb
    @property
    def actionverb(self):
        return self.ActionVerb

    @actionverb.setter
    def actionverb(self, value):
        self.ActionVerb = value

    @property
    def AnimateAction(self):
        return self.com_object.AnimateAction

    @AnimateAction.setter
    def AnimateAction(self, value):
        self.com_object.AnimateAction = value

    # Lower case aliases for AnimateAction
    @property
    def animateaction(self):
        return self.AnimateAction

    @animateaction.setter
    def animateaction(self, value):
        self.AnimateAction = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Hyperlink(self):
        return Hyperlink(self.com_object.Hyperlink)

    # Lower case aliases for Hyperlink
    @property
    def hyperlink(self):
        return self.Hyperlink

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Run(self):
        return self.com_object.Run

    @Run.setter
    def Run(self, value):
        self.com_object.Run = value

    # Lower case aliases for Run
    @property
    def run(self):
        return self.Run

    @run.setter
    def run(self, value):
        self.Run = value

    @property
    def ShowAndReturn(self):
        return self.com_object.ShowAndReturn

    @ShowAndReturn.setter
    def ShowAndReturn(self, value):
        self.com_object.ShowAndReturn = value

    # Lower case aliases for ShowAndReturn
    @property
    def showandreturn(self):
        return self.ShowAndReturn

    @showandreturn.setter
    def showandreturn(self, value):
        self.ShowAndReturn = value

    @property
    def SlideShowName(self):
        return self.com_object.SlideShowName

    @SlideShowName.setter
    def SlideShowName(self, value):
        self.com_object.SlideShowName = value

    # Lower case aliases for SlideShowName
    @property
    def slideshowname(self):
        return self.SlideShowName

    @slideshowname.setter
    def slideshowname(self, value):
        self.SlideShowName = value

    @property
    def SoundEffect(self):
        return SoundEffect(self.com_object.SoundEffect)

    # Lower case aliases for SoundEffect
    @property
    def soundeffect(self):
        return self.SoundEffect


class ActionSettings:

    def __init__(self, actionsettings=None):
        self.com_object= actionsettings

    def __call__(self, item):
        return ActionSetting(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return ActionSetting(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class AddIn:

    def __init__(self, addin=None):
        self.com_object= addin

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AutoLoad(self):
        return self.com_object.AutoLoad

    @AutoLoad.setter
    def AutoLoad(self, value):
        self.com_object.AutoLoad = value

    # Lower case aliases for AutoLoad
    @property
    def autoload(self):
        return self.AutoLoad

    @autoload.setter
    def autoload(self, value):
        self.AutoLoad = value

    @property
    def FullName(self):
        return self.com_object.FullName

    # Lower case aliases for FullName
    @property
    def fullname(self):
        return self.FullName

    @property
    def Loaded(self):
        return self.com_object.Loaded

    @Loaded.setter
    def Loaded(self, value):
        self.com_object.Loaded = value

    # Lower case aliases for Loaded
    @property
    def loaded(self):
        return self.Loaded

    @loaded.setter
    def loaded(self, value):
        self.Loaded = value

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
    def Path(self):
        return AddIn(self.com_object.Path)

    # Lower case aliases for Path
    @property
    def path(self):
        return self.Path

    @property
    def Registered(self):
        return self.com_object.Registered

    @Registered.setter
    def Registered(self, value):
        self.com_object.Registered = value

    # Lower case aliases for Registered
    @property
    def registered(self):
        return self.Registered

    @registered.setter
    def registered(self, value):
        self.Registered = value


class AddIns:

    def __init__(self, addins=None):
        self.com_object= addins

    def __call__(self, item):
        return AddIn(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Add(self, FileName=None):
        arguments = com_arguments([unwrap(a) for a in [FileName]])
        return AddIn(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, FileName=None):
        arguments = [FileName]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return AddIn(self.com_object.Item(*arguments))

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


class Adjustments:

    def __init__(self, adjustments=None):
        self.com_object= adjustments

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Item(self):
        return self.com_object.Item

    @Item.setter
    def Item(self, value):
        self.com_object.Item = value

    # Lower case aliases for Item
    @property
    def item(self):
        return self.Item

    @item.setter
    def item(self, value):
        self.Item = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent


class AnimationBehavior:

    def __init__(self, animationbehavior=None):
        self.com_object= animationbehavior

    @property
    def Accumulate(self):
        return self.com_object.Accumulate

    @Accumulate.setter
    def Accumulate(self, value):
        self.com_object.Accumulate = value

    # Lower case aliases for Accumulate
    @property
    def accumulate(self):
        return self.Accumulate

    @accumulate.setter
    def accumulate(self, value):
        self.Accumulate = value

    @property
    def Additive(self):
        return self.com_object.Additive

    @Additive.setter
    def Additive(self, value):
        self.com_object.Additive = value

    # Lower case aliases for Additive
    @property
    def additive(self):
        return self.Additive

    @additive.setter
    def additive(self, value):
        self.Additive = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def ColorEffect(self):
        return ColorEffect(self.com_object.ColorEffect)

    # Lower case aliases for ColorEffect
    @property
    def coloreffect(self):
        return self.ColorEffect

    @property
    def CommandEffect(self):
        return CommandEffect(self.com_object.CommandEffect)

    # Lower case aliases for CommandEffect
    @property
    def commandeffect(self):
        return self.CommandEffect

    @property
    def FilterEffect(self):
        return FilterEffect(self.com_object.FilterEffect)

    # Lower case aliases for FilterEffect
    @property
    def filtereffect(self):
        return self.FilterEffect

    @property
    def MotionEffect(self):
        return MotionEffect(self.com_object.MotionEffect)

    # Lower case aliases for MotionEffect
    @property
    def motioneffect(self):
        return self.MotionEffect

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyEffect(self):
        return PropertyEffect(self.com_object.PropertyEffect)

    # Lower case aliases for PropertyEffect
    @property
    def propertyeffect(self):
        return self.PropertyEffect

    @property
    def RotationEffect(self):
        return RotationEffect(self.com_object.RotationEffect)

    # Lower case aliases for RotationEffect
    @property
    def rotationeffect(self):
        return self.RotationEffect

    @property
    def ScaleEffect(self):
        return ScaleEffect(self.com_object.ScaleEffect)

    # Lower case aliases for ScaleEffect
    @property
    def scaleeffect(self):
        return self.ScaleEffect

    @property
    def SetEffect(self):
        return SetEffect(self.com_object.SetEffect)

    # Lower case aliases for SetEffect
    @property
    def seteffect(self):
        return self.SetEffect

    @property
    def Timing(self):
        return Timing(self.com_object.Timing)

    # Lower case aliases for Timing
    @property
    def timing(self):
        return self.Timing

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


class AnimationBehaviors:

    def __init__(self, animationbehaviors=None):
        self.com_object= animationbehaviors

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Add(self, Type=None, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Type, Index]])
        return AnimationBehavior(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Type=None, Index=None):
        arguments = [Type, Index]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return AnimationBehavior(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class AnimationPoint:

    def __init__(self, animationpoint=None):
        self.com_object= animationpoint

    @property
    def Application(self):
        return Application(self.com_object.Application)

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
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

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

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()


class AnimationPoints:

    def __init__(self, animationpoints=None):
        self.com_object= animationpoints

    @property
    def Application(self):
        return Application(self.com_object.Application)

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
    def Smooth(self):
        return self.com_object.Smooth

    @Smooth.setter
    def Smooth(self, value):
        self.com_object.Smooth = value

    # Lower case aliases for Smooth
    @property
    def smooth(self):
        return self.Smooth

    @smooth.setter
    def smooth(self, value):
        self.Smooth = value

    def Add(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return AnimationPoint(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Index=None):
        arguments = [Index]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return AnimationPoint(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class AnimationSettings:

    def __init__(self, animationsettings=None):
        self.com_object= animationsettings

    @property
    def AdvanceMode(self):
        return self.com_object.AdvanceMode

    @AdvanceMode.setter
    def AdvanceMode(self, value):
        self.com_object.AdvanceMode = value

    # Lower case aliases for AdvanceMode
    @property
    def advancemode(self):
        return self.AdvanceMode

    @advancemode.setter
    def advancemode(self, value):
        self.AdvanceMode = value

    @property
    def AdvanceTime(self):
        return self.com_object.AdvanceTime

    @AdvanceTime.setter
    def AdvanceTime(self, value):
        self.com_object.AdvanceTime = value

    # Lower case aliases for AdvanceTime
    @property
    def advancetime(self):
        return self.AdvanceTime

    @advancetime.setter
    def advancetime(self, value):
        self.AdvanceTime = value

    @property
    def AfterEffect(self):
        return PpAfterEffect(self.com_object.AfterEffect)

    @AfterEffect.setter
    def AfterEffect(self, value):
        self.com_object.AfterEffect = value

    # Lower case aliases for AfterEffect
    @property
    def aftereffect(self):
        return self.AfterEffect

    @aftereffect.setter
    def aftereffect(self, value):
        self.AfterEffect = value

    @property
    def Animate(self):
        return self.com_object.Animate

    @Animate.setter
    def Animate(self, value):
        self.com_object.Animate = value

    # Lower case aliases for Animate
    @property
    def animate(self):
        return self.Animate

    @animate.setter
    def animate(self, value):
        self.Animate = value

    @property
    def AnimateBackground(self):
        return self.com_object.AnimateBackground

    @AnimateBackground.setter
    def AnimateBackground(self, value):
        self.com_object.AnimateBackground = value

    # Lower case aliases for AnimateBackground
    @property
    def animatebackground(self):
        return self.AnimateBackground

    @animatebackground.setter
    def animatebackground(self, value):
        self.AnimateBackground = value

    @property
    def AnimateTextInReverse(self):
        return self.com_object.AnimateTextInReverse

    @AnimateTextInReverse.setter
    def AnimateTextInReverse(self, value):
        self.com_object.AnimateTextInReverse = value

    # Lower case aliases for AnimateTextInReverse
    @property
    def animatetextinreverse(self):
        return self.AnimateTextInReverse

    @animatetextinreverse.setter
    def animatetextinreverse(self, value):
        self.AnimateTextInReverse = value

    @property
    def AnimationOrder(self):
        return self.com_object.AnimationOrder

    @AnimationOrder.setter
    def AnimationOrder(self, value):
        self.com_object.AnimationOrder = value

    # Lower case aliases for AnimationOrder
    @property
    def animationorder(self):
        return self.AnimationOrder

    @animationorder.setter
    def animationorder(self, value):
        self.AnimationOrder = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def ChartUnitEffect(self):
        return self.com_object.ChartUnitEffect

    @ChartUnitEffect.setter
    def ChartUnitEffect(self, value):
        self.com_object.ChartUnitEffect = value

    # Lower case aliases for ChartUnitEffect
    @property
    def chartuniteffect(self):
        return self.ChartUnitEffect

    @chartuniteffect.setter
    def chartuniteffect(self, value):
        self.ChartUnitEffect = value

    @property
    def DimColor(self):
        return ColorFormat(self.com_object.DimColor)

    @DimColor.setter
    def DimColor(self, value):
        self.com_object.DimColor = value

    # Lower case aliases for DimColor
    @property
    def dimcolor(self):
        return self.DimColor

    @dimcolor.setter
    def dimcolor(self, value):
        self.DimColor = value

    @property
    def EntryEffect(self):
        return self.com_object.EntryEffect

    @EntryEffect.setter
    def EntryEffect(self, value):
        self.com_object.EntryEffect = value

    # Lower case aliases for EntryEffect
    @property
    def entryeffect(self):
        return self.EntryEffect

    @entryeffect.setter
    def entryeffect(self, value):
        self.EntryEffect = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PlaySettings(self):
        return PlaySettings(self.com_object.PlaySettings)

    # Lower case aliases for PlaySettings
    @property
    def playsettings(self):
        return self.PlaySettings

    @property
    def SoundEffect(self):
        return SoundEffect(self.com_object.SoundEffect)

    # Lower case aliases for SoundEffect
    @property
    def soundeffect(self):
        return self.SoundEffect

    @property
    def TextLevelEffect(self):
        return self.com_object.TextLevelEffect

    @TextLevelEffect.setter
    def TextLevelEffect(self, value):
        self.com_object.TextLevelEffect = value

    # Lower case aliases for TextLevelEffect
    @property
    def textleveleffect(self):
        return self.TextLevelEffect

    @textleveleffect.setter
    def textleveleffect(self, value):
        self.TextLevelEffect = value

    @property
    def TextUnitEffect(self):
        return self.com_object.TextUnitEffect

    @TextUnitEffect.setter
    def TextUnitEffect(self, value):
        self.com_object.TextUnitEffect = value

    # Lower case aliases for TextUnitEffect
    @property
    def textuniteffect(self):
        return self.TextUnitEffect

    @textuniteffect.setter
    def textuniteffect(self, value):
        self.TextUnitEffect = value


class Application:

    def __init__(self, application=None):
        self.com_object= application

    def new(self):
        self.com_object = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
        return self

    @property
    def Active(self):
        return self.com_object.Active

    # Lower case aliases for Active
    @property
    def active(self):
        return self.Active

    @property
    def ActiveEncryptionSession(self):
        return self.com_object.ActiveEncryptionSession

    # Lower case aliases for ActiveEncryptionSession
    @property
    def activeencryptionsession(self):
        return self.ActiveEncryptionSession

    @property
    def ActivePresentation(self):
        return Presentation(self.com_object.ActivePresentation)

    # Lower case aliases for ActivePresentation
    @property
    def activepresentation(self):
        return self.ActivePresentation

    @property
    def ActivePrinter(self):
        return self.com_object.ActivePrinter

    # Lower case aliases for ActivePrinter
    @property
    def activeprinter(self):
        return self.ActivePrinter

    @property
    def ActiveProtectedViewWindow(self):
        return ProtectedViewWindow(self.com_object.ActiveProtectedViewWindow)

    # Lower case aliases for ActiveProtectedViewWindow
    @property
    def activeprotectedviewwindow(self):
        return self.ActiveProtectedViewWindow

    @property
    def ActiveWindow(self):
        return DocumentWindow(self.com_object.ActiveWindow)

    # Lower case aliases for ActiveWindow
    @property
    def activewindow(self):
        return self.ActiveWindow

    @property
    def AddIns(self):
        return AddIns(self.com_object.AddIns)

    # Lower case aliases for AddIns
    @property
    def addins(self):
        return self.AddIns

    @property
    def Assistance(self):
        return self.com_object.Assistance

    # Lower case aliases for Assistance
    @property
    def assistance(self):
        return self.Assistance

    @property
    def AutoCorrect(self):
        return AutoCorrect(self.com_object.AutoCorrect)

    # Lower case aliases for AutoCorrect
    @property
    def autocorrect(self):
        return self.AutoCorrect

    @property
    def AutomationSecurity(self):
        return self.com_object.AutomationSecurity

    @AutomationSecurity.setter
    def AutomationSecurity(self, value):
        self.com_object.AutomationSecurity = value

    # Lower case aliases for AutomationSecurity
    @property
    def automationsecurity(self):
        return self.AutomationSecurity

    @automationsecurity.setter
    def automationsecurity(self, value):
        self.AutomationSecurity = value

    @property
    def Build(self):
        return self.com_object.Build

    # Lower case aliases for Build
    @property
    def build(self):
        return self.Build

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
    def COMAddIns(self):
        return self.com_object.COMAddIns

    # Lower case aliases for COMAddIns
    @property
    def comaddins(self):
        return self.COMAddIns

    @property
    def CommandBars(self):
        return self.com_object.CommandBars

    # Lower case aliases for CommandBars
    @property
    def commandbars(self):
        return self.CommandBars

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def DisplayAlerts(self):
        return self.com_object.DisplayAlerts

    @DisplayAlerts.setter
    def DisplayAlerts(self, value):
        self.com_object.DisplayAlerts = value

    # Lower case aliases for DisplayAlerts
    @property
    def displayalerts(self):
        return self.DisplayAlerts

    @displayalerts.setter
    def displayalerts(self, value):
        self.DisplayAlerts = value

    @property
    def DisplayDocumentInformationPanel(self):
        return self.com_object.DisplayDocumentInformationPanel

    @DisplayDocumentInformationPanel.setter
    def DisplayDocumentInformationPanel(self, value):
        self.com_object.DisplayDocumentInformationPanel = value

    # Lower case aliases for DisplayDocumentInformationPanel
    @property
    def displaydocumentinformationpanel(self):
        return self.DisplayDocumentInformationPanel

    @displaydocumentinformationpanel.setter
    def displaydocumentinformationpanel(self, value):
        self.DisplayDocumentInformationPanel = value

    @property
    def DisplayGridLines(self):
        return self.com_object.DisplayGridLines

    @DisplayGridLines.setter
    def DisplayGridLines(self, value):
        self.com_object.DisplayGridLines = value

    # Lower case aliases for DisplayGridLines
    @property
    def displaygridlines(self):
        return self.DisplayGridLines

    @displaygridlines.setter
    def displaygridlines(self, value):
        self.DisplayGridLines = value

    @property
    def FeatureInstall(self):
        return self.com_object.FeatureInstall

    @FeatureInstall.setter
    def FeatureInstall(self, value):
        self.com_object.FeatureInstall = value

    # Lower case aliases for FeatureInstall
    @property
    def featureinstall(self):
        return self.FeatureInstall

    @featureinstall.setter
    def featureinstall(self, value):
        self.FeatureInstall = value

    def FileConverters(self, Index1=None, Index2=None):
        arguments = com_arguments([unwrap(a) for a in [Index1, Index2]])
        if hasattr(self.com_object, "GetFileConverters"):
            return self.com_object.GetFileConverters(*arguments)
        else:
            return self.com_object.FileConverters(*arguments)

    # Lower case aliases for FileConverters
    def fileconverters(self, Index1=None, Index2=None):
        arguments = [Index1, Index2]
        return self.FileConverters(*arguments)

    def FileDialog(self, Type=None):
        arguments = com_arguments([unwrap(a) for a in [Type]])
        if hasattr(self.com_object, "GetFileDialog"):
            return self.com_object.GetFileDialog(*arguments)
        else:
            return self.com_object.FileDialog(*arguments)

    # Lower case aliases for FileDialog
    def filedialog(self, Type=None):
        arguments = [Type]
        return self.FileDialog(*arguments)

    @property
    def FileValidation(self):
        return self.com_object.FileValidation

    @FileValidation.setter
    def FileValidation(self, value):
        self.com_object.FileValidation = value

    # Lower case aliases for FileValidation
    @property
    def filevalidation(self):
        return self.FileValidation

    @filevalidation.setter
    def filevalidation(self, value):
        self.FileValidation = value

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
    def IsSandboxed(self):
        return self.com_object.IsSandboxed

    # Lower case aliases for IsSandboxed
    @property
    def issandboxed(self):
        return self.IsSandboxed

    @property
    def LanguageSettings(self):
        return self.com_object.LanguageSettings

    # Lower case aliases for LanguageSettings
    @property
    def languagesettings(self):
        return self.LanguageSettings

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
    def Name(self):
        return self.com_object.Name

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def NewPresentation(self):
        return self.com_object.NewPresentation

    # Lower case aliases for NewPresentation
    @property
    def newpresentation(self):
        return self.NewPresentation

    @property
    def OperatingSystem(self):
        return self.com_object.OperatingSystem

    # Lower case aliases for OperatingSystem
    @property
    def operatingsystem(self):
        return self.OperatingSystem

    @property
    def Options(self):
        return Options(self.com_object.Options)

    # Lower case aliases for Options
    @property
    def options(self):
        return self.Options

    @property
    def Path(self):
        return Application(self.com_object.Path)

    # Lower case aliases for Path
    @property
    def path(self):
        return self.Path

    @property
    def Presentations(self):
        return Presentations(self.com_object.Presentations)

    # Lower case aliases for Presentations
    @property
    def presentations(self):
        return self.Presentations

    @property
    def ProductCode(self):
        return self.com_object.ProductCode

    # Lower case aliases for ProductCode
    @property
    def productcode(self):
        return self.ProductCode

    @property
    def ProtectedViewWindows(self):
        return ProtectedViewWindows(self.com_object.ProtectedViewWindows)

    # Lower case aliases for ProtectedViewWindows
    @property
    def protectedviewwindows(self):
        return self.ProtectedViewWindows

    @property
    def SensitivityLabelPolicy(self):
        return self.com_object.SensitivityLabelPolicy

    # Lower case aliases for SensitivityLabelPolicy
    @property
    def sensitivitylabelpolicy(self):
        return self.SensitivityLabelPolicy

    @property
    def ShowStartupDialog(self):
        return self.com_object.ShowStartupDialog

    @ShowStartupDialog.setter
    def ShowStartupDialog(self, value):
        self.com_object.ShowStartupDialog = value

    # Lower case aliases for ShowStartupDialog
    @property
    def showstartupdialog(self):
        return self.ShowStartupDialog

    @showstartupdialog.setter
    def showstartupdialog(self, value):
        self.ShowStartupDialog = value

    @property
    def ShowWindowsInTaskbar(self):
        return self.com_object.ShowWindowsInTaskbar

    @ShowWindowsInTaskbar.setter
    def ShowWindowsInTaskbar(self, value):
        self.com_object.ShowWindowsInTaskbar = value

    # Lower case aliases for ShowWindowsInTaskbar
    @property
    def showwindowsintaskbar(self):
        return self.ShowWindowsInTaskbar

    @showwindowsintaskbar.setter
    def showwindowsintaskbar(self, value):
        self.ShowWindowsInTaskbar = value

    @property
    def SlideShowWindows(self):
        return SlideShowWindows(self.com_object.SlideShowWindows)

    # Lower case aliases for SlideShowWindows
    @property
    def slideshowwindows(self):
        return self.SlideShowWindows

    @property
    def SmartArtColors(self):
        return Application(self.com_object.SmartArtColors)

    # Lower case aliases for SmartArtColors
    @property
    def smartartcolors(self):
        return self.SmartArtColors

    @property
    def SmartArtLayouts(self):
        return Application(self.com_object.SmartArtLayouts)

    # Lower case aliases for SmartArtLayouts
    @property
    def smartartlayouts(self):
        return self.SmartArtLayouts

    @property
    def SmartArtQuickStyles(self):
        return Application(self.com_object.SmartArtQuickStyles)

    # Lower case aliases for SmartArtQuickStyles
    @property
    def smartartquickstyles(self):
        return self.SmartArtQuickStyles

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
    def VBE(self):
        return self.com_object.VBE

    # Lower case aliases for VBE
    @property
    def vbe(self):
        return self.VBE

    @property
    def Version(self):
        return self.com_object.Version

    # Lower case aliases for Version
    @property
    def version(self):
        return self.Version

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
    def Windows(self):
        return DocumentWindows(self.com_object.Windows)

    # Lower case aliases for Windows
    @property
    def windows(self):
        return self.Windows

    @property
    def WindowState(self):
        return self.com_object.WindowState

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

    def Help(self, HelpFile=None, ContextID=None):
        arguments = com_arguments([unwrap(a) for a in [HelpFile, ContextID]])
        self.com_object.Help(*arguments)

    # Lower case alias for Help
    def help(self, HelpFile=None, ContextID=None):
        arguments = [HelpFile, ContextID]
        return self.Help(*arguments)

    def Quit(self):
        self.com_object.Quit()

    # Lower case alias for Quit
    def quit(self):
        return self.Quit()

    def Run(self, MacroName=None, safeArrayOfParams=None):
        arguments = com_arguments([unwrap(a) for a in [MacroName, safeArrayOfParams]])
        return self.com_object.Run(*arguments)

    # Lower case alias for Run
    def run(self, MacroName=None, safeArrayOfParams=None):
        arguments = [MacroName, safeArrayOfParams]
        return self.Run(*arguments)

    def StartNewUndoEntry(self):
        self.com_object.StartNewUndoEntry()

    # Lower case alias for StartNewUndoEntry
    def startnewundoentry(self):
        return self.StartNewUndoEntry()


class AutoCorrect:

    def __init__(self, autocorrect=None):
        self.com_object= autocorrect

    @property
    def DisplayAutoCorrectOptions(self):
        return self.com_object.DisplayAutoCorrectOptions

    @DisplayAutoCorrectOptions.setter
    def DisplayAutoCorrectOptions(self, value):
        self.com_object.DisplayAutoCorrectOptions = value

    # Lower case aliases for DisplayAutoCorrectOptions
    @property
    def displayautocorrectoptions(self):
        return self.DisplayAutoCorrectOptions

    @displayautocorrectoptions.setter
    def displayautocorrectoptions(self, value):
        self.DisplayAutoCorrectOptions = value

    @property
    def DisplayAutoLayoutOptions(self):
        return self.com_object.DisplayAutoLayoutOptions

    @DisplayAutoLayoutOptions.setter
    def DisplayAutoLayoutOptions(self, value):
        self.com_object.DisplayAutoLayoutOptions = value

    # Lower case aliases for DisplayAutoLayoutOptions
    @property
    def displayautolayoutoptions(self):
        return self.DisplayAutoLayoutOptions

    @displayautolayoutoptions.setter
    def displayautolayoutoptions(self, value):
        self.DisplayAutoLayoutOptions = value


class Axes:

    def __init__(self, axes=None):
        self.com_object= axes

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    def Item(self, Type=None, AxisGroup=None):
        arguments = com_arguments([unwrap(a) for a in [Type, AxisGroup]])
        self.com_object.Item(*arguments)

    # Lower case alias for Item
    def item(self, Type=None, AxisGroup=None):
        arguments = [Type, AxisGroup]
        return self.Item(*arguments)


class Axis:

    def __init__(self, axis=None):
        self.com_object= axis

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def AxisBetweenCategories(self):
        return self.com_object.AxisBetweenCategories

    @AxisBetweenCategories.setter
    def AxisBetweenCategories(self, value):
        self.com_object.AxisBetweenCategories = value

    # Lower case aliases for AxisBetweenCategories
    @property
    def axisbetweencategories(self):
        return self.AxisBetweenCategories

    @axisbetweencategories.setter
    def axisbetweencategories(self, value):
        self.AxisBetweenCategories = value

    @property
    def AxisGroup(self):
        return XlAxisGroup(self.com_object.AxisGroup)

    # Lower case aliases for AxisGroup
    @property
    def axisgroup(self):
        return self.AxisGroup

    @property
    def AxisTitle(self):
        return AxisTitle(self.com_object.AxisTitle)

    # Lower case aliases for AxisTitle
    @property
    def axistitle(self):
        return self.AxisTitle

    @property
    def BaseUnit(self):
        return XlTimeUnit(self.com_object.BaseUnit)

    @BaseUnit.setter
    def BaseUnit(self, value):
        self.com_object.BaseUnit = value

    # Lower case aliases for BaseUnit
    @property
    def baseunit(self):
        return self.BaseUnit

    @baseunit.setter
    def baseunit(self, value):
        self.BaseUnit = value

    @property
    def BaseUnitIsAuto(self):
        return self.com_object.BaseUnitIsAuto

    @BaseUnitIsAuto.setter
    def BaseUnitIsAuto(self, value):
        self.com_object.BaseUnitIsAuto = value

    # Lower case aliases for BaseUnitIsAuto
    @property
    def baseunitisauto(self):
        return self.BaseUnitIsAuto

    @baseunitisauto.setter
    def baseunitisauto(self, value):
        self.BaseUnitIsAuto = value

    @property
    def Border(self):
        return ChartBorder(self.com_object.Border)

    # Lower case aliases for Border
    @property
    def border(self):
        return self.Border

    @property
    def CategoryNames(self):
        return self.com_object.CategoryNames

    @CategoryNames.setter
    def CategoryNames(self, value):
        self.com_object.CategoryNames = value

    # Lower case aliases for CategoryNames
    @property
    def categorynames(self):
        return self.CategoryNames

    @categorynames.setter
    def categorynames(self, value):
        self.CategoryNames = value

    @property
    def CategoryType(self):
        return XlCategoryType(self.com_object.CategoryType)

    @CategoryType.setter
    def CategoryType(self, value):
        self.com_object.CategoryType = value

    # Lower case aliases for CategoryType
    @property
    def categorytype(self):
        return self.CategoryType

    @categorytype.setter
    def categorytype(self, value):
        self.CategoryType = value

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Crosses(self):
        return self.com_object.Crosses

    @Crosses.setter
    def Crosses(self, value):
        self.com_object.Crosses = value

    # Lower case aliases for Crosses
    @property
    def crosses(self):
        return self.Crosses

    @crosses.setter
    def crosses(self, value):
        self.Crosses = value

    @property
    def CrossesAt(self):
        return self.com_object.CrossesAt

    @CrossesAt.setter
    def CrossesAt(self, value):
        self.com_object.CrossesAt = value

    # Lower case aliases for CrossesAt
    @property
    def crossesat(self):
        return self.CrossesAt

    @crossesat.setter
    def crossesat(self, value):
        self.CrossesAt = value

    @property
    def DisplayUnit(self):
        return XlDisplayUnit(self.com_object.DisplayUnit)

    @DisplayUnit.setter
    def DisplayUnit(self, value):
        self.com_object.DisplayUnit = value

    # Lower case aliases for DisplayUnit
    @property
    def displayunit(self):
        return self.DisplayUnit

    @displayunit.setter
    def displayunit(self, value):
        self.DisplayUnit = value

    @property
    def DisplayUnitCustom(self):
        return self.com_object.DisplayUnitCustom

    @DisplayUnitCustom.setter
    def DisplayUnitCustom(self, value):
        self.com_object.DisplayUnitCustom = value

    # Lower case aliases for DisplayUnitCustom
    @property
    def displayunitcustom(self):
        return self.DisplayUnitCustom

    @displayunitcustom.setter
    def displayunitcustom(self, value):
        self.DisplayUnitCustom = value

    @property
    def DisplayUnitLabel(self):
        return DisplayUnitLabel(self.com_object.DisplayUnitLabel)

    # Lower case aliases for DisplayUnitLabel
    @property
    def displayunitlabel(self):
        return self.DisplayUnitLabel

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

    @property
    def HasDisplayUnitLabel(self):
        return self.com_object.HasDisplayUnitLabel

    @HasDisplayUnitLabel.setter
    def HasDisplayUnitLabel(self, value):
        self.com_object.HasDisplayUnitLabel = value

    # Lower case aliases for HasDisplayUnitLabel
    @property
    def hasdisplayunitlabel(self):
        return self.HasDisplayUnitLabel

    @hasdisplayunitlabel.setter
    def hasdisplayunitlabel(self, value):
        self.HasDisplayUnitLabel = value

    @property
    def HasMajorGridlines(self):
        return self.com_object.HasMajorGridlines

    @HasMajorGridlines.setter
    def HasMajorGridlines(self, value):
        self.com_object.HasMajorGridlines = value

    # Lower case aliases for HasMajorGridlines
    @property
    def hasmajorgridlines(self):
        return self.HasMajorGridlines

    @hasmajorgridlines.setter
    def hasmajorgridlines(self, value):
        self.HasMajorGridlines = value

    @property
    def HasMinorGridlines(self):
        return self.com_object.HasMinorGridlines

    @HasMinorGridlines.setter
    def HasMinorGridlines(self, value):
        self.com_object.HasMinorGridlines = value

    # Lower case aliases for HasMinorGridlines
    @property
    def hasminorgridlines(self):
        return self.HasMinorGridlines

    @hasminorgridlines.setter
    def hasminorgridlines(self, value):
        self.HasMinorGridlines = value

    @property
    def HasTitle(self):
        return self.com_object.HasTitle

    @HasTitle.setter
    def HasTitle(self, value):
        self.com_object.HasTitle = value

    # Lower case aliases for HasTitle
    @property
    def hastitle(self):
        return self.HasTitle

    @hastitle.setter
    def hastitle(self, value):
        self.HasTitle = value

    @property
    def Height(self):
        return self.com_object.Height

    # Lower case aliases for Height
    @property
    def height(self):
        return self.Height

    @property
    def Left(self):
        return self.com_object.Left

    # Lower case aliases for Left
    @property
    def left(self):
        return self.Left

    @property
    def LogBase(self):
        return self.com_object.LogBase

    @LogBase.setter
    def LogBase(self, value):
        self.com_object.LogBase = value

    # Lower case aliases for LogBase
    @property
    def logbase(self):
        return self.LogBase

    @logbase.setter
    def logbase(self, value):
        self.LogBase = value

    @property
    def MajorGridlines(self):
        return Gridlines(self.com_object.MajorGridlines)

    # Lower case aliases for MajorGridlines
    @property
    def majorgridlines(self):
        return self.MajorGridlines

    @property
    def MajorTickMark(self):
        return XlTickMark(self.com_object.MajorTickMark)

    @MajorTickMark.setter
    def MajorTickMark(self, value):
        self.com_object.MajorTickMark = value

    # Lower case aliases for MajorTickMark
    @property
    def majortickmark(self):
        return self.MajorTickMark

    @majortickmark.setter
    def majortickmark(self, value):
        self.MajorTickMark = value

    @property
    def MajorUnit(self):
        return self.com_object.MajorUnit

    @MajorUnit.setter
    def MajorUnit(self, value):
        self.com_object.MajorUnit = value

    # Lower case aliases for MajorUnit
    @property
    def majorunit(self):
        return self.MajorUnit

    @majorunit.setter
    def majorunit(self, value):
        self.MajorUnit = value

    @property
    def MajorUnitIsAuto(self):
        return self.com_object.MajorUnitIsAuto

    @MajorUnitIsAuto.setter
    def MajorUnitIsAuto(self, value):
        self.com_object.MajorUnitIsAuto = value

    # Lower case aliases for MajorUnitIsAuto
    @property
    def majorunitisauto(self):
        return self.MajorUnitIsAuto

    @majorunitisauto.setter
    def majorunitisauto(self, value):
        self.MajorUnitIsAuto = value

    @property
    def MajorUnitScale(self):
        return self.com_object.MajorUnitScale

    @MajorUnitScale.setter
    def MajorUnitScale(self, value):
        self.com_object.MajorUnitScale = value

    # Lower case aliases for MajorUnitScale
    @property
    def majorunitscale(self):
        return self.MajorUnitScale

    @majorunitscale.setter
    def majorunitscale(self, value):
        self.MajorUnitScale = value

    @property
    def MaximumScale(self):
        return self.com_object.MaximumScale

    @MaximumScale.setter
    def MaximumScale(self, value):
        self.com_object.MaximumScale = value

    # Lower case aliases for MaximumScale
    @property
    def maximumscale(self):
        return self.MaximumScale

    @maximumscale.setter
    def maximumscale(self, value):
        self.MaximumScale = value

    @property
    def MaximumScaleIsAuto(self):
        return self.com_object.MaximumScaleIsAuto

    @MaximumScaleIsAuto.setter
    def MaximumScaleIsAuto(self, value):
        self.com_object.MaximumScaleIsAuto = value

    # Lower case aliases for MaximumScaleIsAuto
    @property
    def maximumscaleisauto(self):
        return self.MaximumScaleIsAuto

    @maximumscaleisauto.setter
    def maximumscaleisauto(self, value):
        self.MaximumScaleIsAuto = value

    @property
    def MinimumScale(self):
        return self.com_object.MinimumScale

    @MinimumScale.setter
    def MinimumScale(self, value):
        self.com_object.MinimumScale = value

    # Lower case aliases for MinimumScale
    @property
    def minimumscale(self):
        return self.MinimumScale

    @minimumscale.setter
    def minimumscale(self, value):
        self.MinimumScale = value

    @property
    def MinimumScaleIsAuto(self):
        return self.com_object.MinimumScaleIsAuto

    @MinimumScaleIsAuto.setter
    def MinimumScaleIsAuto(self, value):
        self.com_object.MinimumScaleIsAuto = value

    # Lower case aliases for MinimumScaleIsAuto
    @property
    def minimumscaleisauto(self):
        return self.MinimumScaleIsAuto

    @minimumscaleisauto.setter
    def minimumscaleisauto(self, value):
        self.MinimumScaleIsAuto = value

    @property
    def MinorGridlines(self):
        return Gridlines(self.com_object.MinorGridlines)

    # Lower case aliases for MinorGridlines
    @property
    def minorgridlines(self):
        return self.MinorGridlines

    @property
    def MinorTickMark(self):
        return XlTickMark(self.com_object.MinorTickMark)

    @MinorTickMark.setter
    def MinorTickMark(self, value):
        self.com_object.MinorTickMark = value

    # Lower case aliases for MinorTickMark
    @property
    def minortickmark(self):
        return self.MinorTickMark

    @minortickmark.setter
    def minortickmark(self, value):
        self.MinorTickMark = value

    @property
    def MinorUnit(self):
        return self.com_object.MinorUnit

    @MinorUnit.setter
    def MinorUnit(self, value):
        self.com_object.MinorUnit = value

    # Lower case aliases for MinorUnit
    @property
    def minorunit(self):
        return self.MinorUnit

    @minorunit.setter
    def minorunit(self, value):
        self.MinorUnit = value

    @property
    def MinorUnitIsAuto(self):
        return self.com_object.MinorUnitIsAuto

    @MinorUnitIsAuto.setter
    def MinorUnitIsAuto(self, value):
        self.com_object.MinorUnitIsAuto = value

    # Lower case aliases for MinorUnitIsAuto
    @property
    def minorunitisauto(self):
        return self.MinorUnitIsAuto

    @minorunitisauto.setter
    def minorunitisauto(self, value):
        self.MinorUnitIsAuto = value

    @property
    def MinorUnitScale(self):
        return self.com_object.MinorUnitScale

    @MinorUnitScale.setter
    def MinorUnitScale(self, value):
        self.com_object.MinorUnitScale = value

    # Lower case aliases for MinorUnitScale
    @property
    def minorunitscale(self):
        return self.MinorUnitScale

    @minorunitscale.setter
    def minorunitscale(self, value):
        self.MinorUnitScale = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ReversePlotOrder(self):
        return self.com_object.ReversePlotOrder

    @ReversePlotOrder.setter
    def ReversePlotOrder(self, value):
        self.com_object.ReversePlotOrder = value

    # Lower case aliases for ReversePlotOrder
    @property
    def reverseplotorder(self):
        return self.ReversePlotOrder

    @reverseplotorder.setter
    def reverseplotorder(self, value):
        self.ReversePlotOrder = value

    @property
    def ScaleType(self):
        return XlScaleType(self.com_object.ScaleType)

    @ScaleType.setter
    def ScaleType(self, value):
        self.com_object.ScaleType = value

    # Lower case aliases for ScaleType
    @property
    def scaletype(self):
        return self.ScaleType

    @scaletype.setter
    def scaletype(self, value):
        self.ScaleType = value

    @property
    def TickLabelPosition(self):
        return self.com_object.TickLabelPosition

    @TickLabelPosition.setter
    def TickLabelPosition(self, value):
        self.com_object.TickLabelPosition = value

    # Lower case aliases for TickLabelPosition
    @property
    def ticklabelposition(self):
        return self.TickLabelPosition

    @ticklabelposition.setter
    def ticklabelposition(self, value):
        self.TickLabelPosition = value

    @property
    def TickLabels(self):
        return TickLabels(self.com_object.TickLabels)

    # Lower case aliases for TickLabels
    @property
    def ticklabels(self):
        return self.TickLabels

    @property
    def TickLabelSpacing(self):
        return self.com_object.TickLabelSpacing

    @TickLabelSpacing.setter
    def TickLabelSpacing(self, value):
        self.com_object.TickLabelSpacing = value

    # Lower case aliases for TickLabelSpacing
    @property
    def ticklabelspacing(self):
        return self.TickLabelSpacing

    @ticklabelspacing.setter
    def ticklabelspacing(self, value):
        self.TickLabelSpacing = value

    @property
    def TickLabelSpacingIsAuto(self):
        return self.com_object.TickLabelSpacingIsAuto

    @TickLabelSpacingIsAuto.setter
    def TickLabelSpacingIsAuto(self, value):
        self.com_object.TickLabelSpacingIsAuto = value

    # Lower case aliases for TickLabelSpacingIsAuto
    @property
    def ticklabelspacingisauto(self):
        return self.TickLabelSpacingIsAuto

    @ticklabelspacingisauto.setter
    def ticklabelspacingisauto(self, value):
        self.TickLabelSpacingIsAuto = value

    @property
    def TickMarkSpacing(self):
        return self.com_object.TickMarkSpacing

    @TickMarkSpacing.setter
    def TickMarkSpacing(self, value):
        self.com_object.TickMarkSpacing = value

    # Lower case aliases for TickMarkSpacing
    @property
    def tickmarkspacing(self):
        return self.TickMarkSpacing

    @tickmarkspacing.setter
    def tickmarkspacing(self, value):
        self.TickMarkSpacing = value

    @property
    def Top(self):
        return self.com_object.Top

    # Lower case aliases for Top
    @property
    def top(self):
        return self.Top

    @property
    def Type(self):
        return XlAxisType(self.com_object.Type)

    # Lower case aliases for Type
    @property
    def type(self):
        return self.Type

    @property
    def Width(self):
        return self.com_object.Width

    # Lower case aliases for Width
    @property
    def width(self):
        return self.Width

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class AxisTitle:

    def __init__(self, axistitle=None):
        self.com_object= axistitle

    @property
    def Application(self):
        return self.com_object.Application

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

    def Characters(self, Start=None, Length=None):
        arguments = com_arguments([unwrap(a) for a in [Start, Length]])
        if hasattr(self.com_object, "GetCharacters"):
            return ChartCharacters(self.com_object.GetCharacters(*arguments))
        else:
            return ChartCharacters(self.com_object.Characters(*arguments))

    # Lower case aliases for Characters
    def characters(self, Start=None, Length=None):
        arguments = [Start, Length]
        return self.Characters(*arguments)

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

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
    def FormulaLocal(self):
        return self.com_object.FormulaLocal

    @FormulaLocal.setter
    def FormulaLocal(self, value):
        self.com_object.FormulaLocal = value

    # Lower case aliases for FormulaLocal
    @property
    def formulalocal(self):
        return self.FormulaLocal

    @formulalocal.setter
    def formulalocal(self, value):
        self.FormulaLocal = value

    @property
    def FormulaR1C1(self):
        return self.com_object.FormulaR1C1

    @FormulaR1C1.setter
    def FormulaR1C1(self, value):
        self.com_object.FormulaR1C1 = value

    # Lower case aliases for FormulaR1C1
    @property
    def formular1c1(self):
        return self.FormulaR1C1

    @formular1c1.setter
    def formular1c1(self, value):
        self.FormulaR1C1 = value

    @property
    def FormulaR1C1Local(self):
        return self.com_object.FormulaR1C1Local

    @FormulaR1C1Local.setter
    def FormulaR1C1Local(self, value):
        self.com_object.FormulaR1C1Local = value

    # Lower case aliases for FormulaR1C1Local
    @property
    def formular1c1local(self):
        return self.FormulaR1C1Local

    @formular1c1local.setter
    def formular1c1local(self, value):
        self.FormulaR1C1Local = value

    @property
    def Height(self):
        return self.com_object.Height

    # Lower case aliases for Height
    @property
    def height(self):
        return self.Height

    @property
    def HorizontalAlignment(self):
        return self.com_object.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.com_object.HorizontalAlignment = value

    # Lower case aliases for HorizontalAlignment
    @property
    def horizontalalignment(self):
        return self.HorizontalAlignment

    @horizontalalignment.setter
    def horizontalalignment(self, value):
        self.HorizontalAlignment = value

    @property
    def IncludeInLayout(self):
        return self.com_object.IncludeInLayout

    @IncludeInLayout.setter
    def IncludeInLayout(self, value):
        self.com_object.IncludeInLayout = value

    # Lower case aliases for IncludeInLayout
    @property
    def includeinlayout(self):
        return self.IncludeInLayout

    @includeinlayout.setter
    def includeinlayout(self, value):
        self.IncludeInLayout = value

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
    def Name(self):
        return self.com_object.Name

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def Orientation(self):
        return self.com_object.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.com_object.Orientation = value

    # Lower case aliases for Orientation
    @property
    def orientation(self):
        return self.Orientation

    @orientation.setter
    def orientation(self, value):
        self.Orientation = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return XlChartElementPosition(self.com_object.Position)

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
    def ReadingOrder(self):
        return XlReadingOrder(self.com_object.ReadingOrder)

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.com_object.ReadingOrder = value

    # Lower case aliases for ReadingOrder
    @property
    def readingorder(self):
        return self.ReadingOrder

    @readingorder.setter
    def readingorder(self, value):
        self.ReadingOrder = value

    @property
    def Shadow(self):
        return self.com_object.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.com_object.Shadow = value

    # Lower case aliases for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @shadow.setter
    def shadow(self, value):
        self.Shadow = value

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
    def VerticalAlignment(self):
        return self.com_object.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.com_object.VerticalAlignment = value

    # Lower case aliases for VerticalAlignment
    @property
    def verticalalignment(self):
        return self.VerticalAlignment

    @verticalalignment.setter
    def verticalalignment(self, value):
        self.VerticalAlignment = value

    @property
    def Width(self):
        return self.com_object.Width

    # Lower case aliases for Width
    @property
    def width(self):
        return self.Width

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class Borders:

    def __init__(self, borders=None):
        self.com_object= borders

    def __call__(self, item):
        return Border(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Item(self, BorderType=None):
        arguments = com_arguments([unwrap(a) for a in [BorderType]])
        return LineFormat(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, BorderType=None):
        arguments = [BorderType]
        return self.Item(*arguments)


class Broadcast:

    def __init__(self, broadcast=None):
        self.com_object= broadcast

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AttendeeUrl(self):
        return self.com_object.AttendeeUrl

    # Lower case aliases for AttendeeUrl
    @property
    def attendeeurl(self):
        return self.AttendeeUrl

    @property
    def IsBroadcasting(self):
        return self.com_object.IsBroadcasting

    # Lower case aliases for IsBroadcasting
    @property
    def isbroadcasting(self):
        return self.IsBroadcasting

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    def End(self):
        return self.com_object.End()

    # Lower case alias for End
    def end(self):
        return self.End()

    def Start(self, serverUrl=None):
        arguments = com_arguments([unwrap(a) for a in [serverUrl]])
        self.com_object.Start(*arguments)

    # Lower case alias for Start
    def start(self, serverUrl=None):
        arguments = [serverUrl]
        return self.Start(*arguments)


class BulletFormat:

    def __init__(self, bulletformat=None):
        self.com_object= bulletformat

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Character(self):
        return self.com_object.Character

    @Character.setter
    def Character(self, value):
        self.com_object.Character = value

    # Lower case aliases for Character
    @property
    def character(self):
        return self.Character

    @character.setter
    def character(self, value):
        self.Character = value

    @property
    def Font(self):
        return Font(self.com_object.Font)

    # Lower case aliases for Font
    @property
    def font(self):
        return self.Font

    @property
    def Number(self):
        return self.com_object.Number

    # Lower case aliases for Number
    @property
    def number(self):
        return self.Number

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RelativeSize(self):
        return self.com_object.RelativeSize

    @RelativeSize.setter
    def RelativeSize(self, value):
        self.com_object.RelativeSize = value

    # Lower case aliases for RelativeSize
    @property
    def relativesize(self):
        return self.RelativeSize

    @relativesize.setter
    def relativesize(self, value):
        self.RelativeSize = value

    @property
    def StartValue(self):
        return self.com_object.StartValue

    @StartValue.setter
    def StartValue(self, value):
        self.com_object.StartValue = value

    # Lower case aliases for StartValue
    @property
    def startvalue(self):
        return self.StartValue

    @startvalue.setter
    def startvalue(self, value):
        self.StartValue = value

    @property
    def Style(self):
        return self.com_object.Style

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
    def UseTextColor(self):
        return self.com_object.UseTextColor

    @UseTextColor.setter
    def UseTextColor(self, value):
        self.com_object.UseTextColor = value

    # Lower case aliases for UseTextColor
    @property
    def usetextcolor(self):
        return self.UseTextColor

    @usetextcolor.setter
    def usetextcolor(self, value):
        self.UseTextColor = value

    @property
    def UseTextFont(self):
        return self.com_object.UseTextFont

    @UseTextFont.setter
    def UseTextFont(self, value):
        self.com_object.UseTextFont = value

    # Lower case aliases for UseTextFont
    @property
    def usetextfont(self):
        return self.UseTextFont

    @usetextfont.setter
    def usetextfont(self, value):
        self.UseTextFont = value

    def Picture(self):
        self.com_object.Picture()

    # Lower case alias for Picture
    def picture(self):
        return self.Picture()


class CalloutFormat:

    def __init__(self, calloutformat=None):
        self.com_object= calloutformat

    @property
    def Accent(self):
        return self.com_object.Accent

    @Accent.setter
    def Accent(self, value):
        self.com_object.Accent = value

    # Lower case aliases for Accent
    @property
    def accent(self):
        return self.Accent

    @accent.setter
    def accent(self, value):
        self.Accent = value

    @property
    def Angle(self):
        return self.com_object.Angle

    @Angle.setter
    def Angle(self, value):
        self.com_object.Angle = value

    # Lower case aliases for Angle
    @property
    def angle(self):
        return self.Angle

    @angle.setter
    def angle(self, value):
        self.Angle = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AutoAttach(self):
        return self.com_object.AutoAttach

    @AutoAttach.setter
    def AutoAttach(self, value):
        self.com_object.AutoAttach = value

    # Lower case aliases for AutoAttach
    @property
    def autoattach(self):
        return self.AutoAttach

    @autoattach.setter
    def autoattach(self, value):
        self.AutoAttach = value

    @property
    def AutoLength(self):
        return self.com_object.AutoLength

    # Lower case aliases for AutoLength
    @property
    def autolength(self):
        return self.AutoLength

    @property
    def Border(self):
        return self.com_object.Border

    @Border.setter
    def Border(self, value):
        self.com_object.Border = value

    # Lower case aliases for Border
    @property
    def border(self):
        return self.Border

    @border.setter
    def border(self, value):
        self.Border = value

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Drop(self):
        return self.com_object.Drop

    # Lower case aliases for Drop
    @property
    def drop(self):
        return self.Drop

    @property
    def DropType(self):
        return self.com_object.DropType

    # Lower case aliases for DropType
    @property
    def droptype(self):
        return self.DropType

    @property
    def Gap(self):
        return self.com_object.Gap

    @Gap.setter
    def Gap(self, value):
        self.com_object.Gap = value

    # Lower case aliases for Gap
    @property
    def gap(self):
        return self.Gap

    @gap.setter
    def gap(self, value):
        self.Gap = value

    @property
    def Length(self):
        return self.com_object.Length

    # Lower case aliases for Length
    @property
    def length(self):
        return self.Length

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

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

    def AutomaticLength(self):
        self.com_object.AutomaticLength()

    # Lower case alias for AutomaticLength
    def automaticlength(self):
        return self.AutomaticLength()

    def CustomDrop(self, Drop=None):
        arguments = com_arguments([unwrap(a) for a in [Drop]])
        return self.com_object.CustomDrop(*arguments)

    # Lower case alias for CustomDrop
    def customdrop(self, Drop=None):
        arguments = [Drop]
        return self.CustomDrop(*arguments)

    def CustomLength(self, Length=None):
        arguments = com_arguments([unwrap(a) for a in [Length]])
        return self.com_object.CustomLength(*arguments)

    # Lower case alias for CustomLength
    def customlength(self, Length=None):
        arguments = [Length]
        return self.CustomLength(*arguments)

    def PresetDrop(self, DropType=None):
        arguments = com_arguments([unwrap(a) for a in [DropType]])
        self.com_object.PresetDrop(*arguments)

    # Lower case alias for PresetDrop
    def presetdrop(self, DropType=None):
        arguments = [DropType]
        return self.PresetDrop(*arguments)


class Cell:

    def __init__(self, cell=None):
        self.com_object= cell

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Borders(self):
        return Borders(self.com_object.Borders)

    # Lower case aliases for Borders
    @property
    def borders(self):
        return self.Borders

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Selected(self):
        return self.com_object.Selected

    # Lower case aliases for Selected
    @property
    def selected(self):
        return self.Selected

    @property
    def Shape(self):
        return Shape(self.com_object.Shape)

    # Lower case aliases for Shape
    @property
    def shape(self):
        return self.Shape

    def Merge(self, MergeTo=None):
        arguments = com_arguments([unwrap(a) for a in [MergeTo]])
        self.com_object.Merge(*arguments)

    # Lower case alias for Merge
    def merge(self, MergeTo=None):
        arguments = [MergeTo]
        return self.Merge(*arguments)

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()

    def Split(self, NumRows=None, NumColumns=None):
        arguments = com_arguments([unwrap(a) for a in [NumRows, NumColumns]])
        self.com_object.Split(*arguments)

    # Lower case alias for Split
    def split(self, NumRows=None, NumColumns=None):
        arguments = [NumRows, NumColumns]
        return self.Split(*arguments)


class CellRange:

    def __init__(self, cellrange=None):
        self.com_object= cellrange

    def __call__(self, item):
        return CellRange(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Borders(self):
        return Borders(self.com_object.Borders)

    # Lower case aliases for Borders
    @property
    def borders(self):
        return self.Borders

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

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Cell(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class Chart:

    def __init__(self, chart=None):
        self.com_object= chart

    @property
    def AlternativeText(self):
        return self.com_object.AlternativeText

    @AlternativeText.setter
    def AlternativeText(self, value):
        self.com_object.AlternativeText = value

    # Lower case aliases for AlternativeText
    @property
    def alternativetext(self):
        return self.AlternativeText

    @alternativetext.setter
    def alternativetext(self, value):
        self.AlternativeText = value

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def AutoScaling(self):
        return self.com_object.AutoScaling

    @AutoScaling.setter
    def AutoScaling(self, value):
        self.com_object.AutoScaling = value

    # Lower case aliases for AutoScaling
    @property
    def autoscaling(self):
        return self.AutoScaling

    @autoscaling.setter
    def autoscaling(self, value):
        self.AutoScaling = value

    @property
    def BackWall(self):
        return Walls(self.com_object.BackWall)

    # Lower case aliases for BackWall
    @property
    def backwall(self):
        return self.BackWall

    @property
    def BarShape(self):
        return XlBarShape(self.com_object.BarShape)

    @BarShape.setter
    def BarShape(self, value):
        self.com_object.BarShape = value

    # Lower case aliases for BarShape
    @property
    def barshape(self):
        return self.BarShape

    @barshape.setter
    def barshape(self, value):
        self.BarShape = value

    @property
    def ChartArea(self):
        return ChartArea(self.com_object.ChartArea)

    # Lower case aliases for ChartArea
    @property
    def chartarea(self):
        return self.ChartArea

    @property
    def ChartData(self):
        return ChartData(self.com_object.ChartData)

    # Lower case aliases for ChartData
    @property
    def chartdata(self):
        return self.ChartData

    @property
    def ChartStyle(self):
        return self.com_object.ChartStyle

    @ChartStyle.setter
    def ChartStyle(self, value):
        self.com_object.ChartStyle = value

    # Lower case aliases for ChartStyle
    @property
    def chartstyle(self):
        return self.ChartStyle

    @chartstyle.setter
    def chartstyle(self, value):
        self.ChartStyle = value

    @property
    def ChartTitle(self):
        return ChartTitle(self.com_object.ChartTitle)

    # Lower case aliases for ChartTitle
    @property
    def charttitle(self):
        return self.ChartTitle

    @property
    def ChartType(self):
        return self.com_object.ChartType

    @ChartType.setter
    def ChartType(self, value):
        self.com_object.ChartType = value

    # Lower case aliases for ChartType
    @property
    def charttype(self):
        return self.ChartType

    @charttype.setter
    def charttype(self, value):
        self.ChartType = value

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def DataTable(self):
        return DataTable(self.com_object.DataTable)

    # Lower case aliases for DataTable
    @property
    def datatable(self):
        return self.DataTable

    @property
    def DepthPercent(self):
        return self.com_object.DepthPercent

    @DepthPercent.setter
    def DepthPercent(self, value):
        self.com_object.DepthPercent = value

    # Lower case aliases for DepthPercent
    @property
    def depthpercent(self):
        return self.DepthPercent

    @depthpercent.setter
    def depthpercent(self, value):
        self.DepthPercent = value

    @property
    def DisplayBlanksAs(self):
        return XlDisplayBlanksAs(self.com_object.DisplayBlanksAs)

    @DisplayBlanksAs.setter
    def DisplayBlanksAs(self, value):
        self.com_object.DisplayBlanksAs = value

    # Lower case aliases for DisplayBlanksAs
    @property
    def displayblanksas(self):
        return self.DisplayBlanksAs

    @displayblanksas.setter
    def displayblanksas(self, value):
        self.DisplayBlanksAs = value

    @property
    def Elevation(self):
        return self.com_object.Elevation

    @Elevation.setter
    def Elevation(self, value):
        self.com_object.Elevation = value

    # Lower case aliases for Elevation
    @property
    def elevation(self):
        return self.Elevation

    @elevation.setter
    def elevation(self, value):
        self.Elevation = value

    @property
    def Floor(self):
        return Floor(self.com_object.Floor)

    # Lower case aliases for Floor
    @property
    def floor(self):
        return self.Floor

    @property
    def Format(self):
        return self.com_object.Format

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

    @property
    def GapDepth(self):
        return self.com_object.GapDepth

    @GapDepth.setter
    def GapDepth(self, value):
        self.com_object.GapDepth = value

    # Lower case aliases for GapDepth
    @property
    def gapdepth(self):
        return self.GapDepth

    @gapdepth.setter
    def gapdepth(self, value):
        self.GapDepth = value

    @property
    def HasAxis(self):
        return self.com_object.HasAxis

    @HasAxis.setter
    def HasAxis(self, value):
        self.com_object.HasAxis = value

    # Lower case aliases for HasAxis
    @property
    def hasaxis(self):
        return self.HasAxis

    @hasaxis.setter
    def hasaxis(self, value):
        self.HasAxis = value

    @property
    def HasDataTable(self):
        return self.com_object.HasDataTable

    @HasDataTable.setter
    def HasDataTable(self, value):
        self.com_object.HasDataTable = value

    # Lower case aliases for HasDataTable
    @property
    def hasdatatable(self):
        return self.HasDataTable

    @hasdatatable.setter
    def hasdatatable(self, value):
        self.HasDataTable = value

    @property
    def HasLegend(self):
        return self.com_object.HasLegend

    @HasLegend.setter
    def HasLegend(self, value):
        self.com_object.HasLegend = value

    # Lower case aliases for HasLegend
    @property
    def haslegend(self):
        return self.HasLegend

    @haslegend.setter
    def haslegend(self, value):
        self.HasLegend = value

    @property
    def HasTitle(self):
        return self.com_object.HasTitle

    @HasTitle.setter
    def HasTitle(self, value):
        self.com_object.HasTitle = value

    # Lower case aliases for HasTitle
    @property
    def hastitle(self):
        return self.HasTitle

    @hastitle.setter
    def hastitle(self, value):
        self.HasTitle = value

    @property
    def HeightPercent(self):
        return self.com_object.HeightPercent

    @HeightPercent.setter
    def HeightPercent(self, value):
        self.com_object.HeightPercent = value

    # Lower case aliases for HeightPercent
    @property
    def heightpercent(self):
        return self.HeightPercent

    @heightpercent.setter
    def heightpercent(self, value):
        self.HeightPercent = value

    @property
    def Legend(self):
        return Legend(self.com_object.Legend)

    # Lower case aliases for Legend
    @property
    def legend(self):
        return self.Legend

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
    def Perspective(self):
        return self.com_object.Perspective

    @Perspective.setter
    def Perspective(self, value):
        self.com_object.Perspective = value

    # Lower case aliases for Perspective
    @property
    def perspective(self):
        return self.Perspective

    @perspective.setter
    def perspective(self, value):
        self.Perspective = value

    @property
    def PlotArea(self):
        return PlotArea(self.com_object.PlotArea)

    # Lower case aliases for PlotArea
    @property
    def plotarea(self):
        return self.PlotArea

    @property
    def PlotBy(self):
        return self.com_object.PlotBy

    @PlotBy.setter
    def PlotBy(self, value):
        self.com_object.PlotBy = value

    # Lower case aliases for PlotBy
    @property
    def plotby(self):
        return self.PlotBy

    @plotby.setter
    def plotby(self, value):
        self.PlotBy = value

    @property
    def PlotVisibleOnly(self):
        return self.com_object.PlotVisibleOnly

    @PlotVisibleOnly.setter
    def PlotVisibleOnly(self, value):
        self.com_object.PlotVisibleOnly = value

    # Lower case aliases for PlotVisibleOnly
    @property
    def plotvisibleonly(self):
        return self.PlotVisibleOnly

    @plotvisibleonly.setter
    def plotvisibleonly(self, value):
        self.PlotVisibleOnly = value

    @property
    def RightAngleAxes(self):
        return self.com_object.RightAngleAxes

    @RightAngleAxes.setter
    def RightAngleAxes(self, value):
        self.com_object.RightAngleAxes = value

    # Lower case aliases for RightAngleAxes
    @property
    def rightangleaxes(self):
        return self.RightAngleAxes

    @rightangleaxes.setter
    def rightangleaxes(self, value):
        self.RightAngleAxes = value

    @property
    def Rotation(self):
        return self.com_object.Rotation

    @Rotation.setter
    def Rotation(self, value):
        self.com_object.Rotation = value

    # Lower case aliases for Rotation
    @property
    def rotation(self):
        return self.Rotation

    @rotation.setter
    def rotation(self, value):
        self.Rotation = value

    @property
    def Shapes(self):
        return Shapes(self.com_object.Shapes)

    # Lower case aliases for Shapes
    @property
    def shapes(self):
        return self.Shapes

    @property
    def ShowAllFieldButtons(self):
        return self.com_object.ShowAllFieldButtons

    @ShowAllFieldButtons.setter
    def ShowAllFieldButtons(self, value):
        self.com_object.ShowAllFieldButtons = value

    # Lower case aliases for ShowAllFieldButtons
    @property
    def showallfieldbuttons(self):
        return self.ShowAllFieldButtons

    @showallfieldbuttons.setter
    def showallfieldbuttons(self, value):
        self.ShowAllFieldButtons = value

    @property
    def ShowAxisFieldButtons(self):
        return self.com_object.ShowAxisFieldButtons

    @ShowAxisFieldButtons.setter
    def ShowAxisFieldButtons(self, value):
        self.com_object.ShowAxisFieldButtons = value

    # Lower case aliases for ShowAxisFieldButtons
    @property
    def showaxisfieldbuttons(self):
        return self.ShowAxisFieldButtons

    @showaxisfieldbuttons.setter
    def showaxisfieldbuttons(self, value):
        self.ShowAxisFieldButtons = value

    @property
    def ShowDataLabelsOverMaximum(self):
        return self.com_object.ShowDataLabelsOverMaximum

    @ShowDataLabelsOverMaximum.setter
    def ShowDataLabelsOverMaximum(self, value):
        self.com_object.ShowDataLabelsOverMaximum = value

    # Lower case aliases for ShowDataLabelsOverMaximum
    @property
    def showdatalabelsovermaximum(self):
        return self.ShowDataLabelsOverMaximum

    @showdatalabelsovermaximum.setter
    def showdatalabelsovermaximum(self, value):
        self.ShowDataLabelsOverMaximum = value

    @property
    def ShowLegendFieldButtons(self):
        return self.com_object.ShowLegendFieldButtons

    @ShowLegendFieldButtons.setter
    def ShowLegendFieldButtons(self, value):
        self.com_object.ShowLegendFieldButtons = value

    # Lower case aliases for ShowLegendFieldButtons
    @property
    def showlegendfieldbuttons(self):
        return self.ShowLegendFieldButtons

    @showlegendfieldbuttons.setter
    def showlegendfieldbuttons(self, value):
        self.ShowLegendFieldButtons = value

    @property
    def ShowReportFilterFieldButtons(self):
        return self.com_object.ShowReportFilterFieldButtons

    @ShowReportFilterFieldButtons.setter
    def ShowReportFilterFieldButtons(self, value):
        self.com_object.ShowReportFilterFieldButtons = value

    # Lower case aliases for ShowReportFilterFieldButtons
    @property
    def showreportfilterfieldbuttons(self):
        return self.ShowReportFilterFieldButtons

    @showreportfilterfieldbuttons.setter
    def showreportfilterfieldbuttons(self, value):
        self.ShowReportFilterFieldButtons = value

    @property
    def ShowValueFieldButtons(self):
        return self.com_object.ShowValueFieldButtons

    @ShowValueFieldButtons.setter
    def ShowValueFieldButtons(self, value):
        self.com_object.ShowValueFieldButtons = value

    # Lower case aliases for ShowValueFieldButtons
    @property
    def showvaluefieldbuttons(self):
        return self.ShowValueFieldButtons

    @showvaluefieldbuttons.setter
    def showvaluefieldbuttons(self, value):
        self.ShowValueFieldButtons = value

    @property
    def SideWall(self):
        return Walls(self.com_object.SideWall)

    # Lower case aliases for SideWall
    @property
    def sidewall(self):
        return self.SideWall

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
    def Walls(self):
        return Walls(self.com_object.Walls)

    # Lower case aliases for Walls
    @property
    def walls(self):
        return self.Walls

    def ApplyChartTemplate(self, FileName=None):
        arguments = com_arguments([unwrap(a) for a in [FileName]])
        self.com_object.ApplyChartTemplate(*arguments)

    # Lower case alias for ApplyChartTemplate
    def applycharttemplate(self, FileName=None):
        arguments = [FileName]
        return self.ApplyChartTemplate(*arguments)

    def ApplyDataLabels(self, Type=None, LegendKey=None, AutoText=None, HasLeaderLines=None, ShowSeriesName=None, ShowCategoryName=None, ShowValue=None, ShowPercentage=None, ShowBubbleSize=None, Separator=None):
        arguments = com_arguments([unwrap(a) for a in [Type, LegendKey, AutoText, HasLeaderLines, ShowSeriesName, ShowCategoryName, ShowValue, ShowPercentage, ShowBubbleSize, Separator]])
        self.com_object.ApplyDataLabels(*arguments)

    # Lower case alias for ApplyDataLabels
    def applydatalabels(self, Type=None, LegendKey=None, AutoText=None, HasLeaderLines=None, ShowSeriesName=None, ShowCategoryName=None, ShowValue=None, ShowPercentage=None, ShowBubbleSize=None, Separator=None):
        arguments = [Type, LegendKey, AutoText, HasLeaderLines, ShowSeriesName, ShowCategoryName, ShowValue, ShowPercentage, ShowBubbleSize, Separator]
        return self.ApplyDataLabels(*arguments)

    def ApplyLayout(self, Layout=None, ChartType=None):
        arguments = com_arguments([unwrap(a) for a in [Layout, ChartType]])
        self.com_object.ApplyLayout(*arguments)

    # Lower case alias for ApplyLayout
    def applylayout(self, Layout=None, ChartType=None):
        arguments = [Layout, ChartType]
        return self.ApplyLayout(*arguments)

    def Axes(self, Type=None, AxisGroup=None):
        arguments = com_arguments([unwrap(a) for a in [Type, AxisGroup]])
        return self.com_object.Axes(*arguments)

    # Lower case alias for Axes
    def axes(self, Type=None, AxisGroup=None):
        arguments = [Type, AxisGroup]
        return self.Axes(*arguments)

    def ChartGroups(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        self.com_object.ChartGroups(*arguments)

    # Lower case alias for ChartGroups
    def chartgroups(self, Index=None):
        arguments = [Index]
        return self.ChartGroups(*arguments)

    def ChartWizard(self, Source=None, Gallery=None, Format=None, PlotBy=None, CategoryLabels=None, SeriesLabels=None, HasLegend=None, Title=None, CategoryTitle=None, ValueTitle=None, ExtraTitle=None):
        arguments = com_arguments([unwrap(a) for a in [Source, Gallery, Format, PlotBy, CategoryLabels, SeriesLabels, HasLegend, Title, CategoryTitle, ValueTitle, ExtraTitle]])
        self.com_object.ChartWizard(*arguments)

    # Lower case alias for ChartWizard
    def chartwizard(self, Source=None, Gallery=None, Format=None, PlotBy=None, CategoryLabels=None, SeriesLabels=None, HasLegend=None, Title=None, CategoryTitle=None, ValueTitle=None, ExtraTitle=None):
        arguments = [Source, Gallery, Format, PlotBy, CategoryLabels, SeriesLabels, HasLegend, Title, CategoryTitle, ValueTitle, ExtraTitle]
        return self.ChartWizard(*arguments)

    def ClearToMatchStyle(self):
        self.com_object.ClearToMatchStyle()

    # Lower case alias for ClearToMatchStyle
    def cleartomatchstyle(self):
        return self.ClearToMatchStyle()

    def Copy(self, Before=None, After=None):
        arguments = com_arguments([unwrap(a) for a in [Before, After]])
        self.com_object.Copy(*arguments)

    # Lower case alias for Copy
    def copy(self, Before=None, After=None):
        arguments = [Before, After]
        return self.Copy(*arguments)

    def CopyPicture(self, Appearance=None, Format=None, Size=None):
        arguments = com_arguments([unwrap(a) for a in [Appearance, Format, Size]])
        self.com_object.CopyPicture(*arguments)

    # Lower case alias for CopyPicture
    def copypicture(self, Appearance=None, Format=None, Size=None):
        arguments = [Appearance, Format, Size]
        return self.CopyPicture(*arguments)

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Export(self, FileName=None, FilterName=None, Interactive=None):
        arguments = com_arguments([unwrap(a) for a in [FileName, FilterName, Interactive]])
        self.com_object.Export(*arguments)

    # Lower case alias for Export
    def export(self, FileName=None, FilterName=None, Interactive=None):
        arguments = [FileName, FilterName, Interactive]
        return self.Export(*arguments)

    def GetChartElement(self, x=None, y=None, ElementID=None, Arg1=None, Arg2=None):
        arguments = com_arguments([unwrap(a) for a in [x, y, ElementID, Arg1, Arg2]])
        self.com_object.GetChartElement(*arguments)

    # Lower case alias for GetChartElement
    def getchartelement(self, x=None, y=None, ElementID=None, Arg1=None, Arg2=None):
        arguments = [x, y, ElementID, Arg1, Arg2]
        return self.GetChartElement(*arguments)

    def Paste(self, Type=None):
        arguments = com_arguments([unwrap(a) for a in [Type]])
        self.com_object.Paste(*arguments)

    # Lower case alias for Paste
    def paste(self, Type=None):
        arguments = [Type]
        return self.Paste(*arguments)

    def Refresh(self):
        self.com_object.Refresh()

    # Lower case alias for Refresh
    def refresh(self):
        return self.Refresh()

    def SaveChartTemplate(self, FileName=None):
        arguments = com_arguments([unwrap(a) for a in [FileName]])
        self.com_object.SaveChartTemplate(*arguments)

    # Lower case alias for SaveChartTemplate
    def savecharttemplate(self, FileName=None):
        arguments = [FileName]
        return self.SaveChartTemplate(*arguments)

    def Select(self, Replace=None):
        arguments = com_arguments([unwrap(a) for a in [Replace]])
        self.com_object.Select(*arguments)

    # Lower case alias for Select
    def select(self, Replace=None):
        arguments = [Replace]
        return self.Select(*arguments)

    def SeriesCollection(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return SeriesCollection(self.com_object.SeriesCollection(*arguments))

    # Lower case alias for SeriesCollection
    def seriescollection(self, Index=None):
        arguments = [Index]
        return self.SeriesCollection(*arguments)

    def SetBackgroundPicture(self, FileName=None):
        arguments = com_arguments([unwrap(a) for a in [FileName]])
        self.com_object.SetBackgroundPicture(*arguments)

    # Lower case alias for SetBackgroundPicture
    def setbackgroundpicture(self, FileName=None):
        arguments = [FileName]
        return self.SetBackgroundPicture(*arguments)

    def SetDefaultChart(self, Name=None):
        arguments = com_arguments([unwrap(a) for a in [Name]])
        self.com_object.SetDefaultChart(*arguments)

    # Lower case alias for SetDefaultChart
    def setdefaultchart(self, Name=None):
        arguments = [Name]
        return self.SetDefaultChart(*arguments)

    def SetElement(self, Element=None):
        arguments = com_arguments([unwrap(a) for a in [Element]])
        self.com_object.SetElement(*arguments)

    # Lower case alias for SetElement
    def setelement(self, Element=None):
        arguments = [Element]
        return self.SetElement(*arguments)

    def SetSourceData(self, Source=None, PlotBy=None):
        arguments = com_arguments([unwrap(a) for a in [Source, PlotBy]])
        self.com_object.SetSourceData(*arguments)

    # Lower case alias for SetSourceData
    def setsourcedata(self, Source=None, PlotBy=None):
        arguments = [Source, PlotBy]
        return self.SetSourceData(*arguments)


class ChartArea:

    def __init__(self, chartarea=None):
        self.com_object= chartarea

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

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
    def Shadow(self):
        return self.com_object.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.com_object.Shadow = value

    # Lower case aliases for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @shadow.setter
    def shadow(self, value):
        self.Shadow = value

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

    def Clear(self):
        self.com_object.Clear()

    # Lower case alias for Clear
    def clear(self):
        return self.Clear()

    def ClearContents(self):
        self.com_object.ClearContents()

    # Lower case alias for ClearContents
    def clearcontents(self):
        return self.ClearContents()

    def ClearFormats(self):
        self.com_object.ClearFormats()

    # Lower case alias for ClearFormats
    def clearformats(self):
        return self.ClearFormats()

    def Copy(self):
        self.com_object.Copy()

    # Lower case alias for Copy
    def copy(self):
        return self.Copy()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class ChartBorder:

    def __init__(self, chartborder=None):
        self.com_object= chartborder

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Color(self):
        return self.com_object.Color

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
    def ColorIndex(self):
        return self.com_object.ColorIndex

    @ColorIndex.setter
    def ColorIndex(self, value):
        self.com_object.ColorIndex = value

    # Lower case aliases for ColorIndex
    @property
    def colorindex(self):
        return self.ColorIndex

    @colorindex.setter
    def colorindex(self, value):
        self.ColorIndex = value

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def LineStyle(self):
        return XlLineStyle(self.com_object.LineStyle)

    @LineStyle.setter
    def LineStyle(self, value):
        self.com_object.LineStyle = value

    # Lower case aliases for LineStyle
    @property
    def linestyle(self):
        return self.LineStyle

    @linestyle.setter
    def linestyle(self, value):
        self.LineStyle = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Weight(self):
        return XlBorderWeight(self.com_object.Weight)

    @Weight.setter
    def Weight(self, value):
        self.com_object.Weight = value

    # Lower case aliases for Weight
    @property
    def weight(self):
        return self.Weight

    @weight.setter
    def weight(self, value):
        self.Weight = value


class ChartCharacters:

    def __init__(self, chartcharacters=None):
        self.com_object= chartcharacters

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Caption(self):
        return self.com_object.Caption

    # Lower case aliases for Caption
    @property
    def caption(self):
        return self.Caption

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Font(self):
        return ChartFont(self.com_object.Font)

    # Lower case aliases for Font
    @property
    def font(self):
        return self.Font

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PhoneticCharacters(self):
        return self.com_object.PhoneticCharacters

    @PhoneticCharacters.setter
    def PhoneticCharacters(self, value):
        self.com_object.PhoneticCharacters = value

    # Lower case aliases for PhoneticCharacters
    @property
    def phoneticcharacters(self):
        return self.PhoneticCharacters

    @phoneticcharacters.setter
    def phoneticcharacters(self, value):
        self.PhoneticCharacters = value

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

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Insert(self, String=None):
        arguments = com_arguments([unwrap(a) for a in [String]])
        self.com_object.Insert(*arguments)

    # Lower case alias for Insert
    def insert(self, String=None):
        arguments = [String]
        return self.Insert(*arguments)


class ChartData:

    def __init__(self, chartdata=None):
        self.com_object= chartdata

    @property
    def IsLinked(self):
        return self.com_object.IsLinked

    # Lower case aliases for IsLinked
    @property
    def islinked(self):
        return self.IsLinked

    @property
    def Workbook(self):
        return self.com_object.Workbook

    # Lower case aliases for Workbook
    @property
    def workbook(self):
        return self.Workbook

    def Activate(self):
        self.com_object.Activate()

    # Lower case alias for Activate
    def activate(self):
        return self.Activate()

    def BreakLink(self):
        self.com_object.BreakLink()

    # Lower case alias for BreakLink
    def breaklink(self):
        return self.BreakLink()


class ChartFont:

    def __init__(self, chartfont=None):
        self.com_object= chartfont

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Background(self):
        return XlBackground(self.com_object.Background)

    @Background.setter
    def Background(self, value):
        self.com_object.Background = value

    # Lower case aliases for Background
    @property
    def background(self):
        return self.Background

    @background.setter
    def background(self, value):
        self.Background = value

    @property
    def Bold(self):
        return self.com_object.Bold

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
    def Color(self):
        return self.com_object.Color

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
    def ColorIndex(self):
        return self.com_object.ColorIndex

    @ColorIndex.setter
    def ColorIndex(self, value):
        self.com_object.ColorIndex = value

    # Lower case aliases for ColorIndex
    @property
    def colorindex(self):
        return self.ColorIndex

    @colorindex.setter
    def colorindex(self, value):
        self.ColorIndex = value

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def FontStyle(self):
        return self.com_object.FontStyle

    @FontStyle.setter
    def FontStyle(self, value):
        self.com_object.FontStyle = value

    # Lower case aliases for FontStyle
    @property
    def fontstyle(self):
        return self.FontStyle

    @fontstyle.setter
    def fontstyle(self, value):
        self.FontStyle = value

    @property
    def Italic(self):
        return self.com_object.Italic

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
    def StrikeThrough(self):
        return self.com_object.StrikeThrough

    @StrikeThrough.setter
    def StrikeThrough(self, value):
        self.com_object.StrikeThrough = value

    # Lower case aliases for StrikeThrough
    @property
    def strikethrough(self):
        return self.StrikeThrough

    @strikethrough.setter
    def strikethrough(self, value):
        self.StrikeThrough = value

    @property
    def Subscript(self):
        return self.com_object.Subscript

    @Subscript.setter
    def Subscript(self, value):
        self.com_object.Subscript = value

    # Lower case aliases for Subscript
    @property
    def subscript(self):
        return self.Subscript

    @subscript.setter
    def subscript(self, value):
        self.Subscript = value

    @property
    def Underline(self):
        return XlUnderlineStyle(self.com_object.Underline)

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


class ChartFormat:

    def __init__(self, chartformat=None):
        self.com_object= chartformat

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Fill(self):
        return FillFormat(self.com_object.Fill)

    # Lower case aliases for Fill
    @property
    def fill(self):
        return self.Fill

    @property
    def Glow(self):
        return self.com_object.Glow

    # Lower case aliases for Glow
    @property
    def glow(self):
        return self.Glow

    @property
    def Line(self):
        return LineFormat(self.com_object.Line)

    # Lower case aliases for Line
    @property
    def line(self):
        return self.Line

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PictureFormat(self):
        return PictureFormat(self.com_object.PictureFormat)

    # Lower case aliases for PictureFormat
    @property
    def pictureformat(self):
        return self.PictureFormat

    @property
    def Shadow(self):
        return ShadowFormat(self.com_object.Shadow)

    # Lower case aliases for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @property
    def SoftEdge(self):
        return self.com_object.SoftEdge

    # Lower case aliases for SoftEdge
    @property
    def softedge(self):
        return self.SoftEdge

    @property
    def TextFrame2(self):
        return TextFrame2(self.com_object.TextFrame2)

    # Lower case aliases for TextFrame2
    @property
    def textframe2(self):
        return self.TextFrame2

    @property
    def ThreeD(self):
        return ThreeDFormat(self.com_object.ThreeD)

    # Lower case aliases for ThreeD
    @property
    def threed(self):
        return self.ThreeD


class ChartGroup:

    def __init__(self, chartgroup=None):
        self.com_object= chartgroup

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def AxisGroup(self):
        return XlAxisGroup(self.com_object.AxisGroup)

    @AxisGroup.setter
    def AxisGroup(self, value):
        self.com_object.AxisGroup = value

    # Lower case aliases for AxisGroup
    @property
    def axisgroup(self):
        return self.AxisGroup

    @axisgroup.setter
    def axisgroup(self, value):
        self.AxisGroup = value

    @property
    def BubbleScale(self):
        return self.com_object.BubbleScale

    @BubbleScale.setter
    def BubbleScale(self, value):
        self.com_object.BubbleScale = value

    # Lower case aliases for BubbleScale
    @property
    def bubblescale(self):
        return self.BubbleScale

    @bubblescale.setter
    def bubblescale(self, value):
        self.BubbleScale = value

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def DoughnutHoleSize(self):
        return self.com_object.DoughnutHoleSize

    @DoughnutHoleSize.setter
    def DoughnutHoleSize(self, value):
        self.com_object.DoughnutHoleSize = value

    # Lower case aliases for DoughnutHoleSize
    @property
    def doughnutholesize(self):
        return self.DoughnutHoleSize

    @doughnutholesize.setter
    def doughnutholesize(self, value):
        self.DoughnutHoleSize = value

    @property
    def DownBars(self):
        return DownBars(self.com_object.DownBars)

    # Lower case aliases for DownBars
    @property
    def downbars(self):
        return self.DownBars

    @property
    def DropLines(self):
        return DropLines(self.com_object.DropLines)

    # Lower case aliases for DropLines
    @property
    def droplines(self):
        return self.DropLines

    @property
    def FirstSliceAngle(self):
        return self.com_object.FirstSliceAngle

    @FirstSliceAngle.setter
    def FirstSliceAngle(self, value):
        self.com_object.FirstSliceAngle = value

    # Lower case aliases for FirstSliceAngle
    @property
    def firstsliceangle(self):
        return self.FirstSliceAngle

    @firstsliceangle.setter
    def firstsliceangle(self, value):
        self.FirstSliceAngle = value

    @property
    def GapWidth(self):
        return self.com_object.GapWidth

    @GapWidth.setter
    def GapWidth(self, value):
        self.com_object.GapWidth = value

    # Lower case aliases for GapWidth
    @property
    def gapwidth(self):
        return self.GapWidth

    @gapwidth.setter
    def gapwidth(self, value):
        self.GapWidth = value

    @property
    def Has3DShading(self):
        return self.com_object.Has3DShading

    @Has3DShading.setter
    def Has3DShading(self, value):
        self.com_object.Has3DShading = value

    # Lower case aliases for Has3DShading
    @property
    def has3dshading(self):
        return self.Has3DShading

    @has3dshading.setter
    def has3dshading(self, value):
        self.Has3DShading = value

    @property
    def HasDropLines(self):
        return self.com_object.HasDropLines

    @HasDropLines.setter
    def HasDropLines(self, value):
        self.com_object.HasDropLines = value

    # Lower case aliases for HasDropLines
    @property
    def hasdroplines(self):
        return self.HasDropLines

    @hasdroplines.setter
    def hasdroplines(self, value):
        self.HasDropLines = value

    @property
    def HasHiLoLines(self):
        return self.com_object.HasHiLoLines

    @HasHiLoLines.setter
    def HasHiLoLines(self, value):
        self.com_object.HasHiLoLines = value

    # Lower case aliases for HasHiLoLines
    @property
    def hashilolines(self):
        return self.HasHiLoLines

    @hashilolines.setter
    def hashilolines(self, value):
        self.HasHiLoLines = value

    @property
    def HasRadarAxisLabels(self):
        return self.com_object.HasRadarAxisLabels

    @HasRadarAxisLabels.setter
    def HasRadarAxisLabels(self, value):
        self.com_object.HasRadarAxisLabels = value

    # Lower case aliases for HasRadarAxisLabels
    @property
    def hasradaraxislabels(self):
        return self.HasRadarAxisLabels

    @hasradaraxislabels.setter
    def hasradaraxislabels(self, value):
        self.HasRadarAxisLabels = value

    @property
    def HasSeriesLines(self):
        return self.com_object.HasSeriesLines

    @HasSeriesLines.setter
    def HasSeriesLines(self, value):
        self.com_object.HasSeriesLines = value

    # Lower case aliases for HasSeriesLines
    @property
    def hasserieslines(self):
        return self.HasSeriesLines

    @hasserieslines.setter
    def hasserieslines(self, value):
        self.HasSeriesLines = value

    @property
    def HasUpDownBars(self):
        return self.com_object.HasUpDownBars

    @HasUpDownBars.setter
    def HasUpDownBars(self, value):
        self.com_object.HasUpDownBars = value

    # Lower case aliases for HasUpDownBars
    @property
    def hasupdownbars(self):
        return self.HasUpDownBars

    @hasupdownbars.setter
    def hasupdownbars(self, value):
        self.HasUpDownBars = value

    @property
    def HiLoLines(self):
        return HiLoLines(self.com_object.HiLoLines)

    # Lower case aliases for HiLoLines
    @property
    def hilolines(self):
        return self.HiLoLines

    @property
    def Index(self):
        return self.com_object.Index

    # Lower case aliases for Index
    @property
    def index(self):
        return self.Index

    @property
    def Overlap(self):
        return self.com_object.Overlap

    @Overlap.setter
    def Overlap(self, value):
        self.com_object.Overlap = value

    # Lower case aliases for Overlap
    @property
    def overlap(self):
        return self.Overlap

    @overlap.setter
    def overlap(self, value):
        self.Overlap = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RadarAxisLabels(self):
        return TickLabels(self.com_object.RadarAxisLabels)

    # Lower case aliases for RadarAxisLabels
    @property
    def radaraxislabels(self):
        return self.RadarAxisLabels

    @property
    def SecondPlotSize(self):
        return self.com_object.SecondPlotSize

    @SecondPlotSize.setter
    def SecondPlotSize(self, value):
        self.com_object.SecondPlotSize = value

    # Lower case aliases for SecondPlotSize
    @property
    def secondplotsize(self):
        return self.SecondPlotSize

    @secondplotsize.setter
    def secondplotsize(self, value):
        self.SecondPlotSize = value

    @property
    def SeriesLines(self):
        return SeriesLines(self.com_object.SeriesLines)

    # Lower case aliases for SeriesLines
    @property
    def serieslines(self):
        return self.SeriesLines

    @property
    def ShowNegativeBubbles(self):
        return self.com_object.ShowNegativeBubbles

    @ShowNegativeBubbles.setter
    def ShowNegativeBubbles(self, value):
        self.com_object.ShowNegativeBubbles = value

    # Lower case aliases for ShowNegativeBubbles
    @property
    def shownegativebubbles(self):
        return self.ShowNegativeBubbles

    @shownegativebubbles.setter
    def shownegativebubbles(self, value):
        self.ShowNegativeBubbles = value

    @property
    def SizeRepresents(self):
        return self.com_object.SizeRepresents

    @SizeRepresents.setter
    def SizeRepresents(self, value):
        self.com_object.SizeRepresents = value

    # Lower case aliases for SizeRepresents
    @property
    def sizerepresents(self):
        return self.SizeRepresents

    @sizerepresents.setter
    def sizerepresents(self, value):
        self.SizeRepresents = value

    @property
    def SplitType(self):
        return XlChartSplitType(self.com_object.SplitType)

    @SplitType.setter
    def SplitType(self, value):
        self.com_object.SplitType = value

    # Lower case aliases for SplitType
    @property
    def splittype(self):
        return self.SplitType

    @splittype.setter
    def splittype(self, value):
        self.SplitType = value

    @property
    def SplitValue(self):
        return self.com_object.SplitValue

    @SplitValue.setter
    def SplitValue(self, value):
        self.com_object.SplitValue = value

    # Lower case aliases for SplitValue
    @property
    def splitvalue(self):
        return self.SplitValue

    @splitvalue.setter
    def splitvalue(self, value):
        self.SplitValue = value

    @property
    def UpBars(self):
        return UpBars(self.com_object.UpBars)

    # Lower case aliases for UpBars
    @property
    def upbars(self):
        return self.UpBars

    @property
    def VaryByCategories(self):
        return self.com_object.VaryByCategories

    @VaryByCategories.setter
    def VaryByCategories(self, value):
        self.com_object.VaryByCategories = value

    # Lower case aliases for VaryByCategories
    @property
    def varybycategories(self):
        return self.VaryByCategories

    @varybycategories.setter
    def varybycategories(self, value):
        self.VaryByCategories = value

    def SeriesCollection(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return SeriesCollection(self.com_object.SeriesCollection(*arguments))

    # Lower case alias for SeriesCollection
    def seriescollection(self, Index=None):
        arguments = [Index]
        return self.SeriesCollection(*arguments)


class ChartGroups:

    def __init__(self, chartgroups=None):
        self.com_object= chartgroups

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return ChartGroup(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class ChartTitle:

    def __init__(self, charttitle=None):
        self.com_object= charttitle

    @property
    def Application(self):
        return self.com_object.Application

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

    def Characters(self, Start=None, Length=None):
        arguments = com_arguments([unwrap(a) for a in [Start, Length]])
        if hasattr(self.com_object, "GetCharacters"):
            return ChartCharacters(self.com_object.GetCharacters(*arguments))
        else:
            return ChartCharacters(self.com_object.Characters(*arguments))

    # Lower case aliases for Characters
    def characters(self, Start=None, Length=None):
        arguments = [Start, Length]
        return self.Characters(*arguments)

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

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
    def FormulaLocal(self):
        return self.com_object.FormulaLocal

    @FormulaLocal.setter
    def FormulaLocal(self, value):
        self.com_object.FormulaLocal = value

    # Lower case aliases for FormulaLocal
    @property
    def formulalocal(self):
        return self.FormulaLocal

    @formulalocal.setter
    def formulalocal(self, value):
        self.FormulaLocal = value

    @property
    def FormulaR1C1(self):
        return self.com_object.FormulaR1C1

    @FormulaR1C1.setter
    def FormulaR1C1(self, value):
        self.com_object.FormulaR1C1 = value

    # Lower case aliases for FormulaR1C1
    @property
    def formular1c1(self):
        return self.FormulaR1C1

    @formular1c1.setter
    def formular1c1(self, value):
        self.FormulaR1C1 = value

    @property
    def FormulaR1C1Local(self):
        return self.com_object.FormulaR1C1Local

    @FormulaR1C1Local.setter
    def FormulaR1C1Local(self, value):
        self.com_object.FormulaR1C1Local = value

    # Lower case aliases for FormulaR1C1Local
    @property
    def formular1c1local(self):
        return self.FormulaR1C1Local

    @formular1c1local.setter
    def formular1c1local(self, value):
        self.FormulaR1C1Local = value

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
    def HorizontalAlignment(self):
        return self.com_object.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.com_object.HorizontalAlignment = value

    # Lower case aliases for HorizontalAlignment
    @property
    def horizontalalignment(self):
        return self.HorizontalAlignment

    @horizontalalignment.setter
    def horizontalalignment(self, value):
        self.HorizontalAlignment = value

    @property
    def IncludeInLayout(self):
        return self.com_object.IncludeInLayout

    @IncludeInLayout.setter
    def IncludeInLayout(self, value):
        self.com_object.IncludeInLayout = value

    # Lower case aliases for IncludeInLayout
    @property
    def includeinlayout(self):
        return self.IncludeInLayout

    @includeinlayout.setter
    def includeinlayout(self, value):
        self.IncludeInLayout = value

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
    def Name(self):
        return self.com_object.Name

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def Orientation(self):
        return self.com_object.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.com_object.Orientation = value

    # Lower case aliases for Orientation
    @property
    def orientation(self):
        return self.Orientation

    @orientation.setter
    def orientation(self, value):
        self.Orientation = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return XlChartElementPosition(self.com_object.Position)

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
    def ReadingOrder(self):
        return XlReadingOrder(self.com_object.ReadingOrder)

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.com_object.ReadingOrder = value

    # Lower case aliases for ReadingOrder
    @property
    def readingorder(self):
        return self.ReadingOrder

    @readingorder.setter
    def readingorder(self, value):
        self.ReadingOrder = value

    @property
    def Shadow(self):
        return self.com_object.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.com_object.Shadow = value

    # Lower case aliases for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @shadow.setter
    def shadow(self, value):
        self.Shadow = value

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
    def VerticalAlignment(self):
        return self.com_object.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.com_object.VerticalAlignment = value

    # Lower case aliases for VerticalAlignment
    @property
    def verticalalignment(self):
        return self.VerticalAlignment

    @verticalalignment.setter
    def verticalalignment(self, value):
        self.VerticalAlignment = value

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

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class Coauthoring:

    def __init__(self, coauthoring=None):
        self.com_object= coauthoring

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def FavorServerEditsDuringMerge(self):
        return self.com_object.FavorServerEditsDuringMerge

    @FavorServerEditsDuringMerge.setter
    def FavorServerEditsDuringMerge(self, value):
        self.com_object.FavorServerEditsDuringMerge = value

    # Lower case aliases for FavorServerEditsDuringMerge
    @property
    def favorservereditsduringmerge(self):
        return self.FavorServerEditsDuringMerge

    @favorservereditsduringmerge.setter
    def favorservereditsduringmerge(self, value):
        self.FavorServerEditsDuringMerge = value

    @property
    def MergeMode(self):
        return self.com_object.MergeMode

    # Lower case aliases for MergeMode
    @property
    def mergemode(self):
        return self.MergeMode

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PendingUpdates(self):
        return self.com_object.PendingUpdates

    # Lower case aliases for PendingUpdates
    @property
    def pendingupdates(self):
        return self.PendingUpdates

    def EndReview(self):
        self.com_object.EndReview()

    # Lower case alias for EndReview
    def endreview(self):
        return self.EndReview()


class ColorEffect:

    def __init__(self, coloreffect=None):
        self.com_object= coloreffect

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def By(self):
        return ColorFormat(self.com_object.By)

    # Lower case aliases for By
    @property
    def by(self):
        return self.By

    @property
    def From(self):
        return self.com_object.From

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

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


class ColorFormat:

    def __init__(self, colorformat=None):
        self.com_object= colorformat

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Brightness(self):
        return self.com_object.Brightness

    @Brightness.setter
    def Brightness(self, value):
        self.com_object.Brightness = value

    # Lower case aliases for Brightness
    @property
    def brightness(self):
        return self.Brightness

    @brightness.setter
    def brightness(self, value):
        self.Brightness = value

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def ObjectThemeColor(self):
        return ColorFormat(self.com_object.ObjectThemeColor)

    @ObjectThemeColor.setter
    def ObjectThemeColor(self, value):
        self.com_object.ObjectThemeColor = value

    # Lower case aliases for ObjectThemeColor
    @property
    def objectthemecolor(self):
        return self.ObjectThemeColor

    @objectthemecolor.setter
    def objectthemecolor(self, value):
        self.ObjectThemeColor = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RGB(self):
        return self.com_object.RGB

    @RGB.setter
    def RGB(self, value):
        self.com_object.RGB = value

    # Lower case aliases for RGB
    @property
    def rgb(self):
        return self.RGB

    @rgb.setter
    def rgb(self, value):
        self.RGB = value

    @property
    def SchemeColor(self):
        return self.com_object.SchemeColor

    @SchemeColor.setter
    def SchemeColor(self, value):
        self.com_object.SchemeColor = value

    # Lower case aliases for SchemeColor
    @property
    def schemecolor(self):
        return self.SchemeColor

    @schemecolor.setter
    def schemecolor(self, value):
        self.SchemeColor = value

    @property
    def TintAndShade(self):
        return self.com_object.TintAndShade

    @TintAndShade.setter
    def TintAndShade(self, value):
        self.com_object.TintAndShade = value

    # Lower case aliases for TintAndShade
    @property
    def tintandshade(self):
        return self.TintAndShade

    @tintandshade.setter
    def tintandshade(self, value):
        self.TintAndShade = value

    @property
    def Type(self):
        return self.com_object.Type

    # Lower case aliases for Type
    @property
    def type(self):
        return self.Type


class ColorScheme:

    def __init__(self, colorscheme=None):
        self.com_object= colorscheme

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Colors(self, SchemeColor=None):
        arguments = com_arguments([unwrap(a) for a in [SchemeColor]])
        return RGBColor(self.com_object.Colors(*arguments))

    # Lower case alias for Colors
    def colors(self, SchemeColor=None):
        arguments = [SchemeColor]
        return self.Colors(*arguments)

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()


class ColorSchemes:

    def __init__(self, colorschemes=None):
        self.com_object= colorschemes

    def __call__(self, item):
        return ColorScheme(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Add(self, Scheme=None):
        arguments = com_arguments([unwrap(a) for a in [Scheme]])
        return ColorScheme(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Scheme=None):
        arguments = [Scheme]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return ColorScheme(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class Column:

    def __init__(self, column=None):
        self.com_object= column

    @property
    def Application(self):
        return Application(self.com_object.Application)

    def Cells(self, RowIndex=None, ColumnIndex=None):
        arguments = com_arguments([unwrap(a) for a in [RowIndex, ColumnIndex]])
        if hasattr(self.com_object, "GetCells"):
            return CellRange(self.com_object.GetCells(*arguments))
        else:
            return CellRange(self.com_object.Cells(*arguments))

    # Lower case aliases for Cells
    def cells(self, RowIndex=None, ColumnIndex=None):
        arguments = [RowIndex, ColumnIndex]
        return self.Cells(*arguments)

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

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

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class Columns:

    def __init__(self, columns=None):
        self.com_object= columns

    def __call__(self, item):
        return Column(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Add(self, BeforeColumn=None):
        arguments = com_arguments([unwrap(a) for a in [BeforeColumn]])
        return Column(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, BeforeColumn=None):
        arguments = [BeforeColumn]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Column(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class CommandEffect:

    def __init__(self, commandeffect=None):
        self.com_object= commandeffect

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Bookmark(self):
        return self.com_object.Bookmark

    @Bookmark.setter
    def Bookmark(self, value):
        self.com_object.Bookmark = value

    # Lower case aliases for Bookmark
    @property
    def bookmark(self):
        return self.Bookmark

    @bookmark.setter
    def bookmark(self, value):
        self.Bookmark = value

    @property
    def Command(self):
        return self.com_object.Command

    @Command.setter
    def Command(self, value):
        self.com_object.Command = value

    # Lower case aliases for Command
    @property
    def command(self):
        return self.Command

    @command.setter
    def command(self, value):
        self.Command = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

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


class Comment:

    def __init__(self, comment=None):
        self.com_object= comment

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Author(self):
        return Comment(self.com_object.Author)

    # Lower case aliases for Author
    @property
    def author(self):
        return self.Author

    @property
    def AuthorIndex(self):
        return self.com_object.AuthorIndex

    # Lower case aliases for AuthorIndex
    @property
    def authorindex(self):
        return self.AuthorIndex

    @property
    def AuthorInitials(self):
        return Comment(self.com_object.AuthorInitials)

    # Lower case aliases for AuthorInitials
    @property
    def authorinitials(self):
        return self.AuthorInitials

    @property
    def DateTime(self):
        return self.com_object.DateTime

    # Lower case aliases for DateTime
    @property
    def datetime(self):
        return self.DateTime

    @property
    def Left(self):
        return self.com_object.Left

    # Lower case aliases for Left
    @property
    def left(self):
        return self.Left

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Text(self):
        return self.com_object.Text

    # Lower case aliases for Text
    @property
    def text(self):
        return self.Text

    @property
    def Top(self):
        return self.com_object.Top

    # Lower case aliases for Top
    @property
    def top(self):
        return self.Top

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()


class Comments:

    def __init__(self, comments=None):
        self.com_object= comments

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Add(self, Left=None, Top=None, Author=None, AuthorInitials=None, Text=None):
        arguments = com_arguments([unwrap(a) for a in [Left, Top, Author, AuthorInitials, Text]])
        return Comment(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Left=None, Top=None, Author=None, AuthorInitials=None, Text=None):
        arguments = [Left, Top, Author, AuthorInitials, Text]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Comment(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class ConnectorFormat:

    def __init__(self, connectorformat=None):
        self.com_object= connectorformat

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def BeginConnected(self):
        return self.com_object.BeginConnected

    @BeginConnected.setter
    def BeginConnected(self, value):
        self.com_object.BeginConnected = value

    # Lower case aliases for BeginConnected
    @property
    def beginconnected(self):
        return self.BeginConnected

    @beginconnected.setter
    def beginconnected(self, value):
        self.BeginConnected = value

    @property
    def BeginConnectedShape(self):
        return Shape(self.com_object.BeginConnectedShape)

    # Lower case aliases for BeginConnectedShape
    @property
    def beginconnectedshape(self):
        return self.BeginConnectedShape

    @property
    def BeginConnectionSite(self):
        return self.com_object.BeginConnectionSite

    # Lower case aliases for BeginConnectionSite
    @property
    def beginconnectionsite(self):
        return self.BeginConnectionSite

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def EndConnected(self):
        return self.com_object.EndConnected

    # Lower case aliases for EndConnected
    @property
    def endconnected(self):
        return self.EndConnected

    @property
    def EndConnectedShape(self):
        return Shape(self.com_object.EndConnectedShape)

    # Lower case aliases for EndConnectedShape
    @property
    def endconnectedshape(self):
        return self.EndConnectedShape

    @property
    def EndConnectionSite(self):
        return self.com_object.EndConnectionSite

    # Lower case aliases for EndConnectionSite
    @property
    def endconnectionsite(self):
        return self.EndConnectionSite

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

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

    def BeginConnect(self, ConnectedShape=None, ConnectionSite=None):
        arguments = com_arguments([unwrap(a) for a in [ConnectedShape, ConnectionSite]])
        self.com_object.BeginConnect(*arguments)

    # Lower case alias for BeginConnect
    def beginconnect(self, ConnectedShape=None, ConnectionSite=None):
        arguments = [ConnectedShape, ConnectionSite]
        return self.BeginConnect(*arguments)

    def BeginDisconnect(self):
        self.com_object.BeginDisconnect()

    # Lower case alias for BeginDisconnect
    def begindisconnect(self):
        return self.BeginDisconnect()

    def EndConnect(self, ConnectedShape=None, ConnectionSite=None):
        arguments = com_arguments([unwrap(a) for a in [ConnectedShape, ConnectionSite]])
        self.com_object.EndConnect(*arguments)

    # Lower case alias for EndConnect
    def endconnect(self, ConnectedShape=None, ConnectionSite=None):
        arguments = [ConnectedShape, ConnectionSite]
        return self.EndConnect(*arguments)

    def EndDisconnect(self):
        self.com_object.EndDisconnect()

    # Lower case alias for EndDisconnect
    def enddisconnect(self):
        return self.EndDisconnect()


class CustomerData:

    def __init__(self, customerdata=None):
        self.com_object= customerdata

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return CustomerData(self.com_object.Parent)

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    def Add(self):
        return self.com_object.Add()

    # Lower case alias for Add
    def add(self):
        return self.Add()

    def Delete(self, Id=None):
        arguments = com_arguments([unwrap(a) for a in [Id]])
        self.com_object.Delete(*arguments)

    # Lower case alias for Delete
    def delete(self, Id=None):
        arguments = [Id]
        return self.Delete(*arguments)

    def Item(self, Id=None):
        arguments = com_arguments([unwrap(a) for a in [Id]])
        return self.com_object.Item(*arguments)

    # Lower case alias for Item
    def item(self, Id=None):
        arguments = [Id]
        return self.Item(*arguments)


class CustomLayout:

    def __init__(self, customlayout=None):
        self.com_object= customlayout

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Background(self):
        return ShapeRange(self.com_object.Background)

    # Lower case aliases for Background
    @property
    def background(self):
        return self.Background

    @property
    def CustomerData(self):
        return CustomerData(self.com_object.CustomerData)

    # Lower case aliases for CustomerData
    @property
    def customerdata(self):
        return self.CustomerData

    @property
    def Design(self):
        return Design(self.com_object.Design)

    # Lower case aliases for Design
    @property
    def design(self):
        return self.Design

    @property
    def DisplayMasterShapes(self):
        return self.com_object.DisplayMasterShapes

    @DisplayMasterShapes.setter
    def DisplayMasterShapes(self, value):
        self.com_object.DisplayMasterShapes = value

    # Lower case aliases for DisplayMasterShapes
    @property
    def displaymastershapes(self):
        return self.DisplayMasterShapes

    @displaymastershapes.setter
    def displaymastershapes(self, value):
        self.DisplayMasterShapes = value

    @property
    def FollowMasterBackground(self):
        return self.com_object.FollowMasterBackground

    @FollowMasterBackground.setter
    def FollowMasterBackground(self, value):
        self.com_object.FollowMasterBackground = value

    # Lower case aliases for FollowMasterBackground
    @property
    def followmasterbackground(self):
        return self.FollowMasterBackground

    @followmasterbackground.setter
    def followmasterbackground(self, value):
        self.FollowMasterBackground = value

    @property
    def HeadersFooters(self):
        return HeadersFooters(self.com_object.HeadersFooters)

    # Lower case aliases for HeadersFooters
    @property
    def headersfooters(self):
        return self.HeadersFooters

    @property
    def Height(self):
        return self.com_object.Height

    # Lower case aliases for Height
    @property
    def height(self):
        return self.Height

    @property
    def Hyperlinks(self):
        return Hyperlinks(self.com_object.Hyperlinks)

    # Lower case aliases for Hyperlinks
    @property
    def hyperlinks(self):
        return self.Hyperlinks

    @property
    def Index(self):
        return CustomLayouts(self.com_object.Index)

    # Lower case aliases for Index
    @property
    def index(self):
        return self.Index

    @property
    def MatchingName(self):
        return self.com_object.MatchingName

    @MatchingName.setter
    def MatchingName(self, value):
        self.com_object.MatchingName = value

    # Lower case aliases for MatchingName
    @property
    def matchingname(self):
        return self.MatchingName

    @matchingname.setter
    def matchingname(self, value):
        self.MatchingName = value

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
        return CustomLayout(self.com_object.Parent)

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Preserved(self):
        return self.com_object.Preserved

    @Preserved.setter
    def Preserved(self, value):
        self.com_object.Preserved = value

    # Lower case aliases for Preserved
    @property
    def preserved(self):
        return self.Preserved

    @preserved.setter
    def preserved(self, value):
        self.Preserved = value

    @property
    def Shapes(self):
        return Shapes(self.com_object.Shapes)

    # Lower case aliases for Shapes
    @property
    def shapes(self):
        return self.Shapes

    @property
    def SlideShowTransition(self):
        return SlideShowTransition(self.com_object.SlideShowTransition)

    # Lower case aliases for SlideShowTransition
    @property
    def slideshowtransition(self):
        return self.SlideShowTransition

    @property
    def ThemeColorScheme(self):
        return self.com_object.ThemeColorScheme

    # Lower case aliases for ThemeColorScheme
    @property
    def themecolorscheme(self):
        return self.ThemeColorScheme

    @property
    def TimeLine(self):
        return TimeLine(self.com_object.TimeLine)

    # Lower case aliases for TimeLine
    @property
    def timeline(self):
        return self.TimeLine

    @property
    def Width(self):
        return self.com_object.Width

    # Lower case aliases for Width
    @property
    def width(self):
        return self.Width

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

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Duplicate(self):
        return CustomLayout(self.com_object.Duplicate())

    # Lower case alias for Duplicate
    def duplicate(self):
        return self.Duplicate()

    def MoveTo(self, toPos=None):
        arguments = com_arguments([unwrap(a) for a in [toPos]])
        self.com_object.MoveTo(*arguments)

    # Lower case alias for MoveTo
    def moveto(self, toPos=None):
        arguments = [toPos]
        return self.MoveTo(*arguments)

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class CustomLayouts:

    def __init__(self, customlayouts=None):
        self.com_object= customlayouts

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Add(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return CustomLayout(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Index=None):
        arguments = [Index]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return CustomLayout(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Paste(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return CustomLayout(self.com_object.Paste(*arguments))

    # Lower case alias for Paste
    def paste(self, Index=None):
        arguments = [Index]
        return self.Paste(*arguments)


class DataLabel:

    def __init__(self, datalabel=None):
        self.com_object= datalabel

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def AutoText(self):
        return self.com_object.AutoText

    @AutoText.setter
    def AutoText(self, value):
        self.com_object.AutoText = value

    # Lower case aliases for AutoText
    @property
    def autotext(self):
        return self.AutoText

    @autotext.setter
    def autotext(self, value):
        self.AutoText = value

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

    def Characters(self, Start=None, Length=None):
        arguments = com_arguments([unwrap(a) for a in [Start, Length]])
        if hasattr(self.com_object, "GetCharacters"):
            return ChartCharacters(self.com_object.GetCharacters(*arguments))
        else:
            return ChartCharacters(self.com_object.Characters(*arguments))

    # Lower case aliases for Characters
    def characters(self, Start=None, Length=None):
        arguments = [Start, Length]
        return self.Characters(*arguments)

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

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
    def FormulaLocal(self):
        return self.com_object.FormulaLocal

    @FormulaLocal.setter
    def FormulaLocal(self, value):
        self.com_object.FormulaLocal = value

    # Lower case aliases for FormulaLocal
    @property
    def formulalocal(self):
        return self.FormulaLocal

    @formulalocal.setter
    def formulalocal(self, value):
        self.FormulaLocal = value

    @property
    def FormulaR1C1(self):
        return self.com_object.FormulaR1C1

    @FormulaR1C1.setter
    def FormulaR1C1(self, value):
        self.com_object.FormulaR1C1 = value

    # Lower case aliases for FormulaR1C1
    @property
    def formular1c1(self):
        return self.FormulaR1C1

    @formular1c1.setter
    def formular1c1(self, value):
        self.FormulaR1C1 = value

    @property
    def FormulaR1C1Local(self):
        return self.com_object.FormulaR1C1Local

    @FormulaR1C1Local.setter
    def FormulaR1C1Local(self, value):
        self.com_object.FormulaR1C1Local = value

    # Lower case aliases for FormulaR1C1Local
    @property
    def formular1c1local(self):
        return self.FormulaR1C1Local

    @formular1c1local.setter
    def formular1c1local(self, value):
        self.FormulaR1C1Local = value

    @property
    def Height(self):
        return self.com_object.Height

    # Lower case aliases for Height
    @property
    def height(self):
        return self.Height

    @property
    def HorizontalAlignment(self):
        return self.com_object.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.com_object.HorizontalAlignment = value

    # Lower case aliases for HorizontalAlignment
    @property
    def horizontalalignment(self):
        return self.HorizontalAlignment

    @horizontalalignment.setter
    def horizontalalignment(self, value):
        self.HorizontalAlignment = value

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
    def Name(self):
        return self.com_object.Name

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def NumberFormat(self):
        return self.com_object.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value):
        self.com_object.NumberFormat = value

    # Lower case aliases for NumberFormat
    @property
    def numberformat(self):
        return self.NumberFormat

    @numberformat.setter
    def numberformat(self, value):
        self.NumberFormat = value

    @property
    def NumberFormatLinked(self):
        return self.com_object.NumberFormatLinked

    @NumberFormatLinked.setter
    def NumberFormatLinked(self, value):
        self.com_object.NumberFormatLinked = value

    # Lower case aliases for NumberFormatLinked
    @property
    def numberformatlinked(self):
        return self.NumberFormatLinked

    @numberformatlinked.setter
    def numberformatlinked(self, value):
        self.NumberFormatLinked = value

    @property
    def NumberFormatLocal(self):
        return self.com_object.NumberFormatLocal

    @NumberFormatLocal.setter
    def NumberFormatLocal(self, value):
        self.com_object.NumberFormatLocal = value

    # Lower case aliases for NumberFormatLocal
    @property
    def numberformatlocal(self):
        return self.NumberFormatLocal

    @numberformatlocal.setter
    def numberformatlocal(self, value):
        self.NumberFormatLocal = value

    @property
    def Orientation(self):
        return self.com_object.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.com_object.Orientation = value

    # Lower case aliases for Orientation
    @property
    def orientation(self):
        return self.Orientation

    @orientation.setter
    def orientation(self, value):
        self.Orientation = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return XlDataLabelPosition(self.com_object.Position)

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
    def ReadingOrder(self):
        return XlReadingOrder(self.com_object.ReadingOrder)

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.com_object.ReadingOrder = value

    # Lower case aliases for ReadingOrder
    @property
    def readingorder(self):
        return self.ReadingOrder

    @readingorder.setter
    def readingorder(self, value):
        self.ReadingOrder = value

    @property
    def Separator(self):
        return self.com_object.Separator

    @Separator.setter
    def Separator(self, value):
        self.com_object.Separator = value

    # Lower case aliases for Separator
    @property
    def separator(self):
        return self.Separator

    @separator.setter
    def separator(self, value):
        self.Separator = value

    @property
    def Shadow(self):
        return self.com_object.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.com_object.Shadow = value

    # Lower case aliases for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @shadow.setter
    def shadow(self, value):
        self.Shadow = value

    @property
    def ShowBubbleSize(self):
        return self.com_object.ShowBubbleSize

    @ShowBubbleSize.setter
    def ShowBubbleSize(self, value):
        self.com_object.ShowBubbleSize = value

    # Lower case aliases for ShowBubbleSize
    @property
    def showbubblesize(self):
        return self.ShowBubbleSize

    @showbubblesize.setter
    def showbubblesize(self, value):
        self.ShowBubbleSize = value

    @property
    def ShowCategoryName(self):
        return self.com_object.ShowCategoryName

    @ShowCategoryName.setter
    def ShowCategoryName(self, value):
        self.com_object.ShowCategoryName = value

    # Lower case aliases for ShowCategoryName
    @property
    def showcategoryname(self):
        return self.ShowCategoryName

    @showcategoryname.setter
    def showcategoryname(self, value):
        self.ShowCategoryName = value

    @property
    def ShowLegendKey(self):
        return self.com_object.ShowLegendKey

    @ShowLegendKey.setter
    def ShowLegendKey(self, value):
        self.com_object.ShowLegendKey = value

    # Lower case aliases for ShowLegendKey
    @property
    def showlegendkey(self):
        return self.ShowLegendKey

    @showlegendkey.setter
    def showlegendkey(self, value):
        self.ShowLegendKey = value

    @property
    def ShowPercentage(self):
        return self.com_object.ShowPercentage

    @ShowPercentage.setter
    def ShowPercentage(self, value):
        self.com_object.ShowPercentage = value

    # Lower case aliases for ShowPercentage
    @property
    def showpercentage(self):
        return self.ShowPercentage

    @showpercentage.setter
    def showpercentage(self, value):
        self.ShowPercentage = value

    @property
    def ShowSeriesName(self):
        return self.com_object.ShowSeriesName

    @ShowSeriesName.setter
    def ShowSeriesName(self, value):
        self.com_object.ShowSeriesName = value

    # Lower case aliases for ShowSeriesName
    @property
    def showseriesname(self):
        return self.ShowSeriesName

    @showseriesname.setter
    def showseriesname(self, value):
        self.ShowSeriesName = value

    @property
    def ShowValue(self):
        return self.com_object.ShowValue

    @ShowValue.setter
    def ShowValue(self, value):
        self.com_object.ShowValue = value

    # Lower case aliases for ShowValue
    @property
    def showvalue(self):
        return self.ShowValue

    @showvalue.setter
    def showvalue(self, value):
        self.ShowValue = value

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
    def VerticalAlignment(self):
        return self.com_object.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.com_object.VerticalAlignment = value

    # Lower case aliases for VerticalAlignment
    @property
    def verticalalignment(self):
        return self.VerticalAlignment

    @verticalalignment.setter
    def verticalalignment(self, value):
        self.VerticalAlignment = value

    @property
    def Width(self):
        return self.com_object.Width

    # Lower case aliases for Width
    @property
    def width(self):
        return self.Width

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class DataLabels:

    def __init__(self, datalabels=None):
        self.com_object= datalabels

    def __call__(self, item):
        return DataLabel(self.com_object(item))

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def AutoText(self):
        return self.com_object.AutoText

    @AutoText.setter
    def AutoText(self, value):
        self.com_object.AutoText = value

    # Lower case aliases for AutoText
    @property
    def autotext(self):
        return self.AutoText

    @autotext.setter
    def autotext(self, value):
        self.AutoText = value

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

    @property
    def HorizontalAlignment(self):
        return self.com_object.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.com_object.HorizontalAlignment = value

    # Lower case aliases for HorizontalAlignment
    @property
    def horizontalalignment(self):
        return self.HorizontalAlignment

    @horizontalalignment.setter
    def horizontalalignment(self, value):
        self.HorizontalAlignment = value

    @property
    def Name(self):
        return self.com_object.Name

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def NumberFormat(self):
        return self.com_object.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value):
        self.com_object.NumberFormat = value

    # Lower case aliases for NumberFormat
    @property
    def numberformat(self):
        return self.NumberFormat

    @numberformat.setter
    def numberformat(self, value):
        self.NumberFormat = value

    @property
    def NumberFormatLinked(self):
        return self.com_object.NumberFormatLinked

    @NumberFormatLinked.setter
    def NumberFormatLinked(self, value):
        self.com_object.NumberFormatLinked = value

    # Lower case aliases for NumberFormatLinked
    @property
    def numberformatlinked(self):
        return self.NumberFormatLinked

    @numberformatlinked.setter
    def numberformatlinked(self, value):
        self.NumberFormatLinked = value

    @property
    def NumberFormatLocal(self):
        return self.com_object.NumberFormatLocal

    @NumberFormatLocal.setter
    def NumberFormatLocal(self, value):
        self.com_object.NumberFormatLocal = value

    # Lower case aliases for NumberFormatLocal
    @property
    def numberformatlocal(self):
        return self.NumberFormatLocal

    @numberformatlocal.setter
    def numberformatlocal(self, value):
        self.NumberFormatLocal = value

    @property
    def Orientation(self):
        return self.com_object.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.com_object.Orientation = value

    # Lower case aliases for Orientation
    @property
    def orientation(self):
        return self.Orientation

    @orientation.setter
    def orientation(self, value):
        self.Orientation = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ReadingOrder(self):
        return XlReadingOrder(self.com_object.ReadingOrder)

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.com_object.ReadingOrder = value

    # Lower case aliases for ReadingOrder
    @property
    def readingorder(self):
        return self.ReadingOrder

    @readingorder.setter
    def readingorder(self, value):
        self.ReadingOrder = value

    @property
    def Separator(self):
        return self.com_object.Separator

    @Separator.setter
    def Separator(self, value):
        self.com_object.Separator = value

    # Lower case aliases for Separator
    @property
    def separator(self):
        return self.Separator

    @separator.setter
    def separator(self, value):
        self.Separator = value

    @property
    def Shadow(self):
        return self.com_object.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.com_object.Shadow = value

    # Lower case aliases for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @shadow.setter
    def shadow(self, value):
        self.Shadow = value

    @property
    def ShowBubbleSize(self):
        return self.com_object.ShowBubbleSize

    @ShowBubbleSize.setter
    def ShowBubbleSize(self, value):
        self.com_object.ShowBubbleSize = value

    # Lower case aliases for ShowBubbleSize
    @property
    def showbubblesize(self):
        return self.ShowBubbleSize

    @showbubblesize.setter
    def showbubblesize(self, value):
        self.ShowBubbleSize = value

    @property
    def ShowCategoryName(self):
        return self.com_object.ShowCategoryName

    @ShowCategoryName.setter
    def ShowCategoryName(self, value):
        self.com_object.ShowCategoryName = value

    # Lower case aliases for ShowCategoryName
    @property
    def showcategoryname(self):
        return self.ShowCategoryName

    @showcategoryname.setter
    def showcategoryname(self, value):
        self.ShowCategoryName = value

    @property
    def ShowLegendKey(self):
        return self.com_object.ShowLegendKey

    @ShowLegendKey.setter
    def ShowLegendKey(self, value):
        self.com_object.ShowLegendKey = value

    # Lower case aliases for ShowLegendKey
    @property
    def showlegendkey(self):
        return self.ShowLegendKey

    @showlegendkey.setter
    def showlegendkey(self, value):
        self.ShowLegendKey = value

    @property
    def ShowPercentage(self):
        return self.com_object.ShowPercentage

    @ShowPercentage.setter
    def ShowPercentage(self, value):
        self.com_object.ShowPercentage = value

    # Lower case aliases for ShowPercentage
    @property
    def showpercentage(self):
        return self.ShowPercentage

    @showpercentage.setter
    def showpercentage(self, value):
        self.ShowPercentage = value

    @property
    def ShowSeriesName(self):
        return self.com_object.ShowSeriesName

    @ShowSeriesName.setter
    def ShowSeriesName(self, value):
        self.com_object.ShowSeriesName = value

    # Lower case aliases for ShowSeriesName
    @property
    def showseriesname(self):
        return self.ShowSeriesName

    @showseriesname.setter
    def showseriesname(self, value):
        self.ShowSeriesName = value

    @property
    def ShowValue(self):
        return self.com_object.ShowValue

    @ShowValue.setter
    def ShowValue(self, value):
        self.com_object.ShowValue = value

    # Lower case aliases for ShowValue
    @property
    def showvalue(self):
        return self.ShowValue

    @showvalue.setter
    def showvalue(self, value):
        self.ShowValue = value

    @property
    def VerticalAlignment(self):
        return self.com_object.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.com_object.VerticalAlignment = value

    # Lower case aliases for VerticalAlignment
    @property
    def verticalalignment(self):
        return self.VerticalAlignment

    @verticalalignment.setter
    def verticalalignment(self, value):
        self.VerticalAlignment = value

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return DataLabel(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class DataTable:

    def __init__(self, datatable=None):
        self.com_object= datatable

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Border(self):
        return ChartBorder(self.com_object.Border)

    # Lower case aliases for Border
    @property
    def border(self):
        return self.Border

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Font(self):
        return ChartFont(self.com_object.Font)

    # Lower case aliases for Font
    @property
    def font(self):
        return self.Font

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

    @property
    def HasBorderHorizontal(self):
        return self.com_object.HasBorderHorizontal

    @HasBorderHorizontal.setter
    def HasBorderHorizontal(self, value):
        self.com_object.HasBorderHorizontal = value

    # Lower case aliases for HasBorderHorizontal
    @property
    def hasborderhorizontal(self):
        return self.HasBorderHorizontal

    @hasborderhorizontal.setter
    def hasborderhorizontal(self, value):
        self.HasBorderHorizontal = value

    @property
    def HasBorderOutline(self):
        return self.com_object.HasBorderOutline

    @HasBorderOutline.setter
    def HasBorderOutline(self, value):
        self.com_object.HasBorderOutline = value

    # Lower case aliases for HasBorderOutline
    @property
    def hasborderoutline(self):
        return self.HasBorderOutline

    @hasborderoutline.setter
    def hasborderoutline(self, value):
        self.HasBorderOutline = value

    @property
    def HasBorderVertical(self):
        return self.com_object.HasBorderVertical

    @HasBorderVertical.setter
    def HasBorderVertical(self, value):
        self.com_object.HasBorderVertical = value

    # Lower case aliases for HasBorderVertical
    @property
    def hasbordervertical(self):
        return self.HasBorderVertical

    @hasbordervertical.setter
    def hasbordervertical(self, value):
        self.HasBorderVertical = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ShowLegendKey(self):
        return self.com_object.ShowLegendKey

    @ShowLegendKey.setter
    def ShowLegendKey(self, value):
        self.com_object.ShowLegendKey = value

    # Lower case aliases for ShowLegendKey
    @property
    def showlegendkey(self):
        return self.ShowLegendKey

    @showlegendkey.setter
    def showlegendkey(self, value):
        self.ShowLegendKey = value

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class Design:

    def __init__(self, design=None):
        self.com_object= design

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Index(self):
        return self.com_object.Index

    # Lower case aliases for Index
    @property
    def index(self):
        return self.Index

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
    def Preserved(self):
        return self.com_object.Preserved

    @Preserved.setter
    def Preserved(self, value):
        self.com_object.Preserved = value

    # Lower case aliases for Preserved
    @property
    def preserved(self):
        return self.Preserved

    @preserved.setter
    def preserved(self, value):
        self.Preserved = value

    @property
    def SlideMaster(self):
        return Master(self.com_object.SlideMaster)

    # Lower case aliases for SlideMaster
    @property
    def slidemaster(self):
        return self.SlideMaster

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def MoveTo(self, toPos=None):
        arguments = com_arguments([unwrap(a) for a in [toPos]])
        self.com_object.MoveTo(*arguments)

    # Lower case alias for MoveTo
    def moveto(self, toPos=None):
        arguments = [toPos]
        return self.MoveTo(*arguments)


class Designs:

    def __init__(self, designs=None):
        self.com_object= designs

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Add(self, designName=None, Index=None):
        arguments = com_arguments([unwrap(a) for a in [designName, Index]])
        return Design(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, designName=None, Index=None):
        arguments = [designName, Index]
        return self.Add(*arguments)

    def Clone(self, pOriginal=None, Index=None):
        arguments = com_arguments([unwrap(a) for a in [pOriginal, Index]])
        return Design(self.com_object.Clone(*arguments))

    # Lower case alias for Clone
    def clone(self, pOriginal=None, Index=None):
        arguments = [pOriginal, Index]
        return self.Clone(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Design(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Load(self, TemplateName=None, Index=None):
        arguments = com_arguments([unwrap(a) for a in [TemplateName, Index]])
        return Design(self.com_object.Load(*arguments))

    # Lower case alias for Load
    def load(self, TemplateName=None, Index=None):
        arguments = [TemplateName, Index]
        return self.Load(*arguments)


class DisplayUnitLabel:

    def __init__(self, displayunitlabel=None):
        self.com_object= displayunitlabel

    @property
    def Application(self):
        return self.com_object.Application

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

    def Characters(self, Start=None, Length=None):
        arguments = com_arguments([unwrap(a) for a in [Start, Length]])
        if hasattr(self.com_object, "GetCharacters"):
            return ChartCharacters(self.com_object.GetCharacters(*arguments))
        else:
            return ChartCharacters(self.com_object.Characters(*arguments))

    # Lower case aliases for Characters
    def characters(self, Start=None, Length=None):
        arguments = [Start, Length]
        return self.Characters(*arguments)

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

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
    def FormulaLocal(self):
        return self.com_object.FormulaLocal

    @FormulaLocal.setter
    def FormulaLocal(self, value):
        self.com_object.FormulaLocal = value

    # Lower case aliases for FormulaLocal
    @property
    def formulalocal(self):
        return self.FormulaLocal

    @formulalocal.setter
    def formulalocal(self, value):
        self.FormulaLocal = value

    @property
    def FormulaR1C1(self):
        return self.com_object.FormulaR1C1

    @FormulaR1C1.setter
    def FormulaR1C1(self, value):
        self.com_object.FormulaR1C1 = value

    # Lower case aliases for FormulaR1C1
    @property
    def formular1c1(self):
        return self.FormulaR1C1

    @formular1c1.setter
    def formular1c1(self, value):
        self.FormulaR1C1 = value

    @property
    def FormulaR1C1Local(self):
        return self.com_object.FormulaR1C1Local

    @FormulaR1C1Local.setter
    def FormulaR1C1Local(self, value):
        self.com_object.FormulaR1C1Local = value

    # Lower case aliases for FormulaR1C1Local
    @property
    def formular1c1local(self):
        return self.FormulaR1C1Local

    @formular1c1local.setter
    def formular1c1local(self, value):
        self.FormulaR1C1Local = value

    @property
    def Height(self):
        return self.com_object.Height

    # Lower case aliases for Height
    @property
    def height(self):
        return self.Height

    @property
    def HorizontalAlignment(self):
        return self.com_object.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.com_object.HorizontalAlignment = value

    # Lower case aliases for HorizontalAlignment
    @property
    def horizontalalignment(self):
        return self.HorizontalAlignment

    @horizontalalignment.setter
    def horizontalalignment(self, value):
        self.HorizontalAlignment = value

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
    def Name(self):
        return self.com_object.Name

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def Orientation(self):
        return self.com_object.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.com_object.Orientation = value

    # Lower case aliases for Orientation
    @property
    def orientation(self):
        return self.Orientation

    @orientation.setter
    def orientation(self, value):
        self.Orientation = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return XlChartElementPosition(self.com_object.Position)

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
    def ReadingOrder(self):
        return XlReadingOrder(self.com_object.ReadingOrder)

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.com_object.ReadingOrder = value

    # Lower case aliases for ReadingOrder
    @property
    def readingorder(self):
        return self.ReadingOrder

    @readingorder.setter
    def readingorder(self, value):
        self.ReadingOrder = value

    @property
    def Shadow(self):
        return self.com_object.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.com_object.Shadow = value

    # Lower case aliases for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @shadow.setter
    def shadow(self, value):
        self.Shadow = value

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
    def VerticalAlignment(self):
        return self.com_object.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.com_object.VerticalAlignment = value

    # Lower case aliases for VerticalAlignment
    @property
    def verticalalignment(self):
        return self.VerticalAlignment

    @verticalalignment.setter
    def verticalalignment(self, value):
        self.VerticalAlignment = value

    @property
    def Width(self):
        return self.com_object.Width

    # Lower case aliases for Width
    @property
    def width(self):
        return self.Width

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class DocumentWindow:

    def __init__(self, documentwindow=None):
        self.com_object= documentwindow

    @property
    def Active(self):
        return self.com_object.Active

    # Lower case aliases for Active
    @property
    def active(self):
        return self.Active

    @property
    def ActivePane(self):
        return Pane(self.com_object.ActivePane)

    # Lower case aliases for ActivePane
    @property
    def activepane(self):
        return self.ActivePane

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def BlackAndWhite(self):
        return self.com_object.BlackAndWhite

    @BlackAndWhite.setter
    def BlackAndWhite(self, value):
        self.com_object.BlackAndWhite = value

    # Lower case aliases for BlackAndWhite
    @property
    def blackandwhite(self):
        return self.BlackAndWhite

    @blackandwhite.setter
    def blackandwhite(self, value):
        self.BlackAndWhite = value

    @property
    def Caption(self):
        return self.com_object.Caption

    # Lower case aliases for Caption
    @property
    def caption(self):
        return self.Caption

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
    def Presentation(self):
        return Presentation(self.com_object.Presentation)

    # Lower case aliases for Presentation
    @property
    def presentation(self):
        return self.Presentation

    @property
    def Selection(self):
        return Selection(self.com_object.Selection)

    # Lower case aliases for Selection
    @property
    def selection(self):
        return self.Selection

    @property
    def SplitHorizontal(self):
        return self.com_object.SplitHorizontal

    @SplitHorizontal.setter
    def SplitHorizontal(self, value):
        self.com_object.SplitHorizontal = value

    # Lower case aliases for SplitHorizontal
    @property
    def splithorizontal(self):
        return self.SplitHorizontal

    @splithorizontal.setter
    def splithorizontal(self, value):
        self.SplitHorizontal = value

    @property
    def SplitVertical(self):
        return self.com_object.SplitVertical

    @SplitVertical.setter
    def SplitVertical(self, value):
        self.com_object.SplitVertical = value

    # Lower case aliases for SplitVertical
    @property
    def splitvertical(self):
        return self.SplitVertical

    @splitvertical.setter
    def splitvertical(self, value):
        self.SplitVertical = value

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
    def View(self):
        return View(self.com_object.View)

    # Lower case aliases for View
    @property
    def view(self):
        return self.View

    @property
    def ViewType(self):
        return self.com_object.ViewType

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
        return self.com_object.WindowState

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

    def Close(self):
        self.com_object.Close()

    # Lower case alias for Close
    def close(self):
        return self.Close()

    def ExpandSection(self, sectionIndex=None, Expand=None):
        arguments = com_arguments([unwrap(a) for a in [sectionIndex, Expand]])
        self.com_object.ExpandSection(*arguments)

    # Lower case alias for ExpandSection
    def expandsection(self, sectionIndex=None, Expand=None):
        arguments = [sectionIndex, Expand]
        return self.ExpandSection(*arguments)

    def FitToPage(self):
        self.com_object.FitToPage()

    # Lower case alias for FitToPage
    def fittopage(self):
        return self.FitToPage()

    def IsSectionExpanded(self, sectionIndex=None):
        arguments = com_arguments([unwrap(a) for a in [sectionIndex]])
        return self.com_object.IsSectionExpanded(*arguments)

    # Lower case alias for IsSectionExpanded
    def issectionexpanded(self, sectionIndex=None):
        arguments = [sectionIndex]
        return self.IsSectionExpanded(*arguments)

    def LargeScroll(self, Down=None, Up=None, ToRight=None, ToLeft=None):
        arguments = com_arguments([unwrap(a) for a in [Down, Up, ToRight, ToLeft]])
        self.com_object.LargeScroll(*arguments)

    # Lower case alias for LargeScroll
    def largescroll(self, Down=None, Up=None, ToRight=None, ToLeft=None):
        arguments = [Down, Up, ToRight, ToLeft]
        return self.LargeScroll(*arguments)

    def NewWindow(self):
        return DocumentWindow(self.com_object.NewWindow())

    # Lower case alias for NewWindow
    def newwindow(self):
        return self.NewWindow()

    def PointsToScreenPixelsX(self, Points=None):
        arguments = com_arguments([unwrap(a) for a in [Points]])
        return self.com_object.PointsToScreenPixelsX(*arguments)

    # Lower case alias for PointsToScreenPixelsX
    def pointstoscreenpixelsx(self, Points=None):
        arguments = [Points]
        return self.PointsToScreenPixelsX(*arguments)

    def PointsToScreenPixelsY(self, Points=None):
        arguments = com_arguments([unwrap(a) for a in [Points]])
        return self.com_object.PointsToScreenPixelsY(*arguments)

    # Lower case alias for PointsToScreenPixelsY
    def pointstoscreenpixelsy(self, Points=None):
        arguments = [Points]
        return self.PointsToScreenPixelsY(*arguments)

    def RangeFromPoint(self, x=None, y=None):
        arguments = com_arguments([unwrap(a) for a in [x, y]])
        self.com_object.RangeFromPoint(*arguments)

    # Lower case alias for RangeFromPoint
    def rangefrompoint(self, x=None, y=None):
        arguments = [x, y]
        return self.RangeFromPoint(*arguments)

    def ScrollIntoView(self, Left=None, Top=None, Width=None, Height=None, Start=None):
        arguments = com_arguments([unwrap(a) for a in [Left, Top, Width, Height, Start]])
        self.com_object.ScrollIntoView(*arguments)

    # Lower case alias for ScrollIntoView
    def scrollintoview(self, Left=None, Top=None, Width=None, Height=None, Start=None):
        arguments = [Left, Top, Width, Height, Start]
        return self.ScrollIntoView(*arguments)

    def SmallScroll(self, Down=None, Up=None, ToRight=None, ToLeft=None):
        arguments = com_arguments([unwrap(a) for a in [Down, Up, ToRight, ToLeft]])
        self.com_object.SmallScroll(*arguments)

    # Lower case alias for SmallScroll
    def smallscroll(self, Down=None, Up=None, ToRight=None, ToLeft=None):
        arguments = [Down, Up, ToRight, ToLeft]
        return self.SmallScroll(*arguments)


class DocumentWindows:

    def __init__(self, documentwindows=None):
        self.com_object= documentwindows

    def __call__(self, item):
        return DocumentWindow(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Arrange(self, arrangeStyle=None):
        arguments = com_arguments([unwrap(a) for a in [arrangeStyle]])
        return self.com_object.Arrange(*arguments)

    # Lower case alias for Arrange
    def arrange(self, arrangeStyle=None):
        arguments = [arrangeStyle]
        return self.Arrange(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return DocumentWindow(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class DownBars:

    def __init__(self, downbars=None):
        self.com_object= downbars

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

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

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class DropLines:

    def __init__(self, droplines=None):
        self.com_object= droplines

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Border(self):
        return ChartBorder(self.com_object.Border)

    # Lower case aliases for Border
    @property
    def border(self):
        return self.Border

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

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

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class Effect:

    def __init__(self, effect=None):
        self.com_object= effect

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Behaviors(self):
        return AnimationBehaviors(self.com_object.Behaviors)

    # Lower case aliases for Behaviors
    @property
    def behaviors(self):
        return self.Behaviors

    @property
    def DisplayName(self):
        return self.com_object.DisplayName

    # Lower case aliases for DisplayName
    @property
    def displayname(self):
        return self.DisplayName

    @property
    def EffectInformation(self):
        return EffectInformation(self.com_object.EffectInformation)

    # Lower case aliases for EffectInformation
    @property
    def effectinformation(self):
        return self.EffectInformation

    @property
    def EffectParameters(self):
        return EffectParameters(self.com_object.EffectParameters)

    # Lower case aliases for EffectParameters
    @property
    def effectparameters(self):
        return self.EffectParameters

    @property
    def EffectType(self):
        return self.com_object.EffectType

    @EffectType.setter
    def EffectType(self, value):
        self.com_object.EffectType = value

    # Lower case aliases for EffectType
    @property
    def effecttype(self):
        return self.EffectType

    @effecttype.setter
    def effecttype(self, value):
        self.EffectType = value

    @property
    def Exit(self):
        return self.com_object.Exit

    @Exit.setter
    def Exit(self, value):
        self.com_object.Exit = value

    # Lower case aliases for Exit
    @property
    def exit(self):
        return self.Exit

    @exit.setter
    def exit(self, value):
        self.Exit = value

    @property
    def Index(self):
        return self.com_object.Index

    # Lower case aliases for Index
    @property
    def index(self):
        return self.Index

    @property
    def Paragraph(self):
        return self.com_object.Paragraph

    @Paragraph.setter
    def Paragraph(self, value):
        self.com_object.Paragraph = value

    # Lower case aliases for Paragraph
    @property
    def paragraph(self):
        return self.Paragraph

    @paragraph.setter
    def paragraph(self, value):
        self.Paragraph = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Shape(self):
        return Shape(self.com_object.Shape)

    # Lower case aliases for Shape
    @property
    def shape(self):
        return self.Shape

    @property
    def TextRangeLength(self):
        return self.com_object.TextRangeLength

    @TextRangeLength.setter
    def TextRangeLength(self, value):
        self.com_object.TextRangeLength = value

    # Lower case aliases for TextRangeLength
    @property
    def textrangelength(self):
        return self.TextRangeLength

    @textrangelength.setter
    def textrangelength(self, value):
        self.TextRangeLength = value

    @property
    def TextRangeStart(self):
        return self.com_object.TextRangeStart

    @TextRangeStart.setter
    def TextRangeStart(self, value):
        self.com_object.TextRangeStart = value

    # Lower case aliases for TextRangeStart
    @property
    def textrangestart(self):
        return self.TextRangeStart

    @textrangestart.setter
    def textrangestart(self, value):
        self.TextRangeStart = value

    @property
    def Timing(self):
        return Timing(self.com_object.Timing)

    # Lower case aliases for Timing
    @property
    def timing(self):
        return self.Timing

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def MoveAfter(self, Effect=None):
        arguments = com_arguments([unwrap(a) for a in [Effect]])
        self.com_object.MoveAfter(*arguments)

    # Lower case alias for MoveAfter
    def moveafter(self, Effect=None):
        arguments = [Effect]
        return self.MoveAfter(*arguments)

    def MoveBefore(self, Effect=None):
        arguments = com_arguments([unwrap(a) for a in [Effect]])
        self.com_object.MoveBefore(*arguments)

    # Lower case alias for MoveBefore
    def movebefore(self, Effect=None):
        arguments = [Effect]
        return self.MoveBefore(*arguments)

    def MoveTo(self, toPos=None):
        arguments = com_arguments([unwrap(a) for a in [toPos]])
        self.com_object.MoveTo(*arguments)

    # Lower case alias for MoveTo
    def moveto(self, toPos=None):
        arguments = [toPos]
        return self.MoveTo(*arguments)


class EffectInformation:

    def __init__(self, effectinformation=None):
        self.com_object= effectinformation

    @property
    def AfterEffect(self):
        return PpAfterEffect(self.com_object.AfterEffect)

    # Lower case aliases for AfterEffect
    @property
    def aftereffect(self):
        return self.AfterEffect

    @property
    def AnimateBackground(self):
        return self.com_object.AnimateBackground

    # Lower case aliases for AnimateBackground
    @property
    def animatebackground(self):
        return self.AnimateBackground

    @property
    def AnimateTextInReverse(self):
        return self.com_object.AnimateTextInReverse

    @AnimateTextInReverse.setter
    def AnimateTextInReverse(self, value):
        self.com_object.AnimateTextInReverse = value

    # Lower case aliases for AnimateTextInReverse
    @property
    def animatetextinreverse(self):
        return self.AnimateTextInReverse

    @animatetextinreverse.setter
    def animatetextinreverse(self, value):
        self.AnimateTextInReverse = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def BuildByLevelEffect(self):
        return self.com_object.BuildByLevelEffect

    # Lower case aliases for BuildByLevelEffect
    @property
    def buildbyleveleffect(self):
        return self.BuildByLevelEffect

    @property
    def Dim(self):
        return ColorFormat(self.com_object.Dim)

    # Lower case aliases for Dim
    @property
    def dim(self):
        return self.Dim

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PlaySettings(self):
        return PlaySettings(self.com_object.PlaySettings)

    # Lower case aliases for PlaySettings
    @property
    def playsettings(self):
        return self.PlaySettings

    @property
    def SoundEffect(self):
        return SoundEffect(self.com_object.SoundEffect)

    # Lower case aliases for SoundEffect
    @property
    def soundeffect(self):
        return self.SoundEffect

    @property
    def TextUnitEffect(self):
        return self.com_object.TextUnitEffect

    # Lower case aliases for TextUnitEffect
    @property
    def textuniteffect(self):
        return self.TextUnitEffect


class EffectParameters:

    def __init__(self, effectparameters=None):
        self.com_object= effectparameters

    @property
    def Amount(self):
        return self.com_object.Amount

    @Amount.setter
    def Amount(self, value):
        self.com_object.Amount = value

    # Lower case aliases for Amount
    @property
    def amount(self):
        return self.Amount

    @amount.setter
    def amount(self, value):
        self.Amount = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Color2(self):
        return ColorFormat(self.com_object.Color2)

    # Lower case aliases for Color2
    @property
    def color2(self):
        return self.Color2

    @property
    def Direction(self):
        return self.com_object.Direction

    @Direction.setter
    def Direction(self, value):
        self.com_object.Direction = value

    # Lower case aliases for Direction
    @property
    def direction(self):
        return self.Direction

    @direction.setter
    def direction(self, value):
        self.Direction = value

    @property
    def FontName(self):
        return self.com_object.FontName

    @FontName.setter
    def FontName(self, value):
        self.com_object.FontName = value

    # Lower case aliases for FontName
    @property
    def fontname(self):
        return self.FontName

    @fontname.setter
    def fontname(self, value):
        self.FontName = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Relative(self):
        return self.com_object.Relative

    @Relative.setter
    def Relative(self, value):
        self.com_object.Relative = value

    # Lower case aliases for Relative
    @property
    def relative(self):
        return self.Relative

    @relative.setter
    def relative(self, value):
        self.Relative = value

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


class ErrorBars:

    def __init__(self, errorbars=None):
        self.com_object= errorbars

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Border(self):
        return ChartBorder(self.com_object.Border)

    # Lower case aliases for Border
    @property
    def border(self):
        return self.Border

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def EndStyle(self):
        return self.com_object.EndStyle

    @EndStyle.setter
    def EndStyle(self, value):
        self.com_object.EndStyle = value

    # Lower case aliases for EndStyle
    @property
    def endstyle(self):
        return self.EndStyle

    @endstyle.setter
    def endstyle(self, value):
        self.EndStyle = value

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

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

    def ClearFormats(self):
        self.com_object.ClearFormats()

    # Lower case alias for ClearFormats
    def clearformats(self):
        return self.ClearFormats()

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class ExtraColors:

    def __init__(self, extracolors=None):
        self.com_object= extracolors

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Add(self, Type=None):
        arguments = com_arguments([unwrap(a) for a in [Type]])
        self.com_object.Add(*arguments)

    # Lower case alias for Add
    def add(self, Type=None):
        arguments = [Type]
        return self.Add(*arguments)

    def Clear(self):
        self.com_object.Clear()

    # Lower case alias for Clear
    def clear(self):
        return self.Clear()

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return self.com_object.Item(*arguments)

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class FileConverter:

    def __init__(self, fileconverter=None):
        self.com_object= fileconverter

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def CanOpen(self):
        return self.com_object.CanOpen

    # Lower case aliases for CanOpen
    @property
    def canopen(self):
        return self.CanOpen

    @property
    def CanSave(self):
        return self.com_object.CanSave

    # Lower case aliases for CanSave
    @property
    def cansave(self):
        return self.CanSave

    @property
    def ClassName(self):
        return self.com_object.ClassName

    # Lower case aliases for ClassName
    @property
    def classname(self):
        return self.ClassName

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Extensions(self):
        return FileConverter(self.com_object.Extensions)

    # Lower case aliases for Extensions
    @property
    def extensions(self):
        return self.Extensions

    @property
    def FormatName(self):
        return self.com_object.FormatName

    # Lower case aliases for FormatName
    @property
    def formatname(self):
        return self.FormatName

    @property
    def Name(self):
        return self.com_object.Name

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def OpenFormat(self):
        return self.com_object.OpenFormat

    # Lower case aliases for OpenFormat
    @property
    def openformat(self):
        return self.OpenFormat

    @property
    def Parent(self):
        return FileConverter(self.com_object.Parent)

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Path(self):
        return self.com_object.Path

    # Lower case aliases for Path
    @property
    def path(self):
        return self.Path

    @property
    def SaveFormat(self):
        return self.com_object.SaveFormat

    # Lower case aliases for SaveFormat
    @property
    def saveformat(self):
        return self.SaveFormat


class FileConverters:

    def __init__(self, fileconverters=None):
        self.com_object= fileconverters

    def __call__(self, item):
        return FileConverter(self.com_object(item))

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return FileConverter(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class FillFormat:

    def __init__(self, fillformat=None):
        self.com_object= fillformat

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def BackColor(self):
        return ColorFormat(self.com_object.BackColor)

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
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def ForeColor(self):
        return ColorFormat(self.com_object.ForeColor)

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
    def GradientAngle(self):
        return self.com_object.GradientAngle

    @GradientAngle.setter
    def GradientAngle(self, value):
        self.com_object.GradientAngle = value

    # Lower case aliases for GradientAngle
    @property
    def gradientangle(self):
        return self.GradientAngle

    @gradientangle.setter
    def gradientangle(self, value):
        self.GradientAngle = value

    @property
    def GradientColorType(self):
        return self.com_object.GradientColorType

    # Lower case aliases for GradientColorType
    @property
    def gradientcolortype(self):
        return self.GradientColorType

    @property
    def GradientDegree(self):
        return self.com_object.GradientDegree

    # Lower case aliases for GradientDegree
    @property
    def gradientdegree(self):
        return self.GradientDegree

    @property
    def GradientStops(self):
        return self.com_object.GradientStops

    # Lower case aliases for GradientStops
    @property
    def gradientstops(self):
        return self.GradientStops

    @property
    def GradientStyle(self):
        return self.com_object.GradientStyle

    # Lower case aliases for GradientStyle
    @property
    def gradientstyle(self):
        return self.GradientStyle

    @property
    def GradientVariant(self):
        return self.com_object.GradientVariant

    # Lower case aliases for GradientVariant
    @property
    def gradientvariant(self):
        return self.GradientVariant

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Pattern(self):
        return self.com_object.Pattern

    # Lower case aliases for Pattern
    @property
    def pattern(self):
        return self.Pattern

    @property
    def PictureEffects(self):
        return self.com_object.PictureEffects

    # Lower case aliases for PictureEffects
    @property
    def pictureeffects(self):
        return self.PictureEffects

    @property
    def PresetGradientType(self):
        return self.com_object.PresetGradientType

    # Lower case aliases for PresetGradientType
    @property
    def presetgradienttype(self):
        return self.PresetGradientType

    @property
    def PresetTexture(self):
        return self.com_object.PresetTexture

    # Lower case aliases for PresetTexture
    @property
    def presettexture(self):
        return self.PresetTexture

    @property
    def RotateWithObject(self):
        return self.com_object.RotateWithObject

    @RotateWithObject.setter
    def RotateWithObject(self, value):
        self.com_object.RotateWithObject = value

    # Lower case aliases for RotateWithObject
    @property
    def rotatewithobject(self):
        return self.RotateWithObject

    @rotatewithobject.setter
    def rotatewithobject(self, value):
        self.RotateWithObject = value

    @property
    def TextureAlignment(self):
        return self.com_object.TextureAlignment

    @TextureAlignment.setter
    def TextureAlignment(self, value):
        self.com_object.TextureAlignment = value

    # Lower case aliases for TextureAlignment
    @property
    def texturealignment(self):
        return self.TextureAlignment

    @texturealignment.setter
    def texturealignment(self, value):
        self.TextureAlignment = value

    @property
    def TextureHorizontalScale(self):
        return self.com_object.TextureHorizontalScale

    @TextureHorizontalScale.setter
    def TextureHorizontalScale(self, value):
        self.com_object.TextureHorizontalScale = value

    # Lower case aliases for TextureHorizontalScale
    @property
    def texturehorizontalscale(self):
        return self.TextureHorizontalScale

    @texturehorizontalscale.setter
    def texturehorizontalscale(self, value):
        self.TextureHorizontalScale = value

    @property
    def TextureName(self):
        return self.com_object.TextureName

    # Lower case aliases for TextureName
    @property
    def texturename(self):
        return self.TextureName

    @property
    def TextureOffsetX(self):
        return self.com_object.TextureOffsetX

    @TextureOffsetX.setter
    def TextureOffsetX(self, value):
        self.com_object.TextureOffsetX = value

    # Lower case aliases for TextureOffsetX
    @property
    def textureoffsetx(self):
        return self.TextureOffsetX

    @textureoffsetx.setter
    def textureoffsetx(self, value):
        self.TextureOffsetX = value

    @property
    def TextureOffsetY(self):
        return self.com_object.TextureOffsetY

    @TextureOffsetY.setter
    def TextureOffsetY(self, value):
        self.com_object.TextureOffsetY = value

    # Lower case aliases for TextureOffsetY
    @property
    def textureoffsety(self):
        return self.TextureOffsetY

    @textureoffsety.setter
    def textureoffsety(self, value):
        self.TextureOffsetY = value

    @property
    def TextureTile(self):
        return self.com_object.TextureTile

    @TextureTile.setter
    def TextureTile(self, value):
        self.com_object.TextureTile = value

    # Lower case aliases for TextureTile
    @property
    def texturetile(self):
        return self.TextureTile

    @texturetile.setter
    def texturetile(self, value):
        self.TextureTile = value

    @property
    def TextureType(self):
        return self.com_object.TextureType

    # Lower case aliases for TextureType
    @property
    def texturetype(self):
        return self.TextureType

    @property
    def TextureVerticalScale(self):
        return self.com_object.TextureVerticalScale

    @TextureVerticalScale.setter
    def TextureVerticalScale(self, value):
        self.com_object.TextureVerticalScale = value

    # Lower case aliases for TextureVerticalScale
    @property
    def textureverticalscale(self):
        return self.TextureVerticalScale

    @textureverticalscale.setter
    def textureverticalscale(self, value):
        self.TextureVerticalScale = value

    @property
    def Transparency(self):
        return self.com_object.Transparency

    @Transparency.setter
    def Transparency(self, value):
        self.com_object.Transparency = value

    # Lower case aliases for Transparency
    @property
    def transparency(self):
        return self.Transparency

    @transparency.setter
    def transparency(self, value):
        self.Transparency = value

    @property
    def Type(self):
        return self.com_object.Type

    # Lower case aliases for Type
    @property
    def type(self):
        return self.Type

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

    def Background(self):
        self.com_object.Background()

    # Lower case alias for Background
    def background(self):
        return self.Background()

    def OneColorGradient(self, Style=None, Variant=None, Degree=None):
        arguments = com_arguments([unwrap(a) for a in [Style, Variant, Degree]])
        self.com_object.OneColorGradient(*arguments)

    # Lower case alias for OneColorGradient
    def onecolorgradient(self, Style=None, Variant=None, Degree=None):
        arguments = [Style, Variant, Degree]
        return self.OneColorGradient(*arguments)

    def Patterned(self, Pattern=None):
        arguments = com_arguments([unwrap(a) for a in [Pattern]])
        self.com_object.Patterned(*arguments)

    # Lower case alias for Patterned
    def patterned(self, Pattern=None):
        arguments = [Pattern]
        return self.Patterned(*arguments)

    def PresetGradient(self, Style=None, Variant=None, PresetGradientType=None):
        arguments = com_arguments([unwrap(a) for a in [Style, Variant, PresetGradientType]])
        self.com_object.PresetGradient(*arguments)

    # Lower case alias for PresetGradient
    def presetgradient(self, Style=None, Variant=None, PresetGradientType=None):
        arguments = [Style, Variant, PresetGradientType]
        return self.PresetGradient(*arguments)

    def PresetTextured(self, PresetTexture=None):
        arguments = com_arguments([unwrap(a) for a in [PresetTexture]])
        self.com_object.PresetTextured(*arguments)

    # Lower case alias for PresetTextured
    def presettextured(self, PresetTexture=None):
        arguments = [PresetTexture]
        return self.PresetTextured(*arguments)

    def Solid(self):
        self.com_object.Solid()

    # Lower case alias for Solid
    def solid(self):
        return self.Solid()

    def TwoColorGradient(self, Style=None, Variant=None):
        arguments = com_arguments([unwrap(a) for a in [Style, Variant]])
        self.com_object.TwoColorGradient(*arguments)

    # Lower case alias for TwoColorGradient
    def twocolorgradient(self, Style=None, Variant=None):
        arguments = [Style, Variant]
        return self.TwoColorGradient(*arguments)

    def UserPicture(self, PictureFile=None):
        arguments = com_arguments([unwrap(a) for a in [PictureFile]])
        self.com_object.UserPicture(*arguments)

    # Lower case alias for UserPicture
    def userpicture(self, PictureFile=None):
        arguments = [PictureFile]
        return self.UserPicture(*arguments)

    def UserTextured(self, TextureFile=None):
        arguments = com_arguments([unwrap(a) for a in [TextureFile]])
        self.com_object.UserTextured(*arguments)

    # Lower case alias for UserTextured
    def usertextured(self, TextureFile=None):
        arguments = [TextureFile]
        return self.UserTextured(*arguments)


class FilterEffect:

    def __init__(self, filtereffect=None):
        self.com_object= filtereffect

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Reveal(self):
        return self.com_object.Reveal

    @Reveal.setter
    def Reveal(self, value):
        self.com_object.Reveal = value

    # Lower case aliases for Reveal
    @property
    def reveal(self):
        return self.Reveal

    @reveal.setter
    def reveal(self, value):
        self.Reveal = value

    @property
    def Subtype(self):
        return self.com_object.Subtype

    @Subtype.setter
    def Subtype(self, value):
        self.com_object.Subtype = value

    # Lower case aliases for Subtype
    @property
    def subtype(self):
        return self.Subtype

    @subtype.setter
    def subtype(self, value):
        self.Subtype = value

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


class Floor:

    def __init__(self, floor=None):
        self.com_object= floor

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

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
    def PictureType(self):
        return self.com_object.PictureType

    @PictureType.setter
    def PictureType(self, value):
        self.com_object.PictureType = value

    # Lower case aliases for PictureType
    @property
    def picturetype(self):
        return self.PictureType

    @picturetype.setter
    def picturetype(self, value):
        self.PictureType = value

    @property
    def Thickness(self):
        return self.com_object.Thickness

    @Thickness.setter
    def Thickness(self, value):
        self.com_object.Thickness = value

    # Lower case aliases for Thickness
    @property
    def thickness(self):
        return self.Thickness

    @thickness.setter
    def thickness(self, value):
        self.Thickness = value

    def ClearFormats(self):
        self.com_object.ClearFormats()

    # Lower case alias for ClearFormats
    def clearformats(self):
        return self.ClearFormats()

    def Paste(self):
        self.com_object.Paste()

    # Lower case alias for Paste
    def paste(self):
        return self.Paste()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class Font:

    def __init__(self, font=None):
        self.com_object= font

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AutoRotateNumbers(self):
        return self.com_object.AutoRotateNumbers

    @AutoRotateNumbers.setter
    def AutoRotateNumbers(self, value):
        self.com_object.AutoRotateNumbers = value

    # Lower case aliases for AutoRotateNumbers
    @property
    def autorotatenumbers(self):
        return self.AutoRotateNumbers

    @autorotatenumbers.setter
    def autorotatenumbers(self, value):
        self.AutoRotateNumbers = value

    @property
    def BaselineOffset(self):
        return self.com_object.BaselineOffset

    @BaselineOffset.setter
    def BaselineOffset(self, value):
        self.com_object.BaselineOffset = value

    # Lower case aliases for BaselineOffset
    @property
    def baselineoffset(self):
        return self.BaselineOffset

    @baselineoffset.setter
    def baselineoffset(self, value):
        self.BaselineOffset = value

    @property
    def Bold(self):
        return self.com_object.Bold

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
    def Color(self):
        return Font(self.com_object.Color)

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
    def Embeddable(self):
        return self.com_object.Embeddable

    # Lower case aliases for Embeddable
    @property
    def embeddable(self):
        return self.Embeddable

    @property
    def Embedded(self):
        return self.com_object.Embedded

    # Lower case aliases for Embedded
    @property
    def embedded(self):
        return self.Embedded

    @property
    def Emboss(self):
        return self.com_object.Emboss

    @Emboss.setter
    def Emboss(self, value):
        self.com_object.Emboss = value

    # Lower case aliases for Emboss
    @property
    def emboss(self):
        return self.Emboss

    @emboss.setter
    def emboss(self, value):
        self.Emboss = value

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
    def NameAscii(self):
        return self.com_object.NameAscii

    @NameAscii.setter
    def NameAscii(self, value):
        self.com_object.NameAscii = value

    # Lower case aliases for NameAscii
    @property
    def nameascii(self):
        return self.NameAscii

    @nameascii.setter
    def nameascii(self, value):
        self.NameAscii = value

    @property
    def NameComplexScript(self):
        return self.com_object.NameComplexScript

    @NameComplexScript.setter
    def NameComplexScript(self, value):
        self.com_object.NameComplexScript = value

    # Lower case aliases for NameComplexScript
    @property
    def namecomplexscript(self):
        return self.NameComplexScript

    @namecomplexscript.setter
    def namecomplexscript(self, value):
        self.NameComplexScript = value

    @property
    def NameFarEast(self):
        return self.com_object.NameFarEast

    @NameFarEast.setter
    def NameFarEast(self, value):
        self.com_object.NameFarEast = value

    # Lower case aliases for NameFarEast
    @property
    def namefareast(self):
        return self.NameFarEast

    @namefareast.setter
    def namefareast(self, value):
        self.NameFarEast = value

    @property
    def NameOther(self):
        return self.com_object.NameOther

    @NameOther.setter
    def NameOther(self, value):
        self.com_object.NameOther = value

    # Lower case aliases for NameOther
    @property
    def nameother(self):
        return self.NameOther

    @nameother.setter
    def nameother(self, value):
        self.NameOther = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Shadow(self):
        return self.com_object.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.com_object.Shadow = value

    # Lower case aliases for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @shadow.setter
    def shadow(self, value):
        self.Shadow = value

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
    def Subscript(self):
        return self.com_object.Subscript

    @Subscript.setter
    def Subscript(self, value):
        self.com_object.Subscript = value

    # Lower case aliases for Subscript
    @property
    def subscript(self):
        return self.Subscript

    @subscript.setter
    def subscript(self, value):
        self.Subscript = value

    @property
    def Superscript(self):
        return self.com_object.Superscript

    @Superscript.setter
    def Superscript(self, value):
        self.com_object.Superscript = value

    # Lower case aliases for Superscript
    @property
    def superscript(self):
        return self.Superscript

    @superscript.setter
    def superscript(self, value):
        self.Superscript = value

    @property
    def Underline(self):
        return self.com_object.Underline

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


class Fonts:

    def __init__(self, fonts=None):
        self.com_object= fonts

    def __call__(self, item):
        return Font(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Font(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Replace(self, Original=None, Replacement=None):
        arguments = com_arguments([unwrap(a) for a in [Original, Replacement]])
        self.com_object.Replace(*arguments)

    # Lower case alias for Replace
    def replace(self, Original=None, Replacement=None):
        arguments = [Original, Replacement]
        return self.Replace(*arguments)


class FreeformBuilder:

    def __init__(self, freeformbuilder=None):
        self.com_object= freeformbuilder

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    def AddNodes(self, SegmentType=None, EditingType=None, X1=None, Y1=None, X2=None, Y2=None, X3=None, Y3=None):
        arguments = com_arguments([unwrap(a) for a in [SegmentType, EditingType, X1, Y1, X2, Y2, X3, Y3]])
        self.com_object.AddNodes(*arguments)

    # Lower case alias for AddNodes
    def addnodes(self, SegmentType=None, EditingType=None, X1=None, Y1=None, X2=None, Y2=None, X3=None, Y3=None):
        arguments = [SegmentType, EditingType, X1, Y1, X2, Y2, X3, Y3]
        return self.AddNodes(*arguments)

    def ConvertToShape(self):
        return Shape(self.com_object.ConvertToShape())

    # Lower case alias for ConvertToShape
    def converttoshape(self):
        return self.ConvertToShape()


class GridLines:

    def __init__(self, gridlines=None):
        self.com_object= gridlines

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Border(self):
        return ChartBorder(self.com_object.Border)

    # Lower case aliases for Border
    @property
    def border(self):
        return self.Border

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

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

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class GroupShapes:

    def __init__(self, groupshapes=None):
        self.com_object= groupshapes

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Shape(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Range(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return ShapeRange(self.com_object.Range(*arguments))

    # Lower case alias for Range
    def range(self, Index=None):
        arguments = [Index]
        return self.Range(*arguments)


class HeaderFooter:

    def __init__(self, headerfooter=None):
        self.com_object= headerfooter

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Format(self):
        return self.com_object.Format

    @Format.setter
    def Format(self, value):
        self.com_object.Format = value

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

    @format.setter
    def format(self, value):
        self.Format = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

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
    def UseFormat(self):
        return self.com_object.UseFormat

    @UseFormat.setter
    def UseFormat(self, value):
        self.com_object.UseFormat = value

    # Lower case aliases for UseFormat
    @property
    def useformat(self):
        return self.UseFormat

    @useformat.setter
    def useformat(self, value):
        self.UseFormat = value

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


class HeadersFooters:

    def __init__(self, headersfooters=None):
        self.com_object= headersfooters

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def DateAndTime(self):
        return HeaderFooter(self.com_object.DateAndTime)

    # Lower case aliases for DateAndTime
    @property
    def dateandtime(self):
        return self.DateAndTime

    @property
    def DisplayOnTitleSlide(self):
        return self.com_object.DisplayOnTitleSlide

    @DisplayOnTitleSlide.setter
    def DisplayOnTitleSlide(self, value):
        self.com_object.DisplayOnTitleSlide = value

    # Lower case aliases for DisplayOnTitleSlide
    @property
    def displayontitleslide(self):
        return self.DisplayOnTitleSlide

    @displayontitleslide.setter
    def displayontitleslide(self, value):
        self.DisplayOnTitleSlide = value

    @property
    def Footer(self):
        return HeaderFooter(self.com_object.Footer)

    # Lower case aliases for Footer
    @property
    def footer(self):
        return self.Footer

    @property
    def Header(self):
        return HeaderFooter(self.com_object.Header)

    # Lower case aliases for Header
    @property
    def header(self):
        return self.Header

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SlideNumber(self):
        return HeaderFooter(self.com_object.SlideNumber)

    # Lower case aliases for SlideNumber
    @property
    def slidenumber(self):
        return self.SlideNumber

    def Clear(self):
        self.com_object.Clear()

    # Lower case alias for Clear
    def clear(self):
        return self.Clear()


class HiLoLines:

    def __init__(self, hilolines=None):
        self.com_object= hilolines

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Border(self):
        return ChartBorder(self.com_object.Border)

    # Lower case aliases for Border
    @property
    def border(self):
        return self.Border

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

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

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class Hyperlink:

    def __init__(self, hyperlink=None):
        self.com_object= hyperlink

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
    def EmailSubject(self):
        return self.com_object.EmailSubject

    @EmailSubject.setter
    def EmailSubject(self, value):
        self.com_object.EmailSubject = value

    # Lower case aliases for EmailSubject
    @property
    def emailsubject(self):
        return self.EmailSubject

    @emailsubject.setter
    def emailsubject(self, value):
        self.EmailSubject = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ScreenTip(self):
        return self.com_object.ScreenTip

    @ScreenTip.setter
    def ScreenTip(self, value):
        self.com_object.ScreenTip = value

    # Lower case aliases for ScreenTip
    @property
    def screentip(self):
        return self.ScreenTip

    @screentip.setter
    def screentip(self, value):
        self.ScreenTip = value

    @property
    def ShowAndReturn(self):
        return self.com_object.ShowAndReturn

    @ShowAndReturn.setter
    def ShowAndReturn(self, value):
        self.com_object.ShowAndReturn = value

    # Lower case aliases for ShowAndReturn
    @property
    def showandreturn(self):
        return self.ShowAndReturn

    @showandreturn.setter
    def showandreturn(self, value):
        self.ShowAndReturn = value

    @property
    def SubAddress(self):
        return self.com_object.SubAddress

    @SubAddress.setter
    def SubAddress(self, value):
        self.com_object.SubAddress = value

    # Lower case aliases for SubAddress
    @property
    def subaddress(self):
        return self.SubAddress

    @subaddress.setter
    def subaddress(self, value):
        self.SubAddress = value

    @property
    def TextToDisplay(self):
        return self.com_object.TextToDisplay

    @TextToDisplay.setter
    def TextToDisplay(self, value):
        self.com_object.TextToDisplay = value

    # Lower case aliases for TextToDisplay
    @property
    def texttodisplay(self):
        return self.TextToDisplay

    @texttodisplay.setter
    def texttodisplay(self, value):
        self.TextToDisplay = value

    @property
    def Type(self):
        return self.com_object.Type

    # Lower case aliases for Type
    @property
    def type(self):
        return self.Type

    def AddToFavorites(self):
        self.com_object.AddToFavorites()

    # Lower case alias for AddToFavorites
    def addtofavorites(self):
        return self.AddToFavorites()

    def CreateNewDocument(self, FileName=None, EditNow=None, Overwrite=None):
        arguments = com_arguments([unwrap(a) for a in [FileName, EditNow, Overwrite]])
        return self.com_object.CreateNewDocument(*arguments)

    # Lower case alias for CreateNewDocument
    def createnewdocument(self, FileName=None, EditNow=None, Overwrite=None):
        arguments = [FileName, EditNow, Overwrite]
        return self.CreateNewDocument(*arguments)

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Follow(self):
        self.com_object.Follow()

    # Lower case alias for Follow
    def follow(self):
        return self.Follow()


class Hyperlinks:

    def __init__(self, hyperlinks=None):
        self.com_object= hyperlinks

    def __call__(self, item):
        return Hyperlink(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Hyperlink(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class Interior:

    def __init__(self, interior=None):
        self.com_object= interior

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Color(self):
        return self.com_object.Color

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
    def ColorIndex(self):
        return self.com_object.ColorIndex

    @ColorIndex.setter
    def ColorIndex(self, value):
        self.com_object.ColorIndex = value

    # Lower case aliases for ColorIndex
    @property
    def colorindex(self):
        return self.ColorIndex

    @colorindex.setter
    def colorindex(self, value):
        self.ColorIndex = value

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def InvertIfNegative(self):
        return self.com_object.InvertIfNegative

    @InvertIfNegative.setter
    def InvertIfNegative(self, value):
        self.com_object.InvertIfNegative = value

    # Lower case aliases for InvertIfNegative
    @property
    def invertifnegative(self):
        return self.InvertIfNegative

    @invertifnegative.setter
    def invertifnegative(self, value):
        self.InvertIfNegative = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Pattern(self):
        return XlPattern(self.com_object.Pattern)

    @Pattern.setter
    def Pattern(self, value):
        self.com_object.Pattern = value

    # Lower case aliases for Pattern
    @property
    def pattern(self):
        return self.Pattern

    @pattern.setter
    def pattern(self, value):
        self.Pattern = value

    @property
    def PatternColor(self):
        return self.com_object.PatternColor

    @PatternColor.setter
    def PatternColor(self, value):
        self.com_object.PatternColor = value

    # Lower case aliases for PatternColor
    @property
    def patterncolor(self):
        return self.PatternColor

    @patterncolor.setter
    def patterncolor(self, value):
        self.PatternColor = value

    @property
    def PatternColorIndex(self):
        return XlColorIndex(self.com_object.PatternColorIndex)

    @PatternColorIndex.setter
    def PatternColorIndex(self, value):
        self.com_object.PatternColorIndex = value

    # Lower case aliases for PatternColorIndex
    @property
    def patterncolorindex(self):
        return self.PatternColorIndex

    @patterncolorindex.setter
    def patterncolorindex(self, value):
        self.PatternColorIndex = value


class LeaderLines:

    def __init__(self, leaderlines=None):
        self.com_object= leaderlines

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Border(self):
        return ChartBorder(self.com_object.Border)

    # Lower case aliases for Border
    @property
    def border(self):
        return self.Border

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class Legend:

    def __init__(self, legend=None):
        self.com_object= legend

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

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
    def IncludeInLayout(self):
        return self.com_object.IncludeInLayout

    @IncludeInLayout.setter
    def IncludeInLayout(self, value):
        self.com_object.IncludeInLayout = value

    # Lower case aliases for IncludeInLayout
    @property
    def includeinlayout(self):
        return self.IncludeInLayout

    @includeinlayout.setter
    def includeinlayout(self, value):
        self.IncludeInLayout = value

    @property
    def Left(self):
        return self.com_object.Left

    # Lower case aliases for Left
    @property
    def left(self):
        return self.Left

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
    def Position(self):
        return XlLegendPosition(self.com_object.Position)

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
    def Shadow(self):
        return self.com_object.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.com_object.Shadow = value

    # Lower case aliases for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @shadow.setter
    def shadow(self, value):
        self.Shadow = value

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

    def Clear(self):
        self.com_object.Clear()

    # Lower case alias for Clear
    def clear(self):
        return self.Clear()

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def LegendEntries(self):
        return LegendEntries(self.com_object.LegendEntries())

    # Lower case alias for LegendEntries
    def legendentries(self):
        return self.LegendEntries()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class LegendEntries:

    def __init__(self, legendentries=None):
        self.com_object= legendentries

    def __call__(self, item):
        return LegendEntrie(self.com_object(item))

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return LegendEntry(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class LegendEntry:

    def __init__(self, legendentry=None):
        self.com_object= legendentry

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Font(self):
        return ChartFont(self.com_object.Font)

    # Lower case aliases for Font
    @property
    def font(self):
        return self.Font

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

    @property
    def Height(self):
        return self.com_object.Height

    # Lower case aliases for Height
    @property
    def height(self):
        return self.Height

    @property
    def Index(self):
        return self.com_object.Index

    # Lower case aliases for Index
    @property
    def index(self):
        return self.Index

    @property
    def Left(self):
        return self.com_object.Left

    # Lower case aliases for Left
    @property
    def left(self):
        return self.Left

    @property
    def LegendKey(self):
        return LegendKey(self.com_object.LegendKey)

    # Lower case aliases for LegendKey
    @property
    def legendkey(self):
        return self.LegendKey

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Top(self):
        return self.com_object.Top

    # Lower case aliases for Top
    @property
    def top(self):
        return self.Top

    @property
    def Width(self):
        return self.com_object.Width

    # Lower case aliases for Width
    @property
    def width(self):
        return self.Width

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class LegendKey:

    def __init__(self, legendkey=None):
        self.com_object= legendkey

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

    @property
    def Height(self):
        return self.com_object.Height

    # Lower case aliases for Height
    @property
    def height(self):
        return self.Height

    @property
    def InvertIfNegative(self):
        return self.com_object.InvertIfNegative

    @InvertIfNegative.setter
    def InvertIfNegative(self, value):
        self.com_object.InvertIfNegative = value

    # Lower case aliases for InvertIfNegative
    @property
    def invertifnegative(self):
        return self.InvertIfNegative

    @invertifnegative.setter
    def invertifnegative(self, value):
        self.InvertIfNegative = value

    @property
    def Left(self):
        return self.com_object.Left

    # Lower case aliases for Left
    @property
    def left(self):
        return self.Left

    @property
    def MarkerBackgroundColor(self):
        return self.com_object.MarkerBackgroundColor

    @MarkerBackgroundColor.setter
    def MarkerBackgroundColor(self, value):
        self.com_object.MarkerBackgroundColor = value

    # Lower case aliases for MarkerBackgroundColor
    @property
    def markerbackgroundcolor(self):
        return self.MarkerBackgroundColor

    @markerbackgroundcolor.setter
    def markerbackgroundcolor(self, value):
        self.MarkerBackgroundColor = value

    @property
    def MarkerBackgroundColorIndex(self):
        return XlColorIndex(self.com_object.MarkerBackgroundColorIndex)

    @MarkerBackgroundColorIndex.setter
    def MarkerBackgroundColorIndex(self, value):
        self.com_object.MarkerBackgroundColorIndex = value

    # Lower case aliases for MarkerBackgroundColorIndex
    @property
    def markerbackgroundcolorindex(self):
        return self.MarkerBackgroundColorIndex

    @markerbackgroundcolorindex.setter
    def markerbackgroundcolorindex(self, value):
        self.MarkerBackgroundColorIndex = value

    @property
    def MarkerForegroundColor(self):
        return self.com_object.MarkerForegroundColor

    @MarkerForegroundColor.setter
    def MarkerForegroundColor(self, value):
        self.com_object.MarkerForegroundColor = value

    # Lower case aliases for MarkerForegroundColor
    @property
    def markerforegroundcolor(self):
        return self.MarkerForegroundColor

    @markerforegroundcolor.setter
    def markerforegroundcolor(self, value):
        self.MarkerForegroundColor = value

    @property
    def MarkerForegroundColorIndex(self):
        return XlColorIndex(self.com_object.MarkerForegroundColorIndex)

    @MarkerForegroundColorIndex.setter
    def MarkerForegroundColorIndex(self, value):
        self.com_object.MarkerForegroundColorIndex = value

    # Lower case aliases for MarkerForegroundColorIndex
    @property
    def markerforegroundcolorindex(self):
        return self.MarkerForegroundColorIndex

    @markerforegroundcolorindex.setter
    def markerforegroundcolorindex(self, value):
        self.MarkerForegroundColorIndex = value

    @property
    def MarkerSize(self):
        return self.com_object.MarkerSize

    @MarkerSize.setter
    def MarkerSize(self, value):
        self.com_object.MarkerSize = value

    # Lower case aliases for MarkerSize
    @property
    def markersize(self):
        return self.MarkerSize

    @markersize.setter
    def markersize(self, value):
        self.MarkerSize = value

    @property
    def MarkerStyle(self):
        return XlMarkerStyle(self.com_object.MarkerStyle)

    @MarkerStyle.setter
    def MarkerStyle(self, value):
        self.com_object.MarkerStyle = value

    # Lower case aliases for MarkerStyle
    @property
    def markerstyle(self):
        return self.MarkerStyle

    @markerstyle.setter
    def markerstyle(self, value):
        self.MarkerStyle = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PictureType(self):
        return XlChartPictureType(self.com_object.PictureType)

    @PictureType.setter
    def PictureType(self, value):
        self.com_object.PictureType = value

    # Lower case aliases for PictureType
    @property
    def picturetype(self):
        return self.PictureType

    @picturetype.setter
    def picturetype(self, value):
        self.PictureType = value

    @property
    def PictureUnit2(self):
        return self.com_object.PictureUnit2

    @PictureUnit2.setter
    def PictureUnit2(self, value):
        self.com_object.PictureUnit2 = value

    # Lower case aliases for PictureUnit2
    @property
    def pictureunit2(self):
        return self.PictureUnit2

    @pictureunit2.setter
    def pictureunit2(self, value):
        self.PictureUnit2 = value

    @property
    def Shadow(self):
        return self.com_object.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.com_object.Shadow = value

    # Lower case aliases for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @shadow.setter
    def shadow(self, value):
        self.Shadow = value

    @property
    def Smooth(self):
        return self.com_object.Smooth

    @Smooth.setter
    def Smooth(self, value):
        self.com_object.Smooth = value

    # Lower case aliases for Smooth
    @property
    def smooth(self):
        return self.Smooth

    @smooth.setter
    def smooth(self, value):
        self.Smooth = value

    @property
    def Top(self):
        return self.com_object.Top

    # Lower case aliases for Top
    @property
    def top(self):
        return self.Top

    @property
    def Width(self):
        return self.com_object.Width

    # Lower case aliases for Width
    @property
    def width(self):
        return self.Width

    def ClearFormats(self):
        self.com_object.ClearFormats()

    # Lower case alias for ClearFormats
    def clearformats(self):
        return self.ClearFormats()

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()


class LineFormat:

    def __init__(self, lineformat=None):
        self.com_object= lineformat

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def BackColor(self):
        return ColorFormat(self.com_object.BackColor)

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
    def BeginArrowheadLength(self):
        return self.com_object.BeginArrowheadLength

    @BeginArrowheadLength.setter
    def BeginArrowheadLength(self, value):
        self.com_object.BeginArrowheadLength = value

    # Lower case aliases for BeginArrowheadLength
    @property
    def beginarrowheadlength(self):
        return self.BeginArrowheadLength

    @beginarrowheadlength.setter
    def beginarrowheadlength(self, value):
        self.BeginArrowheadLength = value

    @property
    def BeginArrowheadStyle(self):
        return self.com_object.BeginArrowheadStyle

    @BeginArrowheadStyle.setter
    def BeginArrowheadStyle(self, value):
        self.com_object.BeginArrowheadStyle = value

    # Lower case aliases for BeginArrowheadStyle
    @property
    def beginarrowheadstyle(self):
        return self.BeginArrowheadStyle

    @beginarrowheadstyle.setter
    def beginarrowheadstyle(self, value):
        self.BeginArrowheadStyle = value

    @property
    def BeginArrowheadWidth(self):
        return self.com_object.BeginArrowheadWidth

    @BeginArrowheadWidth.setter
    def BeginArrowheadWidth(self, value):
        self.com_object.BeginArrowheadWidth = value

    # Lower case aliases for BeginArrowheadWidth
    @property
    def beginarrowheadwidth(self):
        return self.BeginArrowheadWidth

    @beginarrowheadwidth.setter
    def beginarrowheadwidth(self, value):
        self.BeginArrowheadWidth = value

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def DashStyle(self):
        return self.com_object.DashStyle

    @DashStyle.setter
    def DashStyle(self, value):
        self.com_object.DashStyle = value

    # Lower case aliases for DashStyle
    @property
    def dashstyle(self):
        return self.DashStyle

    @dashstyle.setter
    def dashstyle(self, value):
        self.DashStyle = value

    @property
    def EndArrowheadLength(self):
        return self.com_object.EndArrowheadLength

    @EndArrowheadLength.setter
    def EndArrowheadLength(self, value):
        self.com_object.EndArrowheadLength = value

    # Lower case aliases for EndArrowheadLength
    @property
    def endarrowheadlength(self):
        return self.EndArrowheadLength

    @endarrowheadlength.setter
    def endarrowheadlength(self, value):
        self.EndArrowheadLength = value

    @property
    def EndArrowheadStyle(self):
        return self.com_object.EndArrowheadStyle

    @EndArrowheadStyle.setter
    def EndArrowheadStyle(self, value):
        self.com_object.EndArrowheadStyle = value

    # Lower case aliases for EndArrowheadStyle
    @property
    def endarrowheadstyle(self):
        return self.EndArrowheadStyle

    @endarrowheadstyle.setter
    def endarrowheadstyle(self, value):
        self.EndArrowheadStyle = value

    @property
    def EndArrowheadWidth(self):
        return self.com_object.EndArrowheadWidth

    @EndArrowheadWidth.setter
    def EndArrowheadWidth(self, value):
        self.com_object.EndArrowheadWidth = value

    # Lower case aliases for EndArrowheadWidth
    @property
    def endarrowheadwidth(self):
        return self.EndArrowheadWidth

    @endarrowheadwidth.setter
    def endarrowheadwidth(self, value):
        self.EndArrowheadWidth = value

    @property
    def ForeColor(self):
        return ColorFormat(self.com_object.ForeColor)

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
    def InsetPen(self):
        return self.com_object.InsetPen

    @InsetPen.setter
    def InsetPen(self, value):
        self.com_object.InsetPen = value

    # Lower case aliases for InsetPen
    @property
    def insetpen(self):
        return self.InsetPen

    @insetpen.setter
    def insetpen(self, value):
        self.InsetPen = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Pattern(self):
        return self.com_object.Pattern

    @Pattern.setter
    def Pattern(self, value):
        self.com_object.Pattern = value

    # Lower case aliases for Pattern
    @property
    def pattern(self):
        return self.Pattern

    @pattern.setter
    def pattern(self, value):
        self.Pattern = value

    @property
    def Style(self):
        return self.com_object.Style

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
    def Transparency(self):
        return self.com_object.Transparency

    @Transparency.setter
    def Transparency(self, value):
        self.com_object.Transparency = value

    # Lower case aliases for Transparency
    @property
    def transparency(self):
        return self.Transparency

    @transparency.setter
    def transparency(self, value):
        self.Transparency = value

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

    @property
    def Weight(self):
        return self.com_object.Weight

    @Weight.setter
    def Weight(self, value):
        self.com_object.Weight = value

    # Lower case aliases for Weight
    @property
    def weight(self):
        return self.Weight

    @weight.setter
    def weight(self, value):
        self.Weight = value


class LinkFormat:

    def __init__(self, linkformat=None):
        self.com_object= linkformat

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AutoUpdate(self):
        return self.com_object.AutoUpdate

    @AutoUpdate.setter
    def AutoUpdate(self, value):
        self.com_object.AutoUpdate = value

    # Lower case aliases for AutoUpdate
    @property
    def autoupdate(self):
        return self.AutoUpdate

    @autoupdate.setter
    def autoupdate(self, value):
        self.AutoUpdate = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SourceFullName(self):
        return self.com_object.SourceFullName

    @SourceFullName.setter
    def SourceFullName(self, value):
        self.com_object.SourceFullName = value

    # Lower case aliases for SourceFullName
    @property
    def sourcefullname(self):
        return self.SourceFullName

    @sourcefullname.setter
    def sourcefullname(self, value):
        self.SourceFullName = value

    def BreakLink(self):
        return self.com_object.BreakLink()

    # Lower case alias for BreakLink
    def breaklink(self):
        return self.BreakLink()

    def Update(self):
        self.com_object.Update()

    # Lower case alias for Update
    def update(self):
        return self.Update()


class Master:

    def __init__(self, master=None):
        self.com_object= master

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Background(self):
        return ShapeRange(self.com_object.Background)

    # Lower case aliases for Background
    @property
    def background(self):
        return self.Background

    @property
    def BackgroundStyle(self):
        return self.com_object.BackgroundStyle

    @BackgroundStyle.setter
    def BackgroundStyle(self, value):
        self.com_object.BackgroundStyle = value

    # Lower case aliases for BackgroundStyle
    @property
    def backgroundstyle(self):
        return self.BackgroundStyle

    @backgroundstyle.setter
    def backgroundstyle(self, value):
        self.BackgroundStyle = value

    @property
    def ColorScheme(self):
        return ColorScheme(self.com_object.ColorScheme)

    @ColorScheme.setter
    def ColorScheme(self, value):
        self.com_object.ColorScheme = value

    # Lower case aliases for ColorScheme
    @property
    def colorscheme(self):
        return self.ColorScheme

    @colorscheme.setter
    def colorscheme(self, value):
        self.ColorScheme = value

    @property
    def CustomerData(self):
        return CustomerData(self.com_object.CustomerData)

    # Lower case aliases for CustomerData
    @property
    def customerdata(self):
        return self.CustomerData

    @property
    def CustomLayouts(self):
        return CustomLayouts(self.com_object.CustomLayouts)

    # Lower case aliases for CustomLayouts
    @property
    def customlayouts(self):
        return self.CustomLayouts

    @property
    def Design(self):
        return Design(self.com_object.Design)

    # Lower case aliases for Design
    @property
    def design(self):
        return self.Design

    @property
    def HeadersFooters(self):
        return HeadersFooters(self.com_object.HeadersFooters)

    # Lower case aliases for HeadersFooters
    @property
    def headersfooters(self):
        return self.HeadersFooters

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
    def Hyperlinks(self):
        return Hyperlinks(self.com_object.Hyperlinks)

    # Lower case aliases for Hyperlinks
    @property
    def hyperlinks(self):
        return self.Hyperlinks

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
    def Shapes(self):
        return Shapes(self.com_object.Shapes)

    # Lower case aliases for Shapes
    @property
    def shapes(self):
        return self.Shapes

    @property
    def SlideShowTransition(self):
        return SlideShowTransition(self.com_object.SlideShowTransition)

    # Lower case aliases for SlideShowTransition
    @property
    def slideshowtransition(self):
        return self.SlideShowTransition

    @property
    def TextStyles(self):
        return TextStyles(self.com_object.TextStyles)

    # Lower case aliases for TextStyles
    @property
    def textstyles(self):
        return self.TextStyles

    @property
    def Theme(self):
        return self.com_object.Theme

    # Lower case aliases for Theme
    @property
    def theme(self):
        return self.Theme

    @property
    def TimeLine(self):
        return TimeLine(self.com_object.TimeLine)

    # Lower case aliases for TimeLine
    @property
    def timeline(self):
        return self.TimeLine

    @property
    def Width(self):
        return self.com_object.Width

    # Lower case aliases for Width
    @property
    def width(self):
        return self.Width

    def ApplyTheme(self, themeName=None):
        arguments = com_arguments([unwrap(a) for a in [themeName]])
        self.com_object.ApplyTheme(*arguments)

    # Lower case alias for ApplyTheme
    def applytheme(self, themeName=None):
        arguments = [themeName]
        return self.ApplyTheme(*arguments)

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()


class MediaBookmark:

    def __init__(self, mediabookmark=None):
        self.com_object= mediabookmark

    @property
    def Index(self):
        return self.com_object.Index

    # Lower case aliases for Index
    @property
    def index(self):
        return self.Index

    @property
    def Name(self):
        return self.com_object.Name

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def Position(self):
        return self.com_object.Position

    # Lower case aliases for Position
    @property
    def position(self):
        return self.Position

    def Delete(self):
        return self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()


class MediaBookmarks:

    def __init__(self, mediabookmarks=None):
        self.com_object= mediabookmarks

    def __call__(self, item):
        return MediaBookmark(self.com_object(item))

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    def Add(self, Position=None, Name=None):
        arguments = com_arguments([unwrap(a) for a in [Position, Name]])
        return MediaBookmark(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Position=None, Name=None):
        arguments = [Position, Name]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return MediaBookmark(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class MediaFormat:

    def __init__(self, mediaformat=None):
        self.com_object= mediaformat

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AudioCompressionType(self):
        return self.com_object.AudioCompressionType

    # Lower case aliases for AudioCompressionType
    @property
    def audiocompressiontype(self):
        return self.AudioCompressionType

    @property
    def AudioSamplingRate(self):
        return self.com_object.AudioSamplingRate

    # Lower case aliases for AudioSamplingRate
    @property
    def audiosamplingrate(self):
        return self.AudioSamplingRate

    @property
    def EndPoint(self):
        return self.com_object.EndPoint

    @EndPoint.setter
    def EndPoint(self, value):
        self.com_object.EndPoint = value

    # Lower case aliases for EndPoint
    @property
    def endpoint(self):
        return self.EndPoint

    @endpoint.setter
    def endpoint(self, value):
        self.EndPoint = value

    @property
    def FadeInDuration(self):
        return self.com_object.FadeInDuration

    @FadeInDuration.setter
    def FadeInDuration(self, value):
        self.com_object.FadeInDuration = value

    # Lower case aliases for FadeInDuration
    @property
    def fadeinduration(self):
        return self.FadeInDuration

    @fadeinduration.setter
    def fadeinduration(self, value):
        self.FadeInDuration = value

    @property
    def FadeOutDuration(self):
        return self.com_object.FadeOutDuration

    @FadeOutDuration.setter
    def FadeOutDuration(self, value):
        self.com_object.FadeOutDuration = value

    # Lower case aliases for FadeOutDuration
    @property
    def fadeoutduration(self):
        return self.FadeOutDuration

    @fadeoutduration.setter
    def fadeoutduration(self, value):
        self.FadeOutDuration = value

    @property
    def IsEmbedded(self):
        return self.com_object.IsEmbedded

    # Lower case aliases for IsEmbedded
    @property
    def isembedded(self):
        return self.IsEmbedded

    @property
    def IsLinked(self):
        return self.com_object.IsLinked

    # Lower case aliases for IsLinked
    @property
    def islinked(self):
        return self.IsLinked

    @property
    def Length(self):
        return self.com_object.Length

    # Lower case aliases for Length
    @property
    def length(self):
        return self.Length

    @property
    def MediaBookmarks(self):
        return MediaBookmarks(self.com_object.MediaBookmarks)

    # Lower case aliases for MediaBookmarks
    @property
    def mediabookmarks(self):
        return self.MediaBookmarks

    @property
    def Muted(self):
        return self.com_object.Muted

    @Muted.setter
    def Muted(self, value):
        self.com_object.Muted = value

    # Lower case aliases for Muted
    @property
    def muted(self):
        return self.Muted

    @muted.setter
    def muted(self, value):
        self.Muted = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ResamplingStatus(self):
        return self.com_object.ResamplingStatus

    # Lower case aliases for ResamplingStatus
    @property
    def resamplingstatus(self):
        return self.ResamplingStatus

    @property
    def SampleHeight(self):
        return self.com_object.SampleHeight

    # Lower case aliases for SampleHeight
    @property
    def sampleheight(self):
        return self.SampleHeight

    @property
    def SampleWidth(self):
        return self.com_object.SampleWidth

    # Lower case aliases for SampleWidth
    @property
    def samplewidth(self):
        return self.SampleWidth

    @property
    def StartPoint(self):
        return self.com_object.StartPoint

    @StartPoint.setter
    def StartPoint(self, value):
        self.com_object.StartPoint = value

    # Lower case aliases for StartPoint
    @property
    def startpoint(self):
        return self.StartPoint

    @startpoint.setter
    def startpoint(self, value):
        self.StartPoint = value

    @property
    def VideoCompressionType(self):
        return self.com_object.VideoCompressionType

    # Lower case aliases for VideoCompressionType
    @property
    def videocompressiontype(self):
        return self.VideoCompressionType

    @property
    def VideoFrameRate(self):
        return self.com_object.VideoFrameRate

    # Lower case aliases for VideoFrameRate
    @property
    def videoframerate(self):
        return self.VideoFrameRate

    @property
    def Volume(self):
        return self.com_object.Volume

    @Volume.setter
    def Volume(self, value):
        self.com_object.Volume = value

    # Lower case aliases for Volume
    @property
    def volume(self):
        return self.Volume

    @volume.setter
    def volume(self, value):
        self.Volume = value

    def Resample(self, Trim=None, SampleHeight=None, SampleWidth=None, VideoFrameRate=None, AudioSamplingRate=None, VideoBitRate=None):
        arguments = com_arguments([unwrap(a) for a in [Trim, SampleHeight, SampleWidth, VideoFrameRate, AudioSamplingRate, VideoBitRate]])
        return self.com_object.Resample(*arguments)

    # Lower case alias for Resample
    def resample(self, Trim=None, SampleHeight=None, SampleWidth=None, VideoFrameRate=None, AudioSamplingRate=None, VideoBitRate=None):
        arguments = [Trim, SampleHeight, SampleWidth, VideoFrameRate, AudioSamplingRate, VideoBitRate]
        return self.Resample(*arguments)

    def ResampleFromProfile(self, profile=None):
        arguments = com_arguments([unwrap(a) for a in [profile]])
        return self.com_object.ResampleFromProfile(*arguments)

    # Lower case alias for ResampleFromProfile
    def resamplefromprofile(self, profile=None):
        arguments = [profile]
        return self.ResampleFromProfile(*arguments)

    def SetDisplayPicture(self, Position=None):
        arguments = com_arguments([unwrap(a) for a in [Position]])
        return self.com_object.SetDisplayPicture(*arguments)

    # Lower case alias for SetDisplayPicture
    def setdisplaypicture(self, Position=None):
        arguments = [Position]
        return self.SetDisplayPicture(*arguments)

    def SetDisplayPictureFromFile(self, FilePath=None):
        arguments = com_arguments([unwrap(a) for a in [FilePath]])
        return self.com_object.SetDisplayPictureFromFile(*arguments)

    # Lower case alias for SetDisplayPictureFromFile
    def setdisplaypicturefromfile(self, FilePath=None):
        arguments = [FilePath]
        return self.SetDisplayPictureFromFile(*arguments)


class Model3DFormat:

    def __init__(self, model3dformat=None):
        self.com_object= model3dformat

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AutoFit(self):
        return self.com_object.AutoFit

    @AutoFit.setter
    def AutoFit(self, value):
        self.com_object.AutoFit = value

    # Lower case aliases for AutoFit
    @property
    def autofit(self):
        return self.AutoFit

    @autofit.setter
    def autofit(self, value):
        self.AutoFit = value

    @property
    def CameraPositionX(self):
        return self.com_object.CameraPositionX

    @CameraPositionX.setter
    def CameraPositionX(self, value):
        self.com_object.CameraPositionX = value

    # Lower case aliases for CameraPositionX
    @property
    def camerapositionx(self):
        return self.CameraPositionX

    @camerapositionx.setter
    def camerapositionx(self, value):
        self.CameraPositionX = value

    @property
    def CameraPositionY(self):
        return self.com_object.CameraPositionY

    @CameraPositionY.setter
    def CameraPositionY(self, value):
        self.com_object.CameraPositionY = value

    # Lower case aliases for CameraPositionY
    @property
    def camerapositiony(self):
        return self.CameraPositionY

    @camerapositiony.setter
    def camerapositiony(self, value):
        self.CameraPositionY = value

    @property
    def CameraPositionZ(self):
        return self.com_object.CameraPositionZ

    @CameraPositionZ.setter
    def CameraPositionZ(self, value):
        self.com_object.CameraPositionZ = value

    # Lower case aliases for CameraPositionZ
    @property
    def camerapositionz(self):
        return self.CameraPositionZ

    @camerapositionz.setter
    def camerapositionz(self, value):
        self.CameraPositionZ = value

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def FieldOfView(self):
        return self.com_object.FieldOfView

    @FieldOfView.setter
    def FieldOfView(self, value):
        self.com_object.FieldOfView = value

    # Lower case aliases for FieldOfView
    @property
    def fieldofview(self):
        return self.FieldOfView

    @fieldofview.setter
    def fieldofview(self, value):
        self.FieldOfView = value

    @property
    def LookAtPointX(self):
        return self.com_object.LookAtPointX

    @LookAtPointX.setter
    def LookAtPointX(self, value):
        self.com_object.LookAtPointX = value

    # Lower case aliases for LookAtPointX
    @property
    def lookatpointx(self):
        return self.LookAtPointX

    @lookatpointx.setter
    def lookatpointx(self, value):
        self.LookAtPointX = value

    @property
    def LookAtPointY(self):
        return self.com_object.LookAtPointY

    @LookAtPointY.setter
    def LookAtPointY(self, value):
        self.com_object.LookAtPointY = value

    # Lower case aliases for LookAtPointY
    @property
    def lookatpointy(self):
        return self.LookAtPointY

    @lookatpointy.setter
    def lookatpointy(self, value):
        self.LookAtPointY = value

    @property
    def LookAtPointZ(self):
        return self.com_object.LookAtPointZ

    @LookAtPointZ.setter
    def LookAtPointZ(self, value):
        self.com_object.LookAtPointZ = value

    # Lower case aliases for LookAtPointZ
    @property
    def lookatpointz(self):
        return self.LookAtPointZ

    @lookatpointz.setter
    def lookatpointz(self, value):
        self.LookAtPointZ = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RotationX(self):
        return self.com_object.RotationX

    @RotationX.setter
    def RotationX(self, value):
        self.com_object.RotationX = value

    # Lower case aliases for RotationX
    @property
    def rotationx(self):
        return self.RotationX

    @rotationx.setter
    def rotationx(self, value):
        self.RotationX = value

    @property
    def RotationY(self):
        return self.com_object.RotationY

    @RotationY.setter
    def RotationY(self, value):
        self.com_object.RotationY = value

    # Lower case aliases for RotationY
    @property
    def rotationy(self):
        return self.RotationY

    @rotationy.setter
    def rotationy(self, value):
        self.RotationY = value

    @property
    def RotationZ(self):
        return self.com_object.RotationZ

    @RotationZ.setter
    def RotationZ(self, value):
        self.com_object.RotationZ = value

    # Lower case aliases for RotationZ
    @property
    def rotationz(self):
        return self.RotationZ

    @rotationz.setter
    def rotationz(self, value):
        self.RotationZ = value

    def IncrementRotationX(self, Increment=None):
        arguments = com_arguments([unwrap(a) for a in [Increment]])
        self.com_object.IncrementRotationX(*arguments)

    # Lower case alias for IncrementRotationX
    def incrementrotationx(self, Increment=None):
        arguments = [Increment]
        return self.IncrementRotationX(*arguments)

    def IncrementRotationY(self, Increment=None):
        arguments = com_arguments([unwrap(a) for a in [Increment]])
        self.com_object.IncrementRotationY(*arguments)

    # Lower case alias for IncrementRotationY
    def incrementrotationy(self, Increment=None):
        arguments = [Increment]
        return self.IncrementRotationY(*arguments)

    def IncrementRotationZ(self, Increment=None):
        arguments = com_arguments([unwrap(a) for a in [Increment]])
        self.com_object.IncrementRotationZ(*arguments)

    # Lower case alias for IncrementRotationZ
    def incrementrotationz(self, Increment=None):
        arguments = [Increment]
        return self.IncrementRotationZ(*arguments)

    def ResetModel(self, ResetSize=None):
        arguments = com_arguments([unwrap(a) for a in [ResetSize]])
        self.com_object.ResetModel(*arguments)

    # Lower case alias for ResetModel
    def resetmodel(self, ResetSize=None):
        arguments = [ResetSize]
        return self.ResetModel(*arguments)


class MotionEffect:

    def __init__(self, motioneffect=None):
        self.com_object= motioneffect

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def ByX(self):
        return self.com_object.ByX

    @ByX.setter
    def ByX(self, value):
        self.com_object.ByX = value

    # Lower case aliases for ByX
    @property
    def byx(self):
        return self.ByX

    @byx.setter
    def byx(self, value):
        self.ByX = value

    @property
    def ByY(self):
        return self.com_object.ByY

    @ByY.setter
    def ByY(self, value):
        self.com_object.ByY = value

    # Lower case aliases for ByY
    @property
    def byy(self):
        return self.ByY

    @byy.setter
    def byy(self, value):
        self.ByY = value

    @property
    def FromX(self):
        return self.com_object.FromX

    @FromX.setter
    def FromX(self, value):
        self.com_object.FromX = value

    # Lower case aliases for FromX
    @property
    def fromx(self):
        return self.FromX

    @fromx.setter
    def fromx(self, value):
        self.FromX = value

    @property
    def FromY(self):
        return MotionEffect(self.com_object.FromY)

    @FromY.setter
    def FromY(self, value):
        self.com_object.FromY = value

    # Lower case aliases for FromY
    @property
    def fromy(self):
        return self.FromY

    @fromy.setter
    def fromy(self, value):
        self.FromY = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Path(self):
        return self.com_object.Path

    @Path.setter
    def Path(self, value):
        self.com_object.Path = value

    # Lower case aliases for Path
    @property
    def path(self):
        return self.Path

    @path.setter
    def path(self, value):
        self.Path = value

    @property
    def ToX(self):
        return self.com_object.ToX

    @ToX.setter
    def ToX(self, value):
        self.com_object.ToX = value

    # Lower case aliases for ToX
    @property
    def tox(self):
        return self.ToX

    @tox.setter
    def tox(self, value):
        self.ToX = value

    @property
    def ToY(self):
        return MotionEffect(self.com_object.ToY)

    @ToY.setter
    def ToY(self, value):
        self.com_object.ToY = value

    # Lower case aliases for ToY
    @property
    def toy(self):
        return self.ToY

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
        self.com_object= namedslideshow

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

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
    def SlideIDs(self):
        return self.com_object.SlideIDs

    # Lower case aliases for SlideIDs
    @property
    def slideids(self):
        return self.SlideIDs

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()


class NamedSlideShows:

    def __init__(self, namedslideshows=None):
        self.com_object= namedslideshows

    def __call__(self, item):
        return NamedSlideShow(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Add(self, Name=None, SafeArrayOfSlideIDs=None):
        arguments = com_arguments([unwrap(a) for a in [Name, SafeArrayOfSlideIDs]])
        return NamedSlideShow(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Name=None, SafeArrayOfSlideIDs=None):
        arguments = [Name, SafeArrayOfSlideIDs]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return NamedSlideShow(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class ObjectVerbs:

    def __init__(self, objectverbs=None):
        self.com_object= objectverbs

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return self.com_object.Item(*arguments)

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class OLEFormat:

    def __init__(self, oleformat=None):
        self.com_object= oleformat

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def FollowColors(self):
        return self.com_object.FollowColors

    @FollowColors.setter
    def FollowColors(self, value):
        self.com_object.FollowColors = value

    # Lower case aliases for FollowColors
    @property
    def followcolors(self):
        return self.FollowColors

    @followcolors.setter
    def followcolors(self, value):
        self.FollowColors = value

    @property
    def Object(self):
        return self.com_object.Object

    # Lower case aliases for Object
    @property
    def object(self):
        return self.Object

    @property
    def ObjectVerbs(self):
        return ObjectVerbs(self.com_object.ObjectVerbs)

    # Lower case aliases for ObjectVerbs
    @property
    def objectverbs(self):
        return self.ObjectVerbs

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ProgID(self):
        return self.com_object.ProgID

    # Lower case aliases for ProgID
    @property
    def progid(self):
        return self.ProgID

    def Activate(self):
        self.com_object.Activate()

    # Lower case alias for Activate
    def activate(self):
        return self.Activate()

    def DoVerb(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        self.com_object.DoVerb(*arguments)

    # Lower case alias for DoVerb
    def doverb(self, Index=None):
        arguments = [Index]
        return self.DoVerb(*arguments)


class Options:

    def __init__(self, options=None):
        self.com_object= options

    @property
    def DisplayPasteOptions(self):
        return self.com_object.DisplayPasteOptions

    @DisplayPasteOptions.setter
    def DisplayPasteOptions(self, value):
        self.com_object.DisplayPasteOptions = value

    # Lower case aliases for DisplayPasteOptions
    @property
    def displaypasteoptions(self):
        return self.DisplayPasteOptions

    @displaypasteoptions.setter
    def displaypasteoptions(self, value):
        self.DisplayPasteOptions = value

    @property
    def ShowCoauthoringMergeChanges(self):
        return self.com_object.ShowCoauthoringMergeChanges

    # Lower case aliases for ShowCoauthoringMergeChanges
    @property
    def showcoauthoringmergechanges(self):
        return self.ShowCoauthoringMergeChanges


class PageSetup:

    def __init__(self, pagesetup=None):
        self.com_object= pagesetup

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def FirstSlideNumber(self):
        return self.com_object.FirstSlideNumber

    @FirstSlideNumber.setter
    def FirstSlideNumber(self, value):
        self.com_object.FirstSlideNumber = value

    # Lower case aliases for FirstSlideNumber
    @property
    def firstslidenumber(self):
        return self.FirstSlideNumber

    @firstslidenumber.setter
    def firstslidenumber(self, value):
        self.FirstSlideNumber = value

    @property
    def NotesOrientation(self):
        return self.com_object.NotesOrientation

    @NotesOrientation.setter
    def NotesOrientation(self, value):
        self.com_object.NotesOrientation = value

    # Lower case aliases for NotesOrientation
    @property
    def notesorientation(self):
        return self.NotesOrientation

    @notesorientation.setter
    def notesorientation(self, value):
        self.NotesOrientation = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SlideHeight(self):
        return self.com_object.SlideHeight

    @SlideHeight.setter
    def SlideHeight(self, value):
        self.com_object.SlideHeight = value

    # Lower case aliases for SlideHeight
    @property
    def slideheight(self):
        return self.SlideHeight

    @slideheight.setter
    def slideheight(self, value):
        self.SlideHeight = value

    @property
    def SlideOrientation(self):
        return self.com_object.SlideOrientation

    @SlideOrientation.setter
    def SlideOrientation(self, value):
        self.com_object.SlideOrientation = value

    # Lower case aliases for SlideOrientation
    @property
    def slideorientation(self):
        return self.SlideOrientation

    @slideorientation.setter
    def slideorientation(self, value):
        self.SlideOrientation = value

    @property
    def SlideSize(self):
        return self.com_object.SlideSize

    @SlideSize.setter
    def SlideSize(self, value):
        self.com_object.SlideSize = value

    # Lower case aliases for SlideSize
    @property
    def slidesize(self):
        return self.SlideSize

    @slidesize.setter
    def slidesize(self, value):
        self.SlideSize = value

    @property
    def SlideWidth(self):
        return self.com_object.SlideWidth

    @SlideWidth.setter
    def SlideWidth(self, value):
        self.com_object.SlideWidth = value

    # Lower case aliases for SlideWidth
    @property
    def slidewidth(self):
        return self.SlideWidth

    @slidewidth.setter
    def slidewidth(self, value):
        self.SlideWidth = value


class Pane:

    def __init__(self, pane=None):
        self.com_object= pane

    @property
    def Active(self):
        return self.com_object.Active

    # Lower case aliases for Active
    @property
    def active(self):
        return self.Active

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ViewType(self):
        return self.com_object.ViewType

    # Lower case aliases for ViewType
    @property
    def viewtype(self):
        return self.ViewType

    def Activate(self):
        self.com_object.Activate()

    # Lower case alias for Activate
    def activate(self):
        return self.Activate()


class Panes:

    def __init__(self, panes=None):
        self.com_object= panes

    def __call__(self, item):
        return Pane(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Pane(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class ParagraphFormat:

    def __init__(self, paragraphformat=None):
        self.com_object= paragraphformat

    @property
    def Alignment(self):
        return self.com_object.Alignment

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
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def BaseLineAlignment(self):
        return self.com_object.BaseLineAlignment

    @BaseLineAlignment.setter
    def BaseLineAlignment(self, value):
        self.com_object.BaseLineAlignment = value

    # Lower case aliases for BaseLineAlignment
    @property
    def baselinealignment(self):
        return self.BaseLineAlignment

    @baselinealignment.setter
    def baselinealignment(self, value):
        self.BaseLineAlignment = value

    @property
    def Bullet(self):
        return BulletFormat(self.com_object.Bullet)

    # Lower case aliases for Bullet
    @property
    def bullet(self):
        return self.Bullet

    @property
    def FarEastLineBreakControl(self):
        return self.com_object.FarEastLineBreakControl

    @FarEastLineBreakControl.setter
    def FarEastLineBreakControl(self, value):
        self.com_object.FarEastLineBreakControl = value

    # Lower case aliases for FarEastLineBreakControl
    @property
    def fareastlinebreakcontrol(self):
        return self.FarEastLineBreakControl

    @fareastlinebreakcontrol.setter
    def fareastlinebreakcontrol(self, value):
        self.FarEastLineBreakControl = value

    @property
    def HangingPunctuation(self):
        return self.com_object.HangingPunctuation

    @HangingPunctuation.setter
    def HangingPunctuation(self, value):
        self.com_object.HangingPunctuation = value

    # Lower case aliases for HangingPunctuation
    @property
    def hangingpunctuation(self):
        return self.HangingPunctuation

    @hangingpunctuation.setter
    def hangingpunctuation(self, value):
        self.HangingPunctuation = value

    @property
    def LineRuleAfter(self):
        return self.com_object.LineRuleAfter

    @LineRuleAfter.setter
    def LineRuleAfter(self, value):
        self.com_object.LineRuleAfter = value

    # Lower case aliases for LineRuleAfter
    @property
    def lineruleafter(self):
        return self.LineRuleAfter

    @lineruleafter.setter
    def lineruleafter(self, value):
        self.LineRuleAfter = value

    @property
    def LineRuleBefore(self):
        return self.com_object.LineRuleBefore

    @LineRuleBefore.setter
    def LineRuleBefore(self, value):
        self.com_object.LineRuleBefore = value

    # Lower case aliases for LineRuleBefore
    @property
    def linerulebefore(self):
        return self.LineRuleBefore

    @linerulebefore.setter
    def linerulebefore(self, value):
        self.LineRuleBefore = value

    @property
    def LineRuleWithin(self):
        return self.com_object.LineRuleWithin

    @LineRuleWithin.setter
    def LineRuleWithin(self, value):
        self.com_object.LineRuleWithin = value

    # Lower case aliases for LineRuleWithin
    @property
    def linerulewithin(self):
        return self.LineRuleWithin

    @linerulewithin.setter
    def linerulewithin(self, value):
        self.LineRuleWithin = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SpaceAfter(self):
        return self.com_object.SpaceAfter

    @SpaceAfter.setter
    def SpaceAfter(self, value):
        self.com_object.SpaceAfter = value

    # Lower case aliases for SpaceAfter
    @property
    def spaceafter(self):
        return self.SpaceAfter

    @spaceafter.setter
    def spaceafter(self, value):
        self.SpaceAfter = value

    @property
    def SpaceBefore(self):
        return self.com_object.SpaceBefore

    @SpaceBefore.setter
    def SpaceBefore(self, value):
        self.com_object.SpaceBefore = value

    # Lower case aliases for SpaceBefore
    @property
    def spacebefore(self):
        return self.SpaceBefore

    @spacebefore.setter
    def spacebefore(self, value):
        self.SpaceBefore = value

    @property
    def SpaceWithin(self):
        return self.com_object.SpaceWithin

    @SpaceWithin.setter
    def SpaceWithin(self, value):
        self.com_object.SpaceWithin = value

    # Lower case aliases for SpaceWithin
    @property
    def spacewithin(self):
        return self.SpaceWithin

    @spacewithin.setter
    def spacewithin(self, value):
        self.SpaceWithin = value

    @property
    def TextDirection(self):
        return self.com_object.TextDirection

    @TextDirection.setter
    def TextDirection(self, value):
        self.com_object.TextDirection = value

    # Lower case aliases for TextDirection
    @property
    def textdirection(self):
        return self.TextDirection

    @textdirection.setter
    def textdirection(self, value):
        self.TextDirection = value

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


class PictureFormat:

    def __init__(self, pictureformat=None):
        self.com_object= pictureformat

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Brightness(self):
        return self.com_object.Brightness

    @Brightness.setter
    def Brightness(self, value):
        self.com_object.Brightness = value

    # Lower case aliases for Brightness
    @property
    def brightness(self):
        return self.Brightness

    @brightness.setter
    def brightness(self, value):
        self.Brightness = value

    @property
    def ColorType(self):
        return self.com_object.ColorType

    @ColorType.setter
    def ColorType(self, value):
        self.com_object.ColorType = value

    # Lower case aliases for ColorType
    @property
    def colortype(self):
        return self.ColorType

    @colortype.setter
    def colortype(self, value):
        self.ColorType = value

    @property
    def Contrast(self):
        return self.com_object.Contrast

    @Contrast.setter
    def Contrast(self, value):
        self.com_object.Contrast = value

    # Lower case aliases for Contrast
    @property
    def contrast(self):
        return self.Contrast

    @contrast.setter
    def contrast(self, value):
        self.Contrast = value

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Crop(self):
        return self.com_object.Crop

    @Crop.setter
    def Crop(self, value):
        self.com_object.Crop = value

    # Lower case aliases for Crop
    @property
    def crop(self):
        return self.Crop

    @crop.setter
    def crop(self, value):
        self.Crop = value

    @property
    def CropBottom(self):
        return self.com_object.CropBottom

    @CropBottom.setter
    def CropBottom(self, value):
        self.com_object.CropBottom = value

    # Lower case aliases for CropBottom
    @property
    def cropbottom(self):
        return self.CropBottom

    @cropbottom.setter
    def cropbottom(self, value):
        self.CropBottom = value

    @property
    def CropLeft(self):
        return self.com_object.CropLeft

    @CropLeft.setter
    def CropLeft(self, value):
        self.com_object.CropLeft = value

    # Lower case aliases for CropLeft
    @property
    def cropleft(self):
        return self.CropLeft

    @cropleft.setter
    def cropleft(self, value):
        self.CropLeft = value

    @property
    def CropRight(self):
        return self.com_object.CropRight

    @CropRight.setter
    def CropRight(self, value):
        self.com_object.CropRight = value

    # Lower case aliases for CropRight
    @property
    def cropright(self):
        return self.CropRight

    @cropright.setter
    def cropright(self, value):
        self.CropRight = value

    @property
    def CropTop(self):
        return self.com_object.CropTop

    @CropTop.setter
    def CropTop(self, value):
        self.com_object.CropTop = value

    # Lower case aliases for CropTop
    @property
    def croptop(self):
        return self.CropTop

    @croptop.setter
    def croptop(self, value):
        self.CropTop = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def TransparencyColor(self):
        return self.com_object.TransparencyColor

    @TransparencyColor.setter
    def TransparencyColor(self, value):
        self.com_object.TransparencyColor = value

    # Lower case aliases for TransparencyColor
    @property
    def transparencycolor(self):
        return self.TransparencyColor

    @transparencycolor.setter
    def transparencycolor(self, value):
        self.TransparencyColor = value

    @property
    def TransparentBackground(self):
        return self.com_object.TransparentBackground

    @TransparentBackground.setter
    def TransparentBackground(self, value):
        self.com_object.TransparentBackground = value

    # Lower case aliases for TransparentBackground
    @property
    def transparentbackground(self):
        return self.TransparentBackground

    @transparentbackground.setter
    def transparentbackground(self, value):
        self.TransparentBackground = value

    def IncrementBrightness(self, Increment=None):
        arguments = com_arguments([unwrap(a) for a in [Increment]])
        self.com_object.IncrementBrightness(*arguments)

    # Lower case alias for IncrementBrightness
    def incrementbrightness(self, Increment=None):
        arguments = [Increment]
        return self.IncrementBrightness(*arguments)

    def IncrementContrast(self, Increment=None):
        arguments = com_arguments([unwrap(a) for a in [Increment]])
        self.com_object.IncrementContrast(*arguments)

    # Lower case alias for IncrementContrast
    def incrementcontrast(self, Increment=None):
        arguments = [Increment]
        return self.IncrementContrast(*arguments)


class PlaceholderFormat:

    def __init__(self, placeholderformat=None):
        self.com_object= placeholderformat

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def ContainedType(self):
        return self.com_object.ContainedType

    # Lower case aliases for ContainedType
    @property
    def containedtype(self):
        return self.ContainedType

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
    def Type(self):
        return self.com_object.Type

    # Lower case aliases for Type
    @property
    def type(self):
        return self.Type


class Placeholders:

    def __init__(self, placeholders=None):
        self.com_object= placeholders

    def __call__(self, item):
        return Placeholder(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def FindByName(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Shape(self.com_object.FindByName(*arguments))

    # Lower case alias for FindByName
    def findbyname(self, Index=None):
        arguments = [Index]
        return self.FindByName(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Shape(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class Player:

    def __init__(self, player=None):
        self.com_object= player

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def CurrentPosition(self):
        return self.com_object.CurrentPosition

    @CurrentPosition.setter
    def CurrentPosition(self, value):
        self.com_object.CurrentPosition = value

    # Lower case aliases for CurrentPosition
    @property
    def currentposition(self):
        return self.CurrentPosition

    @currentposition.setter
    def currentposition(self, value):
        self.CurrentPosition = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def State(self):
        return self.com_object.State

    # Lower case aliases for State
    @property
    def state(self):
        return self.State

    def GoToNextBookmark(self):
        self.com_object.GoToNextBookmark()

    # Lower case alias for GoToNextBookmark
    def gotonextbookmark(self):
        return self.GoToNextBookmark()

    def GoToPreviousBookmark(self):
        self.com_object.GoToPreviousBookmark()

    # Lower case alias for GoToPreviousBookmark
    def gotopreviousbookmark(self):
        return self.GoToPreviousBookmark()

    def Pause(self):
        self.com_object.Pause()

    # Lower case alias for Pause
    def pause(self):
        return self.Pause()

    def Stop(self):
        self.com_object.Stop()

    # Lower case alias for Stop
    def stop(self):
        return self.Stop()


class PlaySettings:

    def __init__(self, playsettings=None):
        self.com_object= playsettings

    @property
    def ActionVerb(self):
        return self.com_object.ActionVerb

    @ActionVerb.setter
    def ActionVerb(self, value):
        self.com_object.ActionVerb = value

    # Lower case aliases for ActionVerb
    @property
    def actionverb(self):
        return self.ActionVerb

    @actionverb.setter
    def actionverb(self, value):
        self.ActionVerb = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def HideWhileNotPlaying(self):
        return self.com_object.HideWhileNotPlaying

    @HideWhileNotPlaying.setter
    def HideWhileNotPlaying(self, value):
        self.com_object.HideWhileNotPlaying = value

    # Lower case aliases for HideWhileNotPlaying
    @property
    def hidewhilenotplaying(self):
        return self.HideWhileNotPlaying

    @hidewhilenotplaying.setter
    def hidewhilenotplaying(self, value):
        self.HideWhileNotPlaying = value

    @property
    def LoopUntilStopped(self):
        return self.com_object.LoopUntilStopped

    @LoopUntilStopped.setter
    def LoopUntilStopped(self, value):
        self.com_object.LoopUntilStopped = value

    # Lower case aliases for LoopUntilStopped
    @property
    def loopuntilstopped(self):
        return self.LoopUntilStopped

    @loopuntilstopped.setter
    def loopuntilstopped(self, value):
        self.LoopUntilStopped = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PauseAnimation(self):
        return self.com_object.PauseAnimation

    @PauseAnimation.setter
    def PauseAnimation(self, value):
        self.com_object.PauseAnimation = value

    # Lower case aliases for PauseAnimation
    @property
    def pauseanimation(self):
        return self.PauseAnimation

    @pauseanimation.setter
    def pauseanimation(self, value):
        self.PauseAnimation = value

    @property
    def PlayOnEntry(self):
        return self.com_object.PlayOnEntry

    @PlayOnEntry.setter
    def PlayOnEntry(self, value):
        self.com_object.PlayOnEntry = value

    # Lower case aliases for PlayOnEntry
    @property
    def playonentry(self):
        return self.PlayOnEntry

    @playonentry.setter
    def playonentry(self, value):
        self.PlayOnEntry = value

    @property
    def RewindMovie(self):
        return self.com_object.RewindMovie

    @RewindMovie.setter
    def RewindMovie(self, value):
        self.com_object.RewindMovie = value

    # Lower case aliases for RewindMovie
    @property
    def rewindmovie(self):
        return self.RewindMovie

    @rewindmovie.setter
    def rewindmovie(self, value):
        self.RewindMovie = value

    @property
    def StopAfterSlides(self):
        return self.com_object.StopAfterSlides

    @StopAfterSlides.setter
    def StopAfterSlides(self, value):
        self.com_object.StopAfterSlides = value

    # Lower case aliases for StopAfterSlides
    @property
    def stopafterslides(self):
        return self.StopAfterSlides

    @stopafterslides.setter
    def stopafterslides(self, value):
        self.StopAfterSlides = value


class PlotArea:

    def __init__(self, plotarea=None):
        self.com_object= plotarea

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

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
    def InsideHeight(self):
        return self.com_object.InsideHeight

    @InsideHeight.setter
    def InsideHeight(self, value):
        self.com_object.InsideHeight = value

    # Lower case aliases for InsideHeight
    @property
    def insideheight(self):
        return self.InsideHeight

    @insideheight.setter
    def insideheight(self, value):
        self.InsideHeight = value

    @property
    def InsideLeft(self):
        return self.com_object.InsideLeft

    @InsideLeft.setter
    def InsideLeft(self, value):
        self.com_object.InsideLeft = value

    # Lower case aliases for InsideLeft
    @property
    def insideleft(self):
        return self.InsideLeft

    @insideleft.setter
    def insideleft(self, value):
        self.InsideLeft = value

    @property
    def InsideTop(self):
        return self.com_object.InsideTop

    @InsideTop.setter
    def InsideTop(self, value):
        self.com_object.InsideTop = value

    # Lower case aliases for InsideTop
    @property
    def insidetop(self):
        return self.InsideTop

    @insidetop.setter
    def insidetop(self, value):
        self.InsideTop = value

    @property
    def InsideWidth(self):
        return self.com_object.InsideWidth

    @InsideWidth.setter
    def InsideWidth(self, value):
        self.com_object.InsideWidth = value

    # Lower case aliases for InsideWidth
    @property
    def insidewidth(self):
        return self.InsideWidth

    @insidewidth.setter
    def insidewidth(self, value):
        self.InsideWidth = value

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
    def Position(self):
        return XlChartElementPosition(self.com_object.Position)

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

    def ClearFormats(self):
        self.com_object.ClearFormats()

    # Lower case alias for ClearFormats
    def clearformats(self):
        return self.ClearFormats()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class Point:

    def __init__(self, point=None):
        self.com_object= point

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def ApplyPictToEnd(self):
        return self.com_object.ApplyPictToEnd

    @ApplyPictToEnd.setter
    def ApplyPictToEnd(self, value):
        self.com_object.ApplyPictToEnd = value

    # Lower case aliases for ApplyPictToEnd
    @property
    def applypicttoend(self):
        return self.ApplyPictToEnd

    @applypicttoend.setter
    def applypicttoend(self, value):
        self.ApplyPictToEnd = value

    @property
    def ApplyPictToFront(self):
        return self.com_object.ApplyPictToFront

    @ApplyPictToFront.setter
    def ApplyPictToFront(self, value):
        self.com_object.ApplyPictToFront = value

    # Lower case aliases for ApplyPictToFront
    @property
    def applypicttofront(self):
        return self.ApplyPictToFront

    @applypicttofront.setter
    def applypicttofront(self, value):
        self.ApplyPictToFront = value

    @property
    def ApplyPictToSides(self):
        return self.com_object.ApplyPictToSides

    @ApplyPictToSides.setter
    def ApplyPictToSides(self, value):
        self.com_object.ApplyPictToSides = value

    # Lower case aliases for ApplyPictToSides
    @property
    def applypicttosides(self):
        return self.ApplyPictToSides

    @applypicttosides.setter
    def applypicttosides(self, value):
        self.ApplyPictToSides = value

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def DataLabel(self):
        return DataLabel(self.com_object.DataLabel)

    # Lower case aliases for DataLabel
    @property
    def datalabel(self):
        return self.DataLabel

    @property
    def Explosion(self):
        return self.com_object.Explosion

    @Explosion.setter
    def Explosion(self, value):
        self.com_object.Explosion = value

    # Lower case aliases for Explosion
    @property
    def explosion(self):
        return self.Explosion

    @explosion.setter
    def explosion(self, value):
        self.Explosion = value

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

    @property
    def Has3DEffect(self):
        return self.com_object.Has3DEffect

    @Has3DEffect.setter
    def Has3DEffect(self, value):
        self.com_object.Has3DEffect = value

    # Lower case aliases for Has3DEffect
    @property
    def has3deffect(self):
        return self.Has3DEffect

    @has3deffect.setter
    def has3deffect(self, value):
        self.Has3DEffect = value

    @property
    def HasDataLabel(self):
        return self.com_object.HasDataLabel

    @HasDataLabel.setter
    def HasDataLabel(self, value):
        self.com_object.HasDataLabel = value

    # Lower case aliases for HasDataLabel
    @property
    def hasdatalabel(self):
        return self.HasDataLabel

    @hasdatalabel.setter
    def hasdatalabel(self, value):
        self.HasDataLabel = value

    @property
    def Height(self):
        return self.com_object.Height

    # Lower case aliases for Height
    @property
    def height(self):
        return self.Height

    @property
    def InvertIfNegative(self):
        return self.com_object.InvertIfNegative

    @InvertIfNegative.setter
    def InvertIfNegative(self, value):
        self.com_object.InvertIfNegative = value

    # Lower case aliases for InvertIfNegative
    @property
    def invertifnegative(self):
        return self.InvertIfNegative

    @invertifnegative.setter
    def invertifnegative(self, value):
        self.InvertIfNegative = value

    @property
    def Left(self):
        return self.com_object.Left

    # Lower case aliases for Left
    @property
    def left(self):
        return self.Left

    @property
    def MarkerBackgroundColor(self):
        return self.com_object.MarkerBackgroundColor

    @MarkerBackgroundColor.setter
    def MarkerBackgroundColor(self, value):
        self.com_object.MarkerBackgroundColor = value

    # Lower case aliases for MarkerBackgroundColor
    @property
    def markerbackgroundcolor(self):
        return self.MarkerBackgroundColor

    @markerbackgroundcolor.setter
    def markerbackgroundcolor(self, value):
        self.MarkerBackgroundColor = value

    @property
    def MarkerBackgroundColorIndex(self):
        return XlColorIndex(self.com_object.MarkerBackgroundColorIndex)

    @MarkerBackgroundColorIndex.setter
    def MarkerBackgroundColorIndex(self, value):
        self.com_object.MarkerBackgroundColorIndex = value

    # Lower case aliases for MarkerBackgroundColorIndex
    @property
    def markerbackgroundcolorindex(self):
        return self.MarkerBackgroundColorIndex

    @markerbackgroundcolorindex.setter
    def markerbackgroundcolorindex(self, value):
        self.MarkerBackgroundColorIndex = value

    @property
    def MarkerForegroundColor(self):
        return self.com_object.MarkerForegroundColor

    @MarkerForegroundColor.setter
    def MarkerForegroundColor(self, value):
        self.com_object.MarkerForegroundColor = value

    # Lower case aliases for MarkerForegroundColor
    @property
    def markerforegroundcolor(self):
        return self.MarkerForegroundColor

    @markerforegroundcolor.setter
    def markerforegroundcolor(self, value):
        self.MarkerForegroundColor = value

    @property
    def MarkerForegroundColorIndex(self):
        return XlColorIndex(self.com_object.MarkerForegroundColorIndex)

    @MarkerForegroundColorIndex.setter
    def MarkerForegroundColorIndex(self, value):
        self.com_object.MarkerForegroundColorIndex = value

    # Lower case aliases for MarkerForegroundColorIndex
    @property
    def markerforegroundcolorindex(self):
        return self.MarkerForegroundColorIndex

    @markerforegroundcolorindex.setter
    def markerforegroundcolorindex(self, value):
        self.MarkerForegroundColorIndex = value

    @property
    def MarkerSize(self):
        return self.com_object.MarkerSize

    @MarkerSize.setter
    def MarkerSize(self, value):
        self.com_object.MarkerSize = value

    # Lower case aliases for MarkerSize
    @property
    def markersize(self):
        return self.MarkerSize

    @markersize.setter
    def markersize(self, value):
        self.MarkerSize = value

    @property
    def MarkerStyle(self):
        return XlMarkerStyle(self.com_object.MarkerStyle)

    @MarkerStyle.setter
    def MarkerStyle(self, value):
        self.com_object.MarkerStyle = value

    # Lower case aliases for MarkerStyle
    @property
    def markerstyle(self):
        return self.MarkerStyle

    @markerstyle.setter
    def markerstyle(self, value):
        self.MarkerStyle = value

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
    def PictureType(self):
        return XlChartPictureType(self.com_object.PictureType)

    @PictureType.setter
    def PictureType(self, value):
        self.com_object.PictureType = value

    # Lower case aliases for PictureType
    @property
    def picturetype(self):
        return self.PictureType

    @picturetype.setter
    def picturetype(self, value):
        self.PictureType = value

    @property
    def PictureUnit2(self):
        return self.com_object.PictureUnit2

    @PictureUnit2.setter
    def PictureUnit2(self, value):
        self.com_object.PictureUnit2 = value

    # Lower case aliases for PictureUnit2
    @property
    def pictureunit2(self):
        return self.PictureUnit2

    @pictureunit2.setter
    def pictureunit2(self, value):
        self.PictureUnit2 = value

    @property
    def SecondaryPlot(self):
        return self.com_object.SecondaryPlot

    @SecondaryPlot.setter
    def SecondaryPlot(self, value):
        self.com_object.SecondaryPlot = value

    # Lower case aliases for SecondaryPlot
    @property
    def secondaryplot(self):
        return self.SecondaryPlot

    @secondaryplot.setter
    def secondaryplot(self, value):
        self.SecondaryPlot = value

    @property
    def Shadow(self):
        return self.com_object.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.com_object.Shadow = value

    # Lower case aliases for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @shadow.setter
    def shadow(self, value):
        self.Shadow = value

    @property
    def Top(self):
        return self.com_object.Top

    # Lower case aliases for Top
    @property
    def top(self):
        return self.Top

    @property
    def Width(self):
        return self.com_object.Width

    # Lower case aliases for Width
    @property
    def width(self):
        return self.Width

    def ApplyDataLabels(self, Type=None, LegendKey=None, AutoText=None, HasLeaderLines=None, ShowSeriesName=None, ShowCategoryName=None, ShowValue=None, ShowPercentage=None, ShowBubbleSize=None, Separator=None):
        arguments = com_arguments([unwrap(a) for a in [Type, LegendKey, AutoText, HasLeaderLines, ShowSeriesName, ShowCategoryName, ShowValue, ShowPercentage, ShowBubbleSize, Separator]])
        self.com_object.ApplyDataLabels(*arguments)

    # Lower case alias for ApplyDataLabels
    def applydatalabels(self, Type=None, LegendKey=None, AutoText=None, HasLeaderLines=None, ShowSeriesName=None, ShowCategoryName=None, ShowValue=None, ShowPercentage=None, ShowBubbleSize=None, Separator=None):
        arguments = [Type, LegendKey, AutoText, HasLeaderLines, ShowSeriesName, ShowCategoryName, ShowValue, ShowPercentage, ShowBubbleSize, Separator]
        return self.ApplyDataLabels(*arguments)

    def ClearFormats(self):
        self.com_object.ClearFormats()

    # Lower case alias for ClearFormats
    def clearformats(self):
        return self.ClearFormats()

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

    def Paste(self):
        self.com_object.Paste()

    # Lower case alias for Paste
    def paste(self):
        return self.Paste()

    def PieSliceLocation(self, loc=None, Index=None):
        arguments = com_arguments([unwrap(a) for a in [loc, Index]])
        return self.com_object.PieSliceLocation(*arguments)

    # Lower case alias for PieSliceLocation
    def pieslicelocation(self, loc=None, Index=None):
        arguments = [loc, Index]
        return self.PieSliceLocation(*arguments)

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class Points:

    def __init__(self, points=None):
        self.com_object= points

    def __call__(self, item):
        return Point(self.com_object(item))

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Point(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


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
        self.com_object= presentation

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AutoSaveOn(self):
        return self.com_object.AutoSaveOn

    @AutoSaveOn.setter
    def AutoSaveOn(self, value):
        self.com_object.AutoSaveOn = value

    # Lower case aliases for AutoSaveOn
    @property
    def autosaveon(self):
        return self.AutoSaveOn

    @autosaveon.setter
    def autosaveon(self, value):
        self.AutoSaveOn = value

    @property
    def Broadcast(self):
        return Broadcast(self.com_object.Broadcast)

    # Lower case aliases for Broadcast
    @property
    def broadcast(self):
        return self.Broadcast

    @property
    def BuiltInDocumentProperties(self):
        return self.com_object.BuiltInDocumentProperties

    # Lower case aliases for BuiltInDocumentProperties
    @property
    def builtindocumentproperties(self):
        return self.BuiltInDocumentProperties

    @property
    def Coauthoring(self):
        return Coauthoring(self.com_object.Coauthoring)

    # Lower case aliases for Coauthoring
    @property
    def coauthoring(self):
        return self.Coauthoring

    @property
    def ColorSchemes(self):
        return ColorSchemes(self.com_object.ColorSchemes)

    # Lower case aliases for ColorSchemes
    @property
    def colorschemes(self):
        return self.ColorSchemes

    @property
    def CommandBars(self):
        return self.com_object.CommandBars

    # Lower case aliases for CommandBars
    @property
    def commandbars(self):
        return self.CommandBars

    @property
    def Container(self):
        return self.com_object.Container

    # Lower case aliases for Container
    @property
    def container(self):
        return self.Container

    @property
    def ContentTypeProperties(self):
        return self.com_object.ContentTypeProperties

    # Lower case aliases for ContentTypeProperties
    @property
    def contenttypeproperties(self):
        return self.ContentTypeProperties

    @property
    def CreateVideoStatus(self):
        return Presentation(self.com_object.CreateVideoStatus)

    # Lower case aliases for CreateVideoStatus
    @property
    def createvideostatus(self):
        return self.CreateVideoStatus

    @property
    def CustomDocumentProperties(self):
        return self.com_object.CustomDocumentProperties

    # Lower case aliases for CustomDocumentProperties
    @property
    def customdocumentproperties(self):
        return self.CustomDocumentProperties

    @property
    def CustomerData(self):
        return CustomerData(self.com_object.CustomerData)

    # Lower case aliases for CustomerData
    @property
    def customerdata(self):
        return self.CustomerData

    @property
    def CustomXMLParts(self):
        return self.com_object.CustomXMLParts

    # Lower case aliases for CustomXMLParts
    @property
    def customxmlparts(self):
        return self.CustomXMLParts

    @property
    def DefaultLanguageID(self):
        return self.com_object.DefaultLanguageID

    @DefaultLanguageID.setter
    def DefaultLanguageID(self, value):
        self.com_object.DefaultLanguageID = value

    # Lower case aliases for DefaultLanguageID
    @property
    def defaultlanguageid(self):
        return self.DefaultLanguageID

    @defaultlanguageid.setter
    def defaultlanguageid(self, value):
        self.DefaultLanguageID = value

    @property
    def DefaultShape(self):
        return Shape(self.com_object.DefaultShape)

    # Lower case aliases for DefaultShape
    @property
    def defaultshape(self):
        return self.DefaultShape

    @property
    def Designs(self):
        return Designs(self.com_object.Designs)

    # Lower case aliases for Designs
    @property
    def designs(self):
        return self.Designs

    @property
    def DisplayComments(self):
        return self.com_object.DisplayComments

    @DisplayComments.setter
    def DisplayComments(self, value):
        self.com_object.DisplayComments = value

    # Lower case aliases for DisplayComments
    @property
    def displaycomments(self):
        return self.DisplayComments

    @displaycomments.setter
    def displaycomments(self, value):
        self.DisplayComments = value

    @property
    def DocumentInspectors(self):
        return self.com_object.DocumentInspectors

    # Lower case aliases for DocumentInspectors
    @property
    def documentinspectors(self):
        return self.DocumentInspectors

    @property
    def DocumentLibraryVersions(self):
        return self.com_object.DocumentLibraryVersions

    # Lower case aliases for DocumentLibraryVersions
    @property
    def documentlibraryversions(self):
        return self.DocumentLibraryVersions

    @property
    def EncryptionProvider(self):
        return self.com_object.EncryptionProvider

    @EncryptionProvider.setter
    def EncryptionProvider(self, value):
        self.com_object.EncryptionProvider = value

    # Lower case aliases for EncryptionProvider
    @property
    def encryptionprovider(self):
        return self.EncryptionProvider

    @encryptionprovider.setter
    def encryptionprovider(self, value):
        self.EncryptionProvider = value

    @property
    def EnvelopeVisible(self):
        return self.com_object.EnvelopeVisible

    @EnvelopeVisible.setter
    def EnvelopeVisible(self, value):
        self.com_object.EnvelopeVisible = value

    # Lower case aliases for EnvelopeVisible
    @property
    def envelopevisible(self):
        return self.EnvelopeVisible

    @envelopevisible.setter
    def envelopevisible(self, value):
        self.EnvelopeVisible = value

    @property
    def ExtraColors(self):
        return ExtraColors(self.com_object.ExtraColors)

    # Lower case aliases for ExtraColors
    @property
    def extracolors(self):
        return self.ExtraColors

    @property
    def FarEastLineBreakLanguage(self):
        return self.com_object.FarEastLineBreakLanguage

    @FarEastLineBreakLanguage.setter
    def FarEastLineBreakLanguage(self, value):
        self.com_object.FarEastLineBreakLanguage = value

    # Lower case aliases for FarEastLineBreakLanguage
    @property
    def fareastlinebreaklanguage(self):
        return self.FarEastLineBreakLanguage

    @fareastlinebreaklanguage.setter
    def fareastlinebreaklanguage(self, value):
        self.FarEastLineBreakLanguage = value

    @property
    def FarEastLineBreakLevel(self):
        return self.com_object.FarEastLineBreakLevel

    @FarEastLineBreakLevel.setter
    def FarEastLineBreakLevel(self, value):
        self.com_object.FarEastLineBreakLevel = value

    # Lower case aliases for FarEastLineBreakLevel
    @property
    def fareastlinebreaklevel(self):
        return self.FarEastLineBreakLevel

    @fareastlinebreaklevel.setter
    def fareastlinebreaklevel(self, value):
        self.FarEastLineBreakLevel = value

    @property
    def Final(self):
        return self.com_object.Final

    @Final.setter
    def Final(self, value):
        self.com_object.Final = value

    # Lower case aliases for Final
    @property
    def final(self):
        return self.Final

    @final.setter
    def final(self, value):
        self.Final = value

    @property
    def Fonts(self):
        return Fonts(self.com_object.Fonts)

    # Lower case aliases for Fonts
    @property
    def fonts(self):
        return self.Fonts

    @property
    def FullName(self):
        return self.com_object.FullName

    # Lower case aliases for FullName
    @property
    def fullname(self):
        return self.FullName

    @property
    def GridDistance(self):
        return self.com_object.GridDistance

    @GridDistance.setter
    def GridDistance(self, value):
        self.com_object.GridDistance = value

    # Lower case aliases for GridDistance
    @property
    def griddistance(self):
        return self.GridDistance

    @griddistance.setter
    def griddistance(self, value):
        self.GridDistance = value

    @property
    def HandoutMaster(self):
        return Master(self.com_object.HandoutMaster)

    # Lower case aliases for HandoutMaster
    @property
    def handoutmaster(self):
        return self.HandoutMaster

    @property
    def HasHandoutMaster(self):
        return self.com_object.HasHandoutMaster

    # Lower case aliases for HasHandoutMaster
    @property
    def hashandoutmaster(self):
        return self.HasHandoutMaster

    @property
    def HasNotesMaster(self):
        return self.com_object.HasNotesMaster

    # Lower case aliases for HasNotesMaster
    @property
    def hasnotesmaster(self):
        return self.HasNotesMaster

    @property
    def HasTitleMaster(self):
        return self.com_object.HasTitleMaster

    # Lower case aliases for HasTitleMaster
    @property
    def hastitlemaster(self):
        return self.HasTitleMaster

    @property
    def HasVBProject(self):
        return self.com_object.HasVBProject

    # Lower case aliases for HasVBProject
    @property
    def hasvbproject(self):
        return self.HasVBProject

    @property
    def InMergeMode(self):
        return self.com_object.InMergeMode

    # Lower case aliases for InMergeMode
    @property
    def inmergemode(self):
        return self.InMergeMode

    @property
    def IsFullyDownloaded(self):
        return self.com_object.IsFullyDownloaded

    # Lower case aliases for IsFullyDownloaded
    @property
    def isfullydownloaded(self):
        return self.IsFullyDownloaded

    @property
    def LayoutDirection(self):
        return self.com_object.LayoutDirection

    @LayoutDirection.setter
    def LayoutDirection(self, value):
        self.com_object.LayoutDirection = value

    # Lower case aliases for LayoutDirection
    @property
    def layoutdirection(self):
        return self.LayoutDirection

    @layoutdirection.setter
    def layoutdirection(self, value):
        self.LayoutDirection = value

    @property
    def Name(self):
        return self.com_object.Name

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def NoLineBreakAfter(self):
        return self.com_object.NoLineBreakAfter

    @NoLineBreakAfter.setter
    def NoLineBreakAfter(self, value):
        self.com_object.NoLineBreakAfter = value

    # Lower case aliases for NoLineBreakAfter
    @property
    def nolinebreakafter(self):
        return self.NoLineBreakAfter

    @nolinebreakafter.setter
    def nolinebreakafter(self, value):
        self.NoLineBreakAfter = value

    @property
    def NoLineBreakBefore(self):
        return self.com_object.NoLineBreakBefore

    @NoLineBreakBefore.setter
    def NoLineBreakBefore(self, value):
        self.com_object.NoLineBreakBefore = value

    # Lower case aliases for NoLineBreakBefore
    @property
    def nolinebreakbefore(self):
        return self.NoLineBreakBefore

    @nolinebreakbefore.setter
    def nolinebreakbefore(self, value):
        self.NoLineBreakBefore = value

    @property
    def NotesMaster(self):
        return Master(self.com_object.NotesMaster)

    # Lower case aliases for NotesMaster
    @property
    def notesmaster(self):
        return self.NotesMaster

    @property
    def PageSetup(self):
        return PageSetup(self.com_object.PageSetup)

    # Lower case aliases for PageSetup
    @property
    def pagesetup(self):
        return self.PageSetup

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Password(self):
        return self.com_object.Password

    @Password.setter
    def Password(self, value):
        self.com_object.Password = value

    # Lower case aliases for Password
    @property
    def password(self):
        return self.Password

    @password.setter
    def password(self, value):
        self.Password = value

    @property
    def PasswordEncryptionAlgorithm(self):
        return self.com_object.PasswordEncryptionAlgorithm

    # Lower case aliases for PasswordEncryptionAlgorithm
    @property
    def passwordencryptionalgorithm(self):
        return self.PasswordEncryptionAlgorithm

    @property
    def PasswordEncryptionFileProperties(self):
        return self.com_object.PasswordEncryptionFileProperties

    # Lower case aliases for PasswordEncryptionFileProperties
    @property
    def passwordencryptionfileproperties(self):
        return self.PasswordEncryptionFileProperties

    @property
    def PasswordEncryptionKeyLength(self):
        return self.com_object.PasswordEncryptionKeyLength

    # Lower case aliases for PasswordEncryptionKeyLength
    @property
    def passwordencryptionkeylength(self):
        return self.PasswordEncryptionKeyLength

    @property
    def PasswordEncryptionProvider(self):
        return self.com_object.PasswordEncryptionProvider

    # Lower case aliases for PasswordEncryptionProvider
    @property
    def passwordencryptionprovider(self):
        return self.PasswordEncryptionProvider

    @property
    def Path(self):
        return Presentation(self.com_object.Path)

    # Lower case aliases for Path
    @property
    def path(self):
        return self.Path

    @property
    def PrintOptions(self):
        return PrintOptions(self.com_object.PrintOptions)

    # Lower case aliases for PrintOptions
    @property
    def printoptions(self):
        return self.PrintOptions

    @property
    def ReadOnly(self):
        return self.com_object.ReadOnly

    # Lower case aliases for ReadOnly
    @property
    def readonly(self):
        return self.ReadOnly

    @property
    def ReadOnlyRecommended(self):
        return self.com_object.ReadOnlyRecommended

    # Lower case aliases for ReadOnlyRecommended
    @property
    def readonlyrecommended(self):
        return self.ReadOnlyRecommended

    @property
    def RemovePersonalInformation(self):
        return self.com_object.RemovePersonalInformation

    @RemovePersonalInformation.setter
    def RemovePersonalInformation(self, value):
        self.com_object.RemovePersonalInformation = value

    # Lower case aliases for RemovePersonalInformation
    @property
    def removepersonalinformation(self):
        return self.RemovePersonalInformation

    @removepersonalinformation.setter
    def removepersonalinformation(self, value):
        self.RemovePersonalInformation = value

    @property
    def Research(self):
        return Research(self.com_object.Research)

    # Lower case aliases for Research
    @property
    def research(self):
        return self.Research

    @property
    def Saved(self):
        return self.com_object.Saved

    @Saved.setter
    def Saved(self, value):
        self.com_object.Saved = value

    # Lower case aliases for Saved
    @property
    def saved(self):
        return self.Saved

    @saved.setter
    def saved(self, value):
        self.Saved = value

    @property
    def SectionProperties(self):
        return SectionProperties(self.com_object.SectionProperties)

    # Lower case aliases for SectionProperties
    @property
    def sectionproperties(self):
        return self.SectionProperties

    @property
    def SensitivityLabel(self):
        return self.com_object.SensitivityLabel

    # Lower case aliases for SensitivityLabel
    @property
    def sensitivitylabel(self):
        return self.SensitivityLabel

    @property
    def ServerPolicy(self):
        return self.com_object.ServerPolicy

    # Lower case aliases for ServerPolicy
    @property
    def serverpolicy(self):
        return self.ServerPolicy

    @property
    def SharedWorkspace(self):
        return self.com_object.SharedWorkspace

    # Lower case aliases for SharedWorkspace
    @property
    def sharedworkspace(self):
        return self.SharedWorkspace

    @property
    def Signatures(self):
        return self.com_object.Signatures

    # Lower case aliases for Signatures
    @property
    def signatures(self):
        return self.Signatures

    @property
    def SlideMaster(self):
        return Master(self.com_object.SlideMaster)

    # Lower case aliases for SlideMaster
    @property
    def slidemaster(self):
        return self.SlideMaster

    @property
    def Slides(self):
        return Slides(self.com_object.Slides)

    # Lower case aliases for Slides
    @property
    def slides(self):
        return self.Slides

    @property
    def SlideShowSettings(self):
        return SlideShowSettings(self.com_object.SlideShowSettings)

    # Lower case aliases for SlideShowSettings
    @property
    def slideshowsettings(self):
        return self.SlideShowSettings

    @property
    def SlideShowWindow(self):
        return SlideShowWindow(self.com_object.SlideShowWindow)

    # Lower case aliases for SlideShowWindow
    @property
    def slideshowwindow(self):
        return self.SlideShowWindow

    @property
    def SnapToGrid(self):
        return self.com_object.SnapToGrid

    @SnapToGrid.setter
    def SnapToGrid(self, value):
        self.com_object.SnapToGrid = value

    # Lower case aliases for SnapToGrid
    @property
    def snaptogrid(self):
        return self.SnapToGrid

    @snaptogrid.setter
    def snaptogrid(self, value):
        self.SnapToGrid = value

    @property
    def Sync(self):
        return self.com_object.Sync

    # Lower case aliases for Sync
    @property
    def sync(self):
        return self.Sync

    @property
    def Tags(self):
        return Tags(self.com_object.Tags)

    # Lower case aliases for Tags
    @property
    def tags(self):
        return self.Tags

    @property
    def TemplateName(self):
        return self.com_object.TemplateName

    # Lower case aliases for TemplateName
    @property
    def templatename(self):
        return self.TemplateName

    @property
    def TitleMaster(self):
        return Master(self.com_object.TitleMaster)

    # Lower case aliases for TitleMaster
    @property
    def titlemaster(self):
        return self.TitleMaster

    @property
    def VBASigned(self):
        return self.com_object.VBASigned

    # Lower case aliases for VBASigned
    @property
    def vbasigned(self):
        return self.VBASigned

    @property
    def VBProject(self):
        return self.com_object.VBProject

    # Lower case aliases for VBProject
    @property
    def vbproject(self):
        return self.VBProject

    @property
    def Windows(self):
        return DocumentWindows(self.com_object.Windows)

    # Lower case aliases for Windows
    @property
    def windows(self):
        return self.Windows

    @property
    def WritePassword(self):
        return self.com_object.WritePassword

    @WritePassword.setter
    def WritePassword(self, value):
        self.com_object.WritePassword = value

    # Lower case aliases for WritePassword
    @property
    def writepassword(self):
        return self.WritePassword

    @writepassword.setter
    def writepassword(self, value):
        self.WritePassword = value

    def AcceptAll(self):
        return self.com_object.AcceptAll()

    # Lower case alias for AcceptAll
    def acceptall(self):
        return self.AcceptAll()

    def AddTitleMaster(self):
        return Master(self.com_object.AddTitleMaster())

    # Lower case alias for AddTitleMaster
    def addtitlemaster(self):
        return self.AddTitleMaster()

    def AddToFavorites(self):
        self.com_object.AddToFavorites()

    # Lower case alias for AddToFavorites
    def addtofavorites(self):
        return self.AddToFavorites()

    def ApplyTemplate(self, FileName=None):
        arguments = com_arguments([unwrap(a) for a in [FileName]])
        self.com_object.ApplyTemplate(*arguments)

    # Lower case alias for ApplyTemplate
    def applytemplate(self, FileName=None):
        arguments = [FileName]
        return self.ApplyTemplate(*arguments)

    def ApplyTheme(self, themeName=None):
        arguments = com_arguments([unwrap(a) for a in [themeName]])
        self.com_object.ApplyTheme(*arguments)

    # Lower case alias for ApplyTheme
    def applytheme(self, themeName=None):
        arguments = [themeName]
        return self.ApplyTheme(*arguments)

    def CanCheckIn(self):
        return self.com_object.CanCheckIn()

    # Lower case alias for CanCheckIn
    def cancheckin(self):
        return self.CanCheckIn()

    def CheckIn(self, SaveChanges=None, Comments=None, MakePublic=None):
        arguments = com_arguments([unwrap(a) for a in [SaveChanges, Comments, MakePublic]])
        self.com_object.CheckIn(*arguments)

    # Lower case alias for CheckIn
    def checkin(self, SaveChanges=None, Comments=None, MakePublic=None):
        arguments = [SaveChanges, Comments, MakePublic]
        return self.CheckIn(*arguments)

    def CheckInWithVersion(self, SaveChanges=None, Comments=None, MakePublic=None, VersionType=None):
        arguments = com_arguments([unwrap(a) for a in [SaveChanges, Comments, MakePublic, VersionType]])
        self.com_object.CheckInWithVersion(*arguments)

    # Lower case alias for CheckInWithVersion
    def checkinwithversion(self, SaveChanges=None, Comments=None, MakePublic=None, VersionType=None):
        arguments = [SaveChanges, Comments, MakePublic, VersionType]
        return self.CheckInWithVersion(*arguments)

    def Close(self):
        self.com_object.Close()

    # Lower case alias for Close
    def close(self):
        return self.Close()

    def Convert2(self, FileName=None):
        arguments = com_arguments([unwrap(a) for a in [FileName]])
        self.com_object.Convert2(*arguments)

    # Lower case alias for Convert2
    def convert2(self, FileName=None):
        arguments = [FileName]
        return self.Convert2(*arguments)

    def CreateVideo(self, FileName=None, UseTimingsAndNarrations=None, DefaultSlideDuration=None, VertResolution=None, FramesPerSecond=None, Quality=None):
        arguments = com_arguments([unwrap(a) for a in [FileName, UseTimingsAndNarrations, DefaultSlideDuration, VertResolution, FramesPerSecond, Quality]])
        self.com_object.CreateVideo(*arguments)

    # Lower case alias for CreateVideo
    def createvideo(self, FileName=None, UseTimingsAndNarrations=None, DefaultSlideDuration=None, VertResolution=None, FramesPerSecond=None, Quality=None):
        arguments = [FileName, UseTimingsAndNarrations, DefaultSlideDuration, VertResolution, FramesPerSecond, Quality]
        return self.CreateVideo(*arguments)

    def EndReview(self):
        return self.com_object.EndReview()

    # Lower case alias for EndReview
    def endreview(self):
        return self.EndReview()

    def EnsureAllMediaUpgraded(self):
        self.com_object.EnsureAllMediaUpgraded()

    # Lower case alias for EnsureAllMediaUpgraded
    def ensureallmediaupgraded(self):
        return self.EnsureAllMediaUpgraded()

    def Export(self, Path=None, FilterName=None, ScaleWidth=None, ScaleHeight=None):
        arguments = com_arguments([unwrap(a) for a in [Path, FilterName, ScaleWidth, ScaleHeight]])
        self.com_object.Export(*arguments)

    # Lower case alias for Export
    def export(self, Path=None, FilterName=None, ScaleWidth=None, ScaleHeight=None):
        arguments = [Path, FilterName, ScaleWidth, ScaleHeight]
        return self.Export(*arguments)

    def ExportAsFixedFormat(self, Path=None, FixedFormatType=None, Intent=None, FrameSlides=None, HandoutOrder=None, OutputType=None, PrintHiddenSlides=None, PrintRange=None, RangeType=None, SlideShowName=None, IncludeDocProperties=None, KeepIRMSettings=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005_1=None, ExternalExporter=None):
        arguments = com_arguments([unwrap(a) for a in [Path, FixedFormatType, Intent, FrameSlides, HandoutOrder, OutputType, PrintHiddenSlides, PrintRange, RangeType, SlideShowName, IncludeDocProperties, KeepIRMSettings, DocStructureTags, BitmapMissingFonts, UseISO19005_1, ExternalExporter]])
        self.com_object.ExportAsFixedFormat(*arguments)

    # Lower case alias for ExportAsFixedFormat
    def exportasfixedformat(self, Path=None, FixedFormatType=None, Intent=None, FrameSlides=None, HandoutOrder=None, OutputType=None, PrintHiddenSlides=None, PrintRange=None, RangeType=None, SlideShowName=None, IncludeDocProperties=None, KeepIRMSettings=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005_1=None, ExternalExporter=None):
        arguments = [Path, FixedFormatType, Intent, FrameSlides, HandoutOrder, OutputType, PrintHiddenSlides, PrintRange, RangeType, SlideShowName, IncludeDocProperties, KeepIRMSettings, DocStructureTags, BitmapMissingFonts, UseISO19005_1, ExternalExporter]
        return self.ExportAsFixedFormat(*arguments)

    def FollowHyperlink(self, Address=None, SubAddress=None, NewWindow=None, AddHistory=None, ExtraInfo=None, Method=None, HeaderInfo=None):
        arguments = com_arguments([unwrap(a) for a in [Address, SubAddress, NewWindow, AddHistory, ExtraInfo, Method, HeaderInfo]])
        return self.com_object.FollowHyperlink(*arguments)

    # Lower case alias for FollowHyperlink
    def followhyperlink(self, Address=None, SubAddress=None, NewWindow=None, AddHistory=None, ExtraInfo=None, Method=None, HeaderInfo=None):
        arguments = [Address, SubAddress, NewWindow, AddHistory, ExtraInfo, Method, HeaderInfo]
        return self.FollowHyperlink(*arguments)

    def GetWorkflowTasks(self):
        return self.com_object.GetWorkflowTasks()

    # Lower case alias for GetWorkflowTasks
    def getworkflowtasks(self):
        return self.GetWorkflowTasks()

    def GetWorkflowTemplates(self):
        return self.com_object.GetWorkflowTemplates()

    # Lower case alias for GetWorkflowTemplates
    def getworkflowtemplates(self):
        return self.GetWorkflowTemplates()

    def LockServerFile(self):
        self.com_object.LockServerFile()

    # Lower case alias for LockServerFile
    def lockserverfile(self):
        return self.LockServerFile()

    def MergeWithBaseline(self, withPresentation=None, baselinePresentation=None):
        arguments = com_arguments([unwrap(a) for a in [withPresentation, baselinePresentation]])
        return self.com_object.MergeWithBaseline(*arguments)

    # Lower case alias for MergeWithBaseline
    def mergewithbaseline(self, withPresentation=None, baselinePresentation=None):
        arguments = [withPresentation, baselinePresentation]
        return self.MergeWithBaseline(*arguments)

    def NewWindow(self):
        return DocumentWindow(self.com_object.NewWindow())

    # Lower case alias for NewWindow
    def newwindow(self):
        return self.NewWindow()

    def PrintOut(self, From=None, To=None, PrintToFile=None, Copies=None, Collate=None):
        arguments = com_arguments([unwrap(a) for a in [From, To, PrintToFile, Copies, Collate]])
        self.com_object.PrintOut(*arguments)

    # Lower case alias for PrintOut
    def printout(self, From=None, To=None, PrintToFile=None, Copies=None, Collate=None):
        arguments = [From, To, PrintToFile, Copies, Collate]
        return self.PrintOut(*arguments)

    def PublishSlides(self, SlideLibraryUrl=None, Overwrite=None):
        arguments = com_arguments([unwrap(a) for a in [SlideLibraryUrl, Overwrite]])
        self.com_object.PublishSlides(*arguments)

    # Lower case alias for PublishSlides
    def publishslides(self, SlideLibraryUrl=None, Overwrite=None):
        arguments = [SlideLibraryUrl, Overwrite]
        return self.PublishSlides(*arguments)

    def RejectAll(self):
        return self.com_object.RejectAll()

    # Lower case alias for RejectAll
    def rejectall(self):
        return self.RejectAll()

    def RemoveDocumentInformation(self, Type=None):
        arguments = com_arguments([unwrap(a) for a in [Type]])
        self.com_object.RemoveDocumentInformation(*arguments)

    # Lower case alias for RemoveDocumentInformation
    def removedocumentinformation(self, Type=None):
        arguments = [Type]
        return self.RemoveDocumentInformation(*arguments)

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()

    def SaveAs(self, FileName=None, FileFormat=None, EmbedFonts=None):
        arguments = com_arguments([unwrap(a) for a in [FileName, FileFormat, EmbedFonts]])
        self.com_object.SaveAs(*arguments)

    # Lower case alias for SaveAs
    def saveas(self, FileName=None, FileFormat=None, EmbedFonts=None):
        arguments = [FileName, FileFormat, EmbedFonts]
        return self.SaveAs(*arguments)

    def SaveCopyAs(self, FileName=None, FileFormat=None, EmbedTrueTypeFonts=None):
        arguments = com_arguments([unwrap(a) for a in [FileName, FileFormat, EmbedTrueTypeFonts]])
        self.com_object.SaveCopyAs(*arguments)

    # Lower case alias for SaveCopyAs
    def savecopyas(self, FileName=None, FileFormat=None, EmbedTrueTypeFonts=None):
        arguments = [FileName, FileFormat, EmbedTrueTypeFonts]
        return self.SaveCopyAs(*arguments)

    def SaveCopyAs2(self, FileName=None, FileFormat=None, EmbedTrueTypeFonts=None, ReadOnlyRecommended=None):
        arguments = com_arguments([unwrap(a) for a in [FileName, FileFormat, EmbedTrueTypeFonts, ReadOnlyRecommended]])
        self.com_object.SaveCopyAs2(*arguments)

    # Lower case alias for SaveCopyAs2
    def savecopyas2(self, FileName=None, FileFormat=None, EmbedTrueTypeFonts=None, ReadOnlyRecommended=None):
        arguments = [FileName, FileFormat, EmbedTrueTypeFonts, ReadOnlyRecommended]
        return self.SaveCopyAs2(*arguments)

    def SendFaxOverInternet(self, Recipients=None, Subject=None, ShowMessage=None):
        arguments = com_arguments([unwrap(a) for a in [Recipients, Subject, ShowMessage]])
        self.com_object.SendFaxOverInternet(*arguments)

    # Lower case alias for SendFaxOverInternet
    def sendfaxoverinternet(self, Recipients=None, Subject=None, ShowMessage=None):
        arguments = [Recipients, Subject, ShowMessage]
        return self.SendFaxOverInternet(*arguments)

    def SetPasswordEncryptionOptions(self, PasswordEncryptionProvider=None, PasswordEncryptionAlgorithm=None, PasswordEncryptionKeyLength=None, PasswordEncryptionFileProperties=None):
        arguments = com_arguments([unwrap(a) for a in [PasswordEncryptionProvider, PasswordEncryptionAlgorithm, PasswordEncryptionKeyLength, PasswordEncryptionFileProperties]])
        self.com_object.SetPasswordEncryptionOptions(*arguments)

    # Lower case alias for SetPasswordEncryptionOptions
    def setpasswordencryptionoptions(self, PasswordEncryptionProvider=None, PasswordEncryptionAlgorithm=None, PasswordEncryptionKeyLength=None, PasswordEncryptionFileProperties=None):
        arguments = [PasswordEncryptionProvider, PasswordEncryptionAlgorithm, PasswordEncryptionKeyLength, PasswordEncryptionFileProperties]
        return self.SetPasswordEncryptionOptions(*arguments)

    def UpdateLinks(self):
        self.com_object.UpdateLinks()

    # Lower case alias for UpdateLinks
    def updatelinks(self):
        return self.UpdateLinks()


class Presentations:

    def __init__(self, presentations=None):
        self.com_object= presentations

    def __call__(self, item):
        return Presentation(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Add(self, WithWindow=None):
        arguments = com_arguments([unwrap(a) for a in [WithWindow]])
        return Presentation(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, WithWindow=None):
        arguments = [WithWindow]
        return self.Add(*arguments)

    def CanCheckOut(self, FileName=None):
        arguments = com_arguments([unwrap(a) for a in [FileName]])
        return self.com_object.CanCheckOut(*arguments)

    # Lower case alias for CanCheckOut
    def cancheckout(self, FileName=None):
        arguments = [FileName]
        return self.CanCheckOut(*arguments)

    def CheckOut(self, FileName=None):
        arguments = com_arguments([unwrap(a) for a in [FileName]])
        return self.com_object.CheckOut(*arguments)

    # Lower case alias for CheckOut
    def checkout(self, FileName=None):
        arguments = [FileName]
        return self.CheckOut(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Presentation(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Open(self, FileName=None, ReadOnly=None, Untitled=None, WithWindow=None):
        arguments = com_arguments([unwrap(a) for a in [FileName, ReadOnly, Untitled, WithWindow]])
        return Presentation(self.com_object.Open(*arguments))

    # Lower case alias for Open
    def open(self, FileName=None, ReadOnly=None, Untitled=None, WithWindow=None):
        arguments = [FileName, ReadOnly, Untitled, WithWindow]
        return self.Open(*arguments)

    def Open2007(self, FileName=None, ReadOnly=None, Untitled=None, WithWindow=None, OpenAndRepair=None):
        arguments = com_arguments([unwrap(a) for a in [FileName, ReadOnly, Untitled, WithWindow, OpenAndRepair]])
        return Presentation(self.com_object.Open2007(*arguments))

    # Lower case alias for Open2007
    def open2007(self, FileName=None, ReadOnly=None, Untitled=None, WithWindow=None, OpenAndRepair=None):
        arguments = [FileName, ReadOnly, Untitled, WithWindow, OpenAndRepair]
        return self.Open2007(*arguments)


class PrintOptions:

    def __init__(self, printoptions=None):
        self.com_object= printoptions

    @property
    def ActivePrinter(self):
        return self.com_object.ActivePrinter

    # Lower case aliases for ActivePrinter
    @property
    def activeprinter(self):
        return self.ActivePrinter

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Collate(self):
        return self.com_object.Collate

    @Collate.setter
    def Collate(self, value):
        self.com_object.Collate = value

    # Lower case aliases for Collate
    @property
    def collate(self):
        return self.Collate

    @collate.setter
    def collate(self, value):
        self.Collate = value

    @property
    def FitToPage(self):
        return self.com_object.FitToPage

    @FitToPage.setter
    def FitToPage(self, value):
        self.com_object.FitToPage = value

    # Lower case aliases for FitToPage
    @property
    def fittopage(self):
        return self.FitToPage

    @fittopage.setter
    def fittopage(self, value):
        self.FitToPage = value

    @property
    def FrameSlides(self):
        return self.com_object.FrameSlides

    @FrameSlides.setter
    def FrameSlides(self, value):
        self.com_object.FrameSlides = value

    # Lower case aliases for FrameSlides
    @property
    def frameslides(self):
        return self.FrameSlides

    @frameslides.setter
    def frameslides(self, value):
        self.FrameSlides = value

    @property
    def HandoutOrder(self):
        return self.com_object.HandoutOrder

    @HandoutOrder.setter
    def HandoutOrder(self, value):
        self.com_object.HandoutOrder = value

    # Lower case aliases for HandoutOrder
    @property
    def handoutorder(self):
        return self.HandoutOrder

    @handoutorder.setter
    def handoutorder(self, value):
        self.HandoutOrder = value

    @property
    def HighQuality(self):
        return self.com_object.HighQuality

    @HighQuality.setter
    def HighQuality(self, value):
        self.com_object.HighQuality = value

    # Lower case aliases for HighQuality
    @property
    def highquality(self):
        return self.HighQuality

    @highquality.setter
    def highquality(self, value):
        self.HighQuality = value

    @property
    def NumberOfCopies(self):
        return self.com_object.NumberOfCopies

    @NumberOfCopies.setter
    def NumberOfCopies(self, value):
        self.com_object.NumberOfCopies = value

    # Lower case aliases for NumberOfCopies
    @property
    def numberofcopies(self):
        return self.NumberOfCopies

    @numberofcopies.setter
    def numberofcopies(self, value):
        self.NumberOfCopies = value

    @property
    def OutputType(self):
        return self.com_object.OutputType

    @OutputType.setter
    def OutputType(self, value):
        self.com_object.OutputType = value

    # Lower case aliases for OutputType
    @property
    def outputtype(self):
        return self.OutputType

    @outputtype.setter
    def outputtype(self, value):
        self.OutputType = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PrintColorType(self):
        return self.com_object.PrintColorType

    @PrintColorType.setter
    def PrintColorType(self, value):
        self.com_object.PrintColorType = value

    # Lower case aliases for PrintColorType
    @property
    def printcolortype(self):
        return self.PrintColorType

    @printcolortype.setter
    def printcolortype(self, value):
        self.PrintColorType = value

    @property
    def PrintComments(self):
        return self.com_object.PrintComments

    @PrintComments.setter
    def PrintComments(self, value):
        self.com_object.PrintComments = value

    # Lower case aliases for PrintComments
    @property
    def printcomments(self):
        return self.PrintComments

    @printcomments.setter
    def printcomments(self, value):
        self.PrintComments = value

    @property
    def PrintFontsAsGraphics(self):
        return self.com_object.PrintFontsAsGraphics

    @PrintFontsAsGraphics.setter
    def PrintFontsAsGraphics(self, value):
        self.com_object.PrintFontsAsGraphics = value

    # Lower case aliases for PrintFontsAsGraphics
    @property
    def printfontsasgraphics(self):
        return self.PrintFontsAsGraphics

    @printfontsasgraphics.setter
    def printfontsasgraphics(self, value):
        self.PrintFontsAsGraphics = value

    @property
    def PrintHiddenSlides(self):
        return self.com_object.PrintHiddenSlides

    @PrintHiddenSlides.setter
    def PrintHiddenSlides(self, value):
        self.com_object.PrintHiddenSlides = value

    # Lower case aliases for PrintHiddenSlides
    @property
    def printhiddenslides(self):
        return self.PrintHiddenSlides

    @printhiddenslides.setter
    def printhiddenslides(self, value):
        self.PrintHiddenSlides = value

    @property
    def PrintInBackground(self):
        return self.com_object.PrintInBackground

    @PrintInBackground.setter
    def PrintInBackground(self, value):
        self.com_object.PrintInBackground = value

    # Lower case aliases for PrintInBackground
    @property
    def printinbackground(self):
        return self.PrintInBackground

    @printinbackground.setter
    def printinbackground(self, value):
        self.PrintInBackground = value

    @property
    def Ranges(self):
        return PrintRanges(self.com_object.Ranges)

    # Lower case aliases for Ranges
    @property
    def ranges(self):
        return self.Ranges

    @property
    def RangeType(self):
        return self.com_object.RangeType

    @RangeType.setter
    def RangeType(self, value):
        self.com_object.RangeType = value

    # Lower case aliases for RangeType
    @property
    def rangetype(self):
        return self.RangeType

    @rangetype.setter
    def rangetype(self, value):
        self.RangeType = value

    @property
    def sectionIndex(self):
        return PrintOptions(self.com_object.sectionIndex)

    @sectionIndex.setter
    def sectionIndex(self, value):
        self.com_object.sectionIndex = value

    # Lower case aliases for sectionIndex
    @property
    def sectionindex(self):
        return self.sectionIndex

    @sectionindex.setter
    def sectionindex(self, value):
        self.sectionIndex = value

    @property
    def SlideShowName(self):
        return self.com_object.SlideShowName

    @SlideShowName.setter
    def SlideShowName(self, value):
        self.com_object.SlideShowName = value

    # Lower case aliases for SlideShowName
    @property
    def slideshowname(self):
        return self.SlideShowName

    @slideshowname.setter
    def slideshowname(self, value):
        self.SlideShowName = value


class PrintRange:

    def __init__(self, printrange=None):
        self.com_object= printrange

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def End(self):
        return self.com_object.End

    # Lower case aliases for End
    @property
    def end(self):
        return self.End

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Start(self):
        return self.com_object.Start

    # Lower case aliases for Start
    @property
    def start(self):
        return self.Start

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()


class PrintRanges:

    def __init__(self, printranges=None):
        self.com_object= printranges

    def __call__(self, item):
        return PrintRange(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Add(self, Start=None, End=None):
        arguments = com_arguments([unwrap(a) for a in [Start, End]])
        return PrintRange(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Start=None, End=None):
        arguments = [Start, End]
        return self.Add(*arguments)

    def ClearAll(self):
        return self.com_object.ClearAll()

    # Lower case alias for ClearAll
    def clearall(self):
        return self.ClearAll()

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return PrintRange(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class PropertyEffect:

    def __init__(self, propertyeffect=None):
        self.com_object= propertyeffect

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def From(self):
        return self.com_object.From

    @From.setter
    def From(self, value):
        self.com_object.From = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Points(self):
        return AnimationPoints(self.com_object.Points)

    # Lower case aliases for Points
    @property
    def points(self):
        return self.Points

    @property
    def Property(self):
        return self.com_object.Property

    @Property.setter
    def Property(self, value):
        self.com_object.Property = value

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


class ProtectedViewWindow:

    def __init__(self, protectedviewwindow=None):
        self.com_object= protectedviewwindow

    @property
    def Active(self):
        return self.com_object.Active

    # Lower case aliases for Active
    @property
    def active(self):
        return self.Active

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
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Presentation(self):
        return Presentation(self.com_object.Presentation)

    # Lower case aliases for Presentation
    @property
    def presentation(self):
        return self.Presentation

    @property
    def SourceName(self):
        return ProtectedViewWindow(self.com_object.SourceName)

    # Lower case aliases for SourceName
    @property
    def sourcename(self):
        return self.SourceName

    @property
    def SourcePath(self):
        return ProtectedViewWindow(self.com_object.SourcePath)

    # Lower case aliases for SourcePath
    @property
    def sourcepath(self):
        return self.SourcePath

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
        return self.com_object.WindowState

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

    def Close(self):
        self.com_object.Close()

    # Lower case alias for Close
    def close(self):
        return self.Close()

    def Edit(self, ModifyPassword=None):
        arguments = com_arguments([unwrap(a) for a in [ModifyPassword]])
        return Presentation(self.com_object.Edit(*arguments))

    # Lower case alias for Edit
    def edit(self, ModifyPassword=None):
        arguments = [ModifyPassword]
        return self.Edit(*arguments)


class ProtectedViewWindows:

    def __init__(self, protectedviewwindows=None):
        self.com_object= protectedviewwindows

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return ProtectedViewWindow(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Open(self, FileName=None, ReadPassword=None, OpenAndRepair=None):
        arguments = com_arguments([unwrap(a) for a in [FileName, ReadPassword, OpenAndRepair]])
        return ProtectedViewWindow(self.com_object.Open(*arguments))

    # Lower case alias for Open
    def open(self, FileName=None, ReadPassword=None, OpenAndRepair=None):
        arguments = [FileName, ReadPassword, OpenAndRepair]
        return self.Open(*arguments)


class PublishObject:

    def __init__(self, publishobject=None):
        self.com_object= publishobject

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def FileName(self):
        return self.com_object.FileName

    @FileName.setter
    def FileName(self, value):
        self.com_object.FileName = value

    # Lower case aliases for FileName
    @property
    def filename(self):
        return self.FileName

    @filename.setter
    def filename(self, value):
        self.FileName = value

    @property
    def HTMLVersion(self):
        return self.com_object.HTMLVersion

    @HTMLVersion.setter
    def HTMLVersion(self, value):
        self.com_object.HTMLVersion = value

    # Lower case aliases for HTMLVersion
    @property
    def htmlversion(self):
        return self.HTMLVersion

    @htmlversion.setter
    def htmlversion(self, value):
        self.HTMLVersion = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RangeEnd(self):
        return self.com_object.RangeEnd

    @RangeEnd.setter
    def RangeEnd(self, value):
        self.com_object.RangeEnd = value

    # Lower case aliases for RangeEnd
    @property
    def rangeend(self):
        return self.RangeEnd

    @rangeend.setter
    def rangeend(self, value):
        self.RangeEnd = value

    @property
    def RangeStart(self):
        return self.com_object.RangeStart

    @RangeStart.setter
    def RangeStart(self, value):
        self.com_object.RangeStart = value

    # Lower case aliases for RangeStart
    @property
    def rangestart(self):
        return self.RangeStart

    @rangestart.setter
    def rangestart(self, value):
        self.RangeStart = value

    @property
    def SlideShowName(self):
        return self.com_object.SlideShowName

    @SlideShowName.setter
    def SlideShowName(self, value):
        self.com_object.SlideShowName = value

    # Lower case aliases for SlideShowName
    @property
    def slideshowname(self):
        return self.SlideShowName

    @slideshowname.setter
    def slideshowname(self, value):
        self.SlideShowName = value

    @property
    def SourceType(self):
        return self.com_object.SourceType

    @SourceType.setter
    def SourceType(self, value):
        self.com_object.SourceType = value

    # Lower case aliases for SourceType
    @property
    def sourcetype(self):
        return self.SourceType

    @sourcetype.setter
    def sourcetype(self, value):
        self.SourceType = value

    @property
    def SpeakerNotes(self):
        return self.com_object.SpeakerNotes

    @SpeakerNotes.setter
    def SpeakerNotes(self, value):
        self.com_object.SpeakerNotes = value

    # Lower case aliases for SpeakerNotes
    @property
    def speakernotes(self):
        return self.SpeakerNotes

    @speakernotes.setter
    def speakernotes(self, value):
        self.SpeakerNotes = value

    def Publish(self):
        self.com_object.Publish()

    # Lower case alias for Publish
    def publish(self):
        return self.Publish()


class PublishObjects:

    def __init__(self, publishobjects=None):
        self.com_object= publishobjects

    def __call__(self, item):
        return PublishObject(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return PublishObject(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class ResampleMediaTasks:

    def __init__(self, resamplemediatasks=None):
        self.com_object= resamplemediatasks

    def __call__(self, item):
        return ResampleMediaTask(self.com_object(item))

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def PercentComplete(self):
        return self.com_object.PercentComplete

    # Lower case aliases for PercentComplete
    @property
    def percentcomplete(self):
        return self.PercentComplete

    def Cancel(self):
        return self.com_object.Cancel()

    # Lower case alias for Cancel
    def cancel(self):
        return self.Cancel()

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return ResampleMediaTask(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Pause(self):
        self.com_object.Pause()

    # Lower case alias for Pause
    def pause(self):
        return self.Pause()

    def Resume(self):
        return self.com_object.Resume()

    # Lower case alias for Resume
    def resume(self):
        return self.Resume()


class Research:

    def __init__(self, research=None):
        self.com_object= research

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    def IsResearchService(self, ServiceID=None):
        arguments = com_arguments([unwrap(a) for a in [ServiceID]])
        return self.com_object.IsResearchService(*arguments)

    # Lower case alias for IsResearchService
    def isresearchservice(self, ServiceID=None):
        arguments = [ServiceID]
        return self.IsResearchService(*arguments)

    def Query(self, ServiceID=None, QueryString=None, QueryLanguage=None, UseSelection=None, RequeryContextXML=None, NewQueryContextXML=None, LaunchQuery=None):
        arguments = com_arguments([unwrap(a) for a in [ServiceID, QueryString, QueryLanguage, UseSelection, RequeryContextXML, NewQueryContextXML, LaunchQuery]])
        self.com_object.Query(*arguments)

    # Lower case alias for Query
    def query(self, ServiceID=None, QueryString=None, QueryLanguage=None, UseSelection=None, RequeryContextXML=None, NewQueryContextXML=None, LaunchQuery=None):
        arguments = [ServiceID, QueryString, QueryLanguage, UseSelection, RequeryContextXML, NewQueryContextXML, LaunchQuery]
        return self.Query(*arguments)

    def SetLanguagePair(self, Language1=None, Language2=None):
        arguments = com_arguments([unwrap(a) for a in [Language1, Language2]])
        self.com_object.SetLanguagePair(*arguments)

    # Lower case alias for SetLanguagePair
    def setlanguagepair(self, Language1=None, Language2=None):
        arguments = [Language1, Language2]
        return self.SetLanguagePair(*arguments)


class RGBColor:

    def __init__(self, rgbcolor=None):
        self.com_object= rgbcolor

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RGB(self):
        return PpColorSchemeIndex(self.com_object.RGB)

    @RGB.setter
    def RGB(self, value):
        self.com_object.RGB = value

    # Lower case aliases for RGB
    @property
    def rgb(self):
        return self.RGB

    @rgb.setter
    def rgb(self, value):
        self.RGB = value


class RotationEffect:

    def __init__(self, rotationeffect=None):
        self.com_object= rotationeffect

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def By(self):
        return self.com_object.By

    @By.setter
    def By(self, value):
        self.com_object.By = value

    # Lower case aliases for By
    @property
    def by(self):
        return self.By

    @by.setter
    def by(self, value):
        self.By = value

    @property
    def From(self):
        return self.com_object.From

    @From.setter
    def From(self, value):
        self.com_object.From = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

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


class Row:

    def __init__(self, row=None):
        self.com_object= row

    @property
    def Application(self):
        return Application(self.com_object.Application)

    def Cells(self, RowIndex=None, ColumnIndex=None):
        arguments = com_arguments([unwrap(a) for a in [RowIndex, ColumnIndex]])
        if hasattr(self.com_object, "GetCells"):
            return CellRange(self.com_object.GetCells(*arguments))
        else:
            return CellRange(self.com_object.Cells(*arguments))

    # Lower case aliases for Cells
    def cells(self, RowIndex=None, ColumnIndex=None):
        arguments = [RowIndex, ColumnIndex]
        return self.Cells(*arguments)

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
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class Rows:

    def __init__(self, rows=None):
        self.com_object= rows

    def __call__(self, item):
        return Row(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Add(self, BeforeRow=None):
        arguments = com_arguments([unwrap(a) for a in [BeforeRow]])
        return Row(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, BeforeRow=None):
        arguments = [BeforeRow]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Row(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class Ruler:

    def __init__(self, ruler=None):
        self.com_object= ruler

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Levels(self):
        return RulerLevels(self.com_object.Levels)

    # Lower case aliases for Levels
    @property
    def levels(self):
        return self.Levels

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def TabStops(self):
        return TabStops(self.com_object.TabStops)

    # Lower case aliases for TabStops
    @property
    def tabstops(self):
        return self.TabStops


class RulerLevel:

    def __init__(self, rulerlevel=None):
        self.com_object= rulerlevel

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def FirstMargin(self):
        return self.com_object.FirstMargin

    @FirstMargin.setter
    def FirstMargin(self, value):
        self.com_object.FirstMargin = value

    # Lower case aliases for FirstMargin
    @property
    def firstmargin(self):
        return self.FirstMargin

    @firstmargin.setter
    def firstmargin(self, value):
        self.FirstMargin = value

    @property
    def LeftMargin(self):
        return self.com_object.LeftMargin

    @LeftMargin.setter
    def LeftMargin(self, value):
        self.com_object.LeftMargin = value

    # Lower case aliases for LeftMargin
    @property
    def leftmargin(self):
        return self.LeftMargin

    @leftmargin.setter
    def leftmargin(self, value):
        self.LeftMargin = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent


class RulerLevels:

    def __init__(self, rulerlevels=None):
        self.com_object= rulerlevels

    def __call__(self, item):
        return RulerLevel(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return RulerLevel(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class ScaleEffect:

    def __init__(self, scaleeffect=None):
        self.com_object= scaleeffect

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def ByX(self):
        return self.com_object.ByX

    @ByX.setter
    def ByX(self, value):
        self.com_object.ByX = value

    # Lower case aliases for ByX
    @property
    def byx(self):
        return self.ByX

    @byx.setter
    def byx(self, value):
        self.ByX = value

    @property
    def ByY(self):
        return self.com_object.ByY

    @ByY.setter
    def ByY(self, value):
        self.com_object.ByY = value

    # Lower case aliases for ByY
    @property
    def byy(self):
        return self.ByY

    @byy.setter
    def byy(self, value):
        self.ByY = value

    @property
    def FromX(self):
        return self.com_object.FromX

    @FromX.setter
    def FromX(self, value):
        self.com_object.FromX = value

    # Lower case aliases for FromX
    @property
    def fromx(self):
        return self.FromX

    @fromx.setter
    def fromx(self, value):
        self.FromX = value

    @property
    def FromY(self):
        return ScaleEffect(self.com_object.FromY)

    @FromY.setter
    def FromY(self, value):
        self.com_object.FromY = value

    # Lower case aliases for FromY
    @property
    def fromy(self):
        return self.FromY

    @fromy.setter
    def fromy(self, value):
        self.FromY = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ToX(self):
        return self.com_object.ToX

    @ToX.setter
    def ToX(self, value):
        self.com_object.ToX = value

    # Lower case aliases for ToX
    @property
    def tox(self):
        return self.ToX

    @tox.setter
    def tox(self, value):
        self.ToX = value

    @property
    def ToY(self):
        return ScaleEffect(self.com_object.ToY)

    @ToY.setter
    def ToY(self, value):
        self.com_object.ToY = value

    # Lower case aliases for ToY
    @property
    def toy(self):
        return self.ToY

    @toy.setter
    def toy(self, value):
        self.ToY = value


class SectionProperties:

    def __init__(self, sectionproperties=None):
        self.com_object= sectionproperties

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def AddBeforeSlide(self, SlideIndex=None, sectionName=None):
        arguments = com_arguments([unwrap(a) for a in [SlideIndex, sectionName]])
        return self.com_object.AddBeforeSlide(*arguments)

    # Lower case alias for AddBeforeSlide
    def addbeforeslide(self, SlideIndex=None, sectionName=None):
        arguments = [SlideIndex, sectionName]
        return self.AddBeforeSlide(*arguments)

    def AddSection(self, sectionIndex=None, sectionName=None):
        arguments = com_arguments([unwrap(a) for a in [sectionIndex, sectionName]])
        return self.com_object.AddSection(*arguments)

    # Lower case alias for AddSection
    def addsection(self, sectionIndex=None, sectionName=None):
        arguments = [sectionIndex, sectionName]
        return self.AddSection(*arguments)

    def Delete(self, sectionIndex=None, deleteSlides=None):
        arguments = com_arguments([unwrap(a) for a in [sectionIndex, deleteSlides]])
        self.com_object.Delete(*arguments)

    # Lower case alias for Delete
    def delete(self, sectionIndex=None, deleteSlides=None):
        arguments = [sectionIndex, deleteSlides]
        return self.Delete(*arguments)

    def FirstSlide(self, sectionIndex=None):
        arguments = com_arguments([unwrap(a) for a in [sectionIndex]])
        return self.com_object.FirstSlide(*arguments)

    # Lower case alias for FirstSlide
    def firstslide(self, sectionIndex=None):
        arguments = [sectionIndex]
        return self.FirstSlide(*arguments)

    def Move(self, sectionIndex=None, toPos=None):
        arguments = com_arguments([unwrap(a) for a in [sectionIndex, toPos]])
        self.com_object.Move(*arguments)

    # Lower case alias for Move
    def move(self, sectionIndex=None, toPos=None):
        arguments = [sectionIndex, toPos]
        return self.Move(*arguments)

    def Name(self, sectionIndex=None):
        arguments = com_arguments([unwrap(a) for a in [sectionIndex]])
        return self.com_object.Name(*arguments)

    # Lower case alias for Name
    def name(self, sectionIndex=None):
        arguments = [sectionIndex]
        return self.Name(*arguments)

    def Rename(self, sectionIndex=None, sectionName=None):
        arguments = com_arguments([unwrap(a) for a in [sectionIndex, sectionName]])
        self.com_object.Rename(*arguments)

    # Lower case alias for Rename
    def rename(self, sectionIndex=None, sectionName=None):
        arguments = [sectionIndex, sectionName]
        return self.Rename(*arguments)

    def SectionID(self, sectionIndex=None):
        arguments = com_arguments([unwrap(a) for a in [sectionIndex]])
        return self.com_object.SectionID(*arguments)

    # Lower case alias for SectionID
    def sectionid(self, sectionIndex=None):
        arguments = [sectionIndex]
        return self.SectionID(*arguments)

    def SlidesCount(self, sectionIndex=None):
        arguments = com_arguments([unwrap(a) for a in [sectionIndex]])
        return self.com_object.SlidesCount(*arguments)

    # Lower case alias for SlidesCount
    def slidescount(self, sectionIndex=None):
        arguments = [sectionIndex]
        return self.SlidesCount(*arguments)


class Selection:

    def __init__(self, selection=None):
        self.com_object= selection

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def ChildShapeRange(self):
        return ShapeRange(self.com_object.ChildShapeRange)

    # Lower case aliases for ChildShapeRange
    @property
    def childshaperange(self):
        return self.ChildShapeRange

    @property
    def HasChildShapeRange(self):
        return self.com_object.HasChildShapeRange

    # Lower case aliases for HasChildShapeRange
    @property
    def haschildshaperange(self):
        return self.HasChildShapeRange

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ShapeRange(self):
        return ShapeRange(self.com_object.ShapeRange)

    # Lower case aliases for ShapeRange
    @property
    def shaperange(self):
        return self.ShapeRange

    @property
    def SlideRange(self):
        return SlideRange(self.com_object.SlideRange)

    # Lower case aliases for SlideRange
    @property
    def sliderange(self):
        return self.SlideRange

    @property
    def TextRange(self):
        return TextRange(self.com_object.TextRange)

    # Lower case aliases for TextRange
    @property
    def textrange(self):
        return self.TextRange

    @property
    def TextRange2(self):
        return self.com_object.TextRange2

    # Lower case aliases for TextRange2
    @property
    def textrange2(self):
        return self.TextRange2

    @property
    def Type(self):
        return self.com_object.Type

    # Lower case aliases for Type
    @property
    def type(self):
        return self.Type

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

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Unselect(self):
        self.com_object.Unselect()

    # Lower case alias for Unselect
    def unselect(self):
        return self.Unselect()


class Sequence:

    def __init__(self, sequence=None):
        self.com_object= sequence

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def AddEffect(self, Shape=None, effectId=None, Level=None, trigger=None, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Shape, effectId, Level, trigger, Index]])
        return Effect(self.com_object.AddEffect(*arguments))

    # Lower case alias for AddEffect
    def addeffect(self, Shape=None, effectId=None, Level=None, trigger=None, Index=None):
        arguments = [Shape, effectId, Level, trigger, Index]
        return self.AddEffect(*arguments)

    def AddTriggerEffect(self, pShape=None, effectId=None, trigger=None, pTriggerShape=None, bookmark=None, Level=None):
        arguments = com_arguments([unwrap(a) for a in [pShape, effectId, trigger, pTriggerShape, bookmark, Level]])
        return Effect(self.com_object.AddTriggerEffect(*arguments))

    # Lower case alias for AddTriggerEffect
    def addtriggereffect(self, pShape=None, effectId=None, trigger=None, pTriggerShape=None, bookmark=None, Level=None):
        arguments = [pShape, effectId, trigger, pTriggerShape, bookmark, Level]
        return self.AddTriggerEffect(*arguments)

    def Clone(self, Effect=None, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Effect, Index]])
        return Effect(self.com_object.Clone(*arguments))

    # Lower case alias for Clone
    def clone(self, Effect=None, Index=None):
        arguments = [Effect, Index]
        return self.Clone(*arguments)

    def ConvertToAfterEffect(self, Effect=None, After=None, DimColor=None, DimSchemeColor=None):
        arguments = com_arguments([unwrap(a) for a in [Effect, After, DimColor, DimSchemeColor]])
        return Effect(self.com_object.ConvertToAfterEffect(*arguments))

    # Lower case alias for ConvertToAfterEffect
    def converttoaftereffect(self, Effect=None, After=None, DimColor=None, DimSchemeColor=None):
        arguments = [Effect, After, DimColor, DimSchemeColor]
        return self.ConvertToAfterEffect(*arguments)

    def ConvertToAnimateBackground(self, Effect=None, AnimateBackground=None):
        arguments = com_arguments([unwrap(a) for a in [Effect, AnimateBackground]])
        return Effect(self.com_object.ConvertToAnimateBackground(*arguments))

    # Lower case alias for ConvertToAnimateBackground
    def converttoanimatebackground(self, Effect=None, AnimateBackground=None):
        arguments = [Effect, AnimateBackground]
        return self.ConvertToAnimateBackground(*arguments)

    def ConvertToAnimateInReverse(self, Effect=None, animateInReverse=None):
        arguments = com_arguments([unwrap(a) for a in [Effect, animateInReverse]])
        return Effect(self.com_object.ConvertToAnimateInReverse(*arguments))

    # Lower case alias for ConvertToAnimateInReverse
    def converttoanimateinreverse(self, Effect=None, animateInReverse=None):
        arguments = [Effect, animateInReverse]
        return self.ConvertToAnimateInReverse(*arguments)

    def ConvertToBuildLevel(self, Effect=None, Level=None):
        arguments = com_arguments([unwrap(a) for a in [Effect, Level]])
        return Effect(self.com_object.ConvertToBuildLevel(*arguments))

    # Lower case alias for ConvertToBuildLevel
    def converttobuildlevel(self, Effect=None, Level=None):
        arguments = [Effect, Level]
        return self.ConvertToBuildLevel(*arguments)

    def ConvertToTextUnitEffect(self, Effect=None, unitEffect=None):
        arguments = com_arguments([unwrap(a) for a in [Effect, unitEffect]])
        return Effect(self.com_object.ConvertToTextUnitEffect(*arguments))

    # Lower case alias for ConvertToTextUnitEffect
    def converttotextuniteffect(self, Effect=None, unitEffect=None):
        arguments = [Effect, unitEffect]
        return self.ConvertToTextUnitEffect(*arguments)

    def FindFirstAnimationFor(self, Shape=None):
        arguments = com_arguments([unwrap(a) for a in [Shape]])
        return Effect(self.com_object.FindFirstAnimationFor(*arguments))

    # Lower case alias for FindFirstAnimationFor
    def findfirstanimationfor(self, Shape=None):
        arguments = [Shape]
        return self.FindFirstAnimationFor(*arguments)

    def FindFirstAnimationForClick(self, click=None):
        arguments = com_arguments([unwrap(a) for a in [click]])
        return Effect(self.com_object.FindFirstAnimationForClick(*arguments))

    # Lower case alias for FindFirstAnimationForClick
    def findfirstanimationforclick(self, click=None):
        arguments = [click]
        return self.FindFirstAnimationForClick(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Effect(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class Sequences:

    def __init__(self, sequences=None):
        self.com_object= sequences

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Add(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Sequence(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Index=None):
        arguments = [Index]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Sequence(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class Series:

    def __init__(self, series=None):
        self.com_object= series

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def ApplyPictToEnd(self):
        return self.com_object.ApplyPictToEnd

    @ApplyPictToEnd.setter
    def ApplyPictToEnd(self, value):
        self.com_object.ApplyPictToEnd = value

    # Lower case aliases for ApplyPictToEnd
    @property
    def applypicttoend(self):
        return self.ApplyPictToEnd

    @applypicttoend.setter
    def applypicttoend(self, value):
        self.ApplyPictToEnd = value

    @property
    def ApplyPictToFront(self):
        return self.com_object.ApplyPictToFront

    @ApplyPictToFront.setter
    def ApplyPictToFront(self, value):
        self.com_object.ApplyPictToFront = value

    # Lower case aliases for ApplyPictToFront
    @property
    def applypicttofront(self):
        return self.ApplyPictToFront

    @applypicttofront.setter
    def applypicttofront(self, value):
        self.ApplyPictToFront = value

    @property
    def ApplyPictToSides(self):
        return self.com_object.ApplyPictToSides

    @ApplyPictToSides.setter
    def ApplyPictToSides(self, value):
        self.com_object.ApplyPictToSides = value

    # Lower case aliases for ApplyPictToSides
    @property
    def applypicttosides(self):
        return self.ApplyPictToSides

    @applypicttosides.setter
    def applypicttosides(self, value):
        self.ApplyPictToSides = value

    @property
    def AxisGroup(self):
        return XlAxisGroup(self.com_object.AxisGroup)

    @AxisGroup.setter
    def AxisGroup(self, value):
        self.com_object.AxisGroup = value

    # Lower case aliases for AxisGroup
    @property
    def axisgroup(self):
        return self.AxisGroup

    @axisgroup.setter
    def axisgroup(self, value):
        self.AxisGroup = value

    @property
    def BarShape(self):
        return XlBarShape(self.com_object.BarShape)

    @BarShape.setter
    def BarShape(self, value):
        self.com_object.BarShape = value

    # Lower case aliases for BarShape
    @property
    def barshape(self):
        return self.BarShape

    @barshape.setter
    def barshape(self, value):
        self.BarShape = value

    @property
    def BubbleSizes(self):
        return self.com_object.BubbleSizes

    @BubbleSizes.setter
    def BubbleSizes(self, value):
        self.com_object.BubbleSizes = value

    # Lower case aliases for BubbleSizes
    @property
    def bubblesizes(self):
        return self.BubbleSizes

    @bubblesizes.setter
    def bubblesizes(self, value):
        self.BubbleSizes = value

    @property
    def ChartType(self):
        return self.com_object.ChartType

    @ChartType.setter
    def ChartType(self, value):
        self.com_object.ChartType = value

    # Lower case aliases for ChartType
    @property
    def charttype(self):
        return self.ChartType

    @charttype.setter
    def charttype(self, value):
        self.ChartType = value

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def ErrorBars(self):
        return ErrorBars(self.com_object.ErrorBars)

    # Lower case aliases for ErrorBars
    @property
    def errorbars(self):
        return self.ErrorBars

    @property
    def Explosion(self):
        return self.com_object.Explosion

    @Explosion.setter
    def Explosion(self, value):
        self.com_object.Explosion = value

    # Lower case aliases for Explosion
    @property
    def explosion(self):
        return self.Explosion

    @explosion.setter
    def explosion(self, value):
        self.Explosion = value

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

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
    def FormulaLocal(self):
        return self.com_object.FormulaLocal

    @FormulaLocal.setter
    def FormulaLocal(self, value):
        self.com_object.FormulaLocal = value

    # Lower case aliases for FormulaLocal
    @property
    def formulalocal(self):
        return self.FormulaLocal

    @formulalocal.setter
    def formulalocal(self, value):
        self.FormulaLocal = value

    @property
    def FormulaR1C1(self):
        return self.com_object.FormulaR1C1

    @FormulaR1C1.setter
    def FormulaR1C1(self, value):
        self.com_object.FormulaR1C1 = value

    # Lower case aliases for FormulaR1C1
    @property
    def formular1c1(self):
        return self.FormulaR1C1

    @formular1c1.setter
    def formular1c1(self, value):
        self.FormulaR1C1 = value

    @property
    def FormulaR1C1Local(self):
        return self.com_object.FormulaR1C1Local

    @FormulaR1C1Local.setter
    def FormulaR1C1Local(self, value):
        self.com_object.FormulaR1C1Local = value

    # Lower case aliases for FormulaR1C1Local
    @property
    def formular1c1local(self):
        return self.FormulaR1C1Local

    @formular1c1local.setter
    def formular1c1local(self, value):
        self.FormulaR1C1Local = value

    @property
    def Has3DEffect(self):
        return self.com_object.Has3DEffect

    @Has3DEffect.setter
    def Has3DEffect(self, value):
        self.com_object.Has3DEffect = value

    # Lower case aliases for Has3DEffect
    @property
    def has3deffect(self):
        return self.Has3DEffect

    @has3deffect.setter
    def has3deffect(self, value):
        self.Has3DEffect = value

    @property
    def HasDataLabels(self):
        return self.com_object.HasDataLabels

    @HasDataLabels.setter
    def HasDataLabels(self, value):
        self.com_object.HasDataLabels = value

    # Lower case aliases for HasDataLabels
    @property
    def hasdatalabels(self):
        return self.HasDataLabels

    @hasdatalabels.setter
    def hasdatalabels(self, value):
        self.HasDataLabels = value

    @property
    def HasErrorBars(self):
        return self.com_object.HasErrorBars

    @HasErrorBars.setter
    def HasErrorBars(self, value):
        self.com_object.HasErrorBars = value

    # Lower case aliases for HasErrorBars
    @property
    def haserrorbars(self):
        return self.HasErrorBars

    @haserrorbars.setter
    def haserrorbars(self, value):
        self.HasErrorBars = value

    @property
    def HasLeaderLines(self):
        return self.com_object.HasLeaderLines

    @HasLeaderLines.setter
    def HasLeaderLines(self, value):
        self.com_object.HasLeaderLines = value

    # Lower case aliases for HasLeaderLines
    @property
    def hasleaderlines(self):
        return self.HasLeaderLines

    @hasleaderlines.setter
    def hasleaderlines(self, value):
        self.HasLeaderLines = value

    @property
    def InvertColor(self):
        return self.com_object.InvertColor

    @InvertColor.setter
    def InvertColor(self, value):
        self.com_object.InvertColor = value

    # Lower case aliases for InvertColor
    @property
    def invertcolor(self):
        return self.InvertColor

    @invertcolor.setter
    def invertcolor(self, value):
        self.InvertColor = value

    @property
    def InvertColorIndex(self):
        return self.com_object.InvertColorIndex

    @InvertColorIndex.setter
    def InvertColorIndex(self, value):
        self.com_object.InvertColorIndex = value

    # Lower case aliases for InvertColorIndex
    @property
    def invertcolorindex(self):
        return self.InvertColorIndex

    @invertcolorindex.setter
    def invertcolorindex(self, value):
        self.InvertColorIndex = value

    @property
    def InvertIfNegative(self):
        return self.com_object.InvertIfNegative

    @InvertIfNegative.setter
    def InvertIfNegative(self, value):
        self.com_object.InvertIfNegative = value

    # Lower case aliases for InvertIfNegative
    @property
    def invertifnegative(self):
        return self.InvertIfNegative

    @invertifnegative.setter
    def invertifnegative(self, value):
        self.InvertIfNegative = value

    @property
    def LeaderLines(self):
        return LeaderLines(self.com_object.LeaderLines)

    # Lower case aliases for LeaderLines
    @property
    def leaderlines(self):
        return self.LeaderLines

    @property
    def MarkerBackgroundColor(self):
        return self.com_object.MarkerBackgroundColor

    @MarkerBackgroundColor.setter
    def MarkerBackgroundColor(self, value):
        self.com_object.MarkerBackgroundColor = value

    # Lower case aliases for MarkerBackgroundColor
    @property
    def markerbackgroundcolor(self):
        return self.MarkerBackgroundColor

    @markerbackgroundcolor.setter
    def markerbackgroundcolor(self, value):
        self.MarkerBackgroundColor = value

    @property
    def MarkerBackgroundColorIndex(self):
        return XlColorIndex(self.com_object.MarkerBackgroundColorIndex)

    @MarkerBackgroundColorIndex.setter
    def MarkerBackgroundColorIndex(self, value):
        self.com_object.MarkerBackgroundColorIndex = value

    # Lower case aliases for MarkerBackgroundColorIndex
    @property
    def markerbackgroundcolorindex(self):
        return self.MarkerBackgroundColorIndex

    @markerbackgroundcolorindex.setter
    def markerbackgroundcolorindex(self, value):
        self.MarkerBackgroundColorIndex = value

    @property
    def MarkerForegroundColor(self):
        return self.com_object.MarkerForegroundColor

    @MarkerForegroundColor.setter
    def MarkerForegroundColor(self, value):
        self.com_object.MarkerForegroundColor = value

    # Lower case aliases for MarkerForegroundColor
    @property
    def markerforegroundcolor(self):
        return self.MarkerForegroundColor

    @markerforegroundcolor.setter
    def markerforegroundcolor(self, value):
        self.MarkerForegroundColor = value

    @property
    def MarkerForegroundColorIndex(self):
        return XlColorIndex(self.com_object.MarkerForegroundColorIndex)

    @MarkerForegroundColorIndex.setter
    def MarkerForegroundColorIndex(self, value):
        self.com_object.MarkerForegroundColorIndex = value

    # Lower case aliases for MarkerForegroundColorIndex
    @property
    def markerforegroundcolorindex(self):
        return self.MarkerForegroundColorIndex

    @markerforegroundcolorindex.setter
    def markerforegroundcolorindex(self, value):
        self.MarkerForegroundColorIndex = value

    @property
    def MarkerSize(self):
        return self.com_object.MarkerSize

    @MarkerSize.setter
    def MarkerSize(self, value):
        self.com_object.MarkerSize = value

    # Lower case aliases for MarkerSize
    @property
    def markersize(self):
        return self.MarkerSize

    @markersize.setter
    def markersize(self, value):
        self.MarkerSize = value

    @property
    def MarkerStyle(self):
        return XlMarkerStyle(self.com_object.MarkerStyle)

    @MarkerStyle.setter
    def MarkerStyle(self, value):
        self.com_object.MarkerStyle = value

    # Lower case aliases for MarkerStyle
    @property
    def markerstyle(self):
        return self.MarkerStyle

    @markerstyle.setter
    def markerstyle(self, value):
        self.MarkerStyle = value

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
    def PictureType(self):
        return XlChartPictureType(self.com_object.PictureType)

    @PictureType.setter
    def PictureType(self, value):
        self.com_object.PictureType = value

    # Lower case aliases for PictureType
    @property
    def picturetype(self):
        return self.PictureType

    @picturetype.setter
    def picturetype(self, value):
        self.PictureType = value

    @property
    def PictureUnit2(self):
        return self.com_object.PictureUnit2

    @PictureUnit2.setter
    def PictureUnit2(self, value):
        self.com_object.PictureUnit2 = value

    # Lower case aliases for PictureUnit2
    @property
    def pictureunit2(self):
        return self.PictureUnit2

    @pictureunit2.setter
    def pictureunit2(self, value):
        self.PictureUnit2 = value

    @property
    def PlotColorIndex(self):
        return self.com_object.PlotColorIndex

    # Lower case aliases for PlotColorIndex
    @property
    def plotcolorindex(self):
        return self.PlotColorIndex

    @property
    def PlotOrder(self):
        return self.com_object.PlotOrder

    @PlotOrder.setter
    def PlotOrder(self, value):
        self.com_object.PlotOrder = value

    # Lower case aliases for PlotOrder
    @property
    def plotorder(self):
        return self.PlotOrder

    @plotorder.setter
    def plotorder(self, value):
        self.PlotOrder = value

    @property
    def Shadow(self):
        return self.com_object.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.com_object.Shadow = value

    # Lower case aliases for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @shadow.setter
    def shadow(self, value):
        self.Shadow = value

    @property
    def Smooth(self):
        return self.com_object.Smooth

    @Smooth.setter
    def Smooth(self, value):
        self.com_object.Smooth = value

    # Lower case aliases for Smooth
    @property
    def smooth(self):
        return self.Smooth

    @smooth.setter
    def smooth(self, value):
        self.Smooth = value

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
    def Values(self):
        return self.com_object.Values

    @Values.setter
    def Values(self, value):
        self.com_object.Values = value

    # Lower case aliases for Values
    @property
    def values(self):
        return self.Values

    @values.setter
    def values(self, value):
        self.Values = value

    @property
    def XValues(self):
        return self.com_object.XValues

    @XValues.setter
    def XValues(self, value):
        self.com_object.XValues = value

    # Lower case aliases for XValues
    @property
    def xvalues(self):
        return self.XValues

    @xvalues.setter
    def xvalues(self, value):
        self.XValues = value

    def ApplyDataLabels(self, Type=None, LegendKey=None, AutoText=None, HasLeaderLines=None, ShowSeriesName=None, ShowCategoryName=None, ShowValue=None, ShowPercentage=None, ShowBubbleSize=None, Separator=None):
        arguments = com_arguments([unwrap(a) for a in [Type, LegendKey, AutoText, HasLeaderLines, ShowSeriesName, ShowCategoryName, ShowValue, ShowPercentage, ShowBubbleSize, Separator]])
        self.com_object.ApplyDataLabels(*arguments)

    # Lower case alias for ApplyDataLabels
    def applydatalabels(self, Type=None, LegendKey=None, AutoText=None, HasLeaderLines=None, ShowSeriesName=None, ShowCategoryName=None, ShowValue=None, ShowPercentage=None, ShowBubbleSize=None, Separator=None):
        arguments = [Type, LegendKey, AutoText, HasLeaderLines, ShowSeriesName, ShowCategoryName, ShowValue, ShowPercentage, ShowBubbleSize, Separator]
        return self.ApplyDataLabels(*arguments)

    def ClearFormats(self):
        self.com_object.ClearFormats()

    # Lower case alias for ClearFormats
    def clearformats(self):
        return self.ClearFormats()

    def Copy(self):
        self.com_object.Copy()

    # Lower case alias for Copy
    def copy(self):
        return self.Copy()

    def DataLabels(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return DataLabel(self.com_object.DataLabels(*arguments))

    # Lower case alias for DataLabels
    def datalabels(self, Index=None):
        arguments = [Index]
        return self.DataLabels(*arguments)

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def ErrorBar(self, Direction=None, Include=None, Type=None, Amount=None, MinusValues=None):
        arguments = com_arguments([unwrap(a) for a in [Direction, Include, Type, Amount, MinusValues]])
        self.com_object.ErrorBar(*arguments)

    # Lower case alias for ErrorBar
    def errorbar(self, Direction=None, Include=None, Type=None, Amount=None, MinusValues=None):
        arguments = [Direction, Include, Type, Amount, MinusValues]
        return self.ErrorBar(*arguments)

    def Paste(self):
        self.com_object.Paste()

    # Lower case alias for Paste
    def paste(self):
        return self.Paste()

    def Points(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Points(self.com_object.Points(*arguments))

    # Lower case alias for Points
    def points(self, Index=None):
        arguments = [Index]
        return self.Points(*arguments)

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()

    def Trendlines(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Trendlines(self.com_object.Trendlines(*arguments))

    # Lower case alias for Trendlines
    def trendlines(self, Index=None):
        arguments = [Index]
        return self.Trendlines(*arguments)


class SeriesCollection:

    def __init__(self, seriescollection=None):
        self.com_object= seriescollection

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    def Add(self, Source=None, Rowcol=None, SeriesLabels=None, CategoryLabels=None, Replace=None):
        arguments = com_arguments([unwrap(a) for a in [Source, Rowcol, SeriesLabels, CategoryLabels, Replace]])
        return Series(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Source=None, Rowcol=None, SeriesLabels=None, CategoryLabels=None, Replace=None):
        arguments = [Source, Rowcol, SeriesLabels, CategoryLabels, Replace]
        return self.Add(*arguments)

    def Extend(self, Source=None, Rowcol=None, CategoryLabels=None):
        arguments = com_arguments([unwrap(a) for a in [Source, Rowcol, CategoryLabels]])
        self.com_object.Extend(*arguments)

    # Lower case alias for Extend
    def extend(self, Source=None, Rowcol=None, CategoryLabels=None):
        arguments = [Source, Rowcol, CategoryLabels]
        return self.Extend(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Series(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def NewSeries(self):
        return Series(self.com_object.NewSeries())

    # Lower case alias for NewSeries
    def newseries(self):
        return self.NewSeries()


class SeriesLines:

    def __init__(self, serieslines=None):
        self.com_object= serieslines

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Border(self):
        return ChartBorder(self.com_object.Border)

    # Lower case aliases for Border
    @property
    def border(self):
        return self.Border

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

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

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class SetEffect:

    def __init__(self, seteffect=None):
        self.com_object= seteffect

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Property(self):
        return self.com_object.Property

    @Property.setter
    def Property(self, value):
        self.com_object.Property = value

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


class ShadowFormat:

    def __init__(self, shadowformat=None):
        self.com_object= shadowformat

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Blur(self):
        return self.com_object.Blur

    @Blur.setter
    def Blur(self, value):
        self.com_object.Blur = value

    # Lower case aliases for Blur
    @property
    def blur(self):
        return self.Blur

    @blur.setter
    def blur(self, value):
        self.Blur = value

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def ForeColor(self):
        return ColorFormat(self.com_object.ForeColor)

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
    def Obscured(self):
        return self.com_object.Obscured

    @Obscured.setter
    def Obscured(self, value):
        self.com_object.Obscured = value

    # Lower case aliases for Obscured
    @property
    def obscured(self):
        return self.Obscured

    @obscured.setter
    def obscured(self, value):
        self.Obscured = value

    @property
    def OffsetX(self):
        return self.com_object.OffsetX

    @OffsetX.setter
    def OffsetX(self, value):
        self.com_object.OffsetX = value

    # Lower case aliases for OffsetX
    @property
    def offsetx(self):
        return self.OffsetX

    @offsetx.setter
    def offsetx(self, value):
        self.OffsetX = value

    @property
    def OffsetY(self):
        return self.com_object.OffsetY

    @OffsetY.setter
    def OffsetY(self, value):
        self.com_object.OffsetY = value

    # Lower case aliases for OffsetY
    @property
    def offsety(self):
        return self.OffsetY

    @offsety.setter
    def offsety(self, value):
        self.OffsetY = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RotateWithShape(self):
        return self.com_object.RotateWithShape

    @RotateWithShape.setter
    def RotateWithShape(self, value):
        self.com_object.RotateWithShape = value

    # Lower case aliases for RotateWithShape
    @property
    def rotatewithshape(self):
        return self.RotateWithShape

    @rotatewithshape.setter
    def rotatewithshape(self, value):
        self.RotateWithShape = value

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
    def Style(self):
        return self.com_object.Style

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
    def Transparency(self):
        return self.com_object.Transparency

    @Transparency.setter
    def Transparency(self, value):
        self.com_object.Transparency = value

    # Lower case aliases for Transparency
    @property
    def transparency(self):
        return self.Transparency

    @transparency.setter
    def transparency(self, value):
        self.Transparency = value

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

    def IncrementOffsetX(self, Increment=None):
        arguments = com_arguments([unwrap(a) for a in [Increment]])
        self.com_object.IncrementOffsetX(*arguments)

    # Lower case alias for IncrementOffsetX
    def incrementoffsetx(self, Increment=None):
        arguments = [Increment]
        return self.IncrementOffsetX(*arguments)

    def IncrementOffsetY(self, Increment=None):
        arguments = com_arguments([unwrap(a) for a in [Increment]])
        self.com_object.IncrementOffsetY(*arguments)

    # Lower case alias for IncrementOffsetY
    def incrementoffsety(self, Increment=None):
        arguments = [Increment]
        return self.IncrementOffsetY(*arguments)


class Shape:

    def __init__(self, shape=None):
        self.com_object= shape

    @property
    def ActionSettings(self):
        return ActionSettings(self.com_object.ActionSettings)

    # Lower case aliases for ActionSettings
    @property
    def actionsettings(self):
        return self.ActionSettings

    @property
    def Adjustments(self):
        return Adjustments(self.com_object.Adjustments)

    # Lower case aliases for Adjustments
    @property
    def adjustments(self):
        return self.Adjustments

    @property
    def AlternativeText(self):
        return self.com_object.AlternativeText

    @AlternativeText.setter
    def AlternativeText(self, value):
        self.com_object.AlternativeText = value

    # Lower case aliases for AlternativeText
    @property
    def alternativetext(self):
        return self.AlternativeText

    @alternativetext.setter
    def alternativetext(self, value):
        self.AlternativeText = value

    @property
    def AnimationSettings(self):
        return AnimationSettings(self.com_object.AnimationSettings)

    # Lower case aliases for AnimationSettings
    @property
    def animationsettings(self):
        return self.AnimationSettings

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AutoShapeType(self):
        return Shape(self.com_object.AutoShapeType)

    @AutoShapeType.setter
    def AutoShapeType(self, value):
        self.com_object.AutoShapeType = value

    # Lower case aliases for AutoShapeType
    @property
    def autoshapetype(self):
        return self.AutoShapeType

    @autoshapetype.setter
    def autoshapetype(self, value):
        self.AutoShapeType = value

    @property
    def BackgroundStyle(self):
        return self.com_object.BackgroundStyle

    @BackgroundStyle.setter
    def BackgroundStyle(self, value):
        self.com_object.BackgroundStyle = value

    # Lower case aliases for BackgroundStyle
    @property
    def backgroundstyle(self):
        return self.BackgroundStyle

    @backgroundstyle.setter
    def backgroundstyle(self, value):
        self.BackgroundStyle = value

    @property
    def BlackWhiteMode(self):
        return self.com_object.BlackWhiteMode

    @BlackWhiteMode.setter
    def BlackWhiteMode(self, value):
        self.com_object.BlackWhiteMode = value

    # Lower case aliases for BlackWhiteMode
    @property
    def blackwhitemode(self):
        return self.BlackWhiteMode

    @blackwhitemode.setter
    def blackwhitemode(self, value):
        self.BlackWhiteMode = value

    @property
    def Callout(self):
        return CalloutFormat(self.com_object.Callout)

    # Lower case aliases for Callout
    @property
    def callout(self):
        return self.Callout

    @property
    def Chart(self):
        return Chart(self.com_object.Chart)

    # Lower case aliases for Chart
    @property
    def chart(self):
        return self.Chart

    @property
    def Child(self):
        return self.com_object.Child

    # Lower case aliases for Child
    @property
    def child(self):
        return self.Child

    @property
    def ConnectionSiteCount(self):
        return self.com_object.ConnectionSiteCount

    # Lower case aliases for ConnectionSiteCount
    @property
    def connectionsitecount(self):
        return self.ConnectionSiteCount

    @property
    def Connector(self):
        return self.com_object.Connector

    # Lower case aliases for Connector
    @property
    def connector(self):
        return self.Connector

    @property
    def ConnectorFormat(self):
        return ConnectorFormat(self.com_object.ConnectorFormat)

    # Lower case aliases for ConnectorFormat
    @property
    def connectorformat(self):
        return self.ConnectorFormat

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def CustomerData(self):
        return CustomerData(self.com_object.CustomerData)

    # Lower case aliases for CustomerData
    @property
    def customerdata(self):
        return self.CustomerData

    @property
    def Decorative(self):
        return self.com_object.Decorative

    @Decorative.setter
    def Decorative(self, value):
        self.com_object.Decorative = value

    # Lower case aliases for Decorative
    @property
    def decorative(self):
        return self.Decorative

    @decorative.setter
    def decorative(self, value):
        self.Decorative = value

    @property
    def Fill(self):
        return FillFormat(self.com_object.Fill)

    # Lower case aliases for Fill
    @property
    def fill(self):
        return self.Fill

    @property
    def Glow(self):
        return self.com_object.Glow

    # Lower case aliases for Glow
    @property
    def glow(self):
        return self.Glow

    @property
    def GraphicStyle(self):
        return self.com_object.GraphicStyle

    @GraphicStyle.setter
    def GraphicStyle(self, value):
        self.com_object.GraphicStyle = value

    # Lower case aliases for GraphicStyle
    @property
    def graphicstyle(self):
        return self.GraphicStyle

    @graphicstyle.setter
    def graphicstyle(self, value):
        self.GraphicStyle = value

    @property
    def GroupItems(self):
        return GroupShapes(self.com_object.GroupItems)

    # Lower case aliases for GroupItems
    @property
    def groupitems(self):
        return self.GroupItems

    @property
    def HasChart(self):
        return self.com_object.HasChart

    # Lower case aliases for HasChart
    @property
    def haschart(self):
        return self.HasChart

    @property
    def HasSmartArt(self):
        return self.com_object.HasSmartArt

    # Lower case aliases for HasSmartArt
    @property
    def hassmartart(self):
        return self.HasSmartArt

    @property
    def HasTable(self):
        return self.com_object.HasTable

    # Lower case aliases for HasTable
    @property
    def hastable(self):
        return self.HasTable

    @property
    def HasTextFrame(self):
        return self.com_object.HasTextFrame

    # Lower case aliases for HasTextFrame
    @property
    def hastextframe(self):
        return self.HasTextFrame

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
    def HorizontalFlip(self):
        return self.com_object.HorizontalFlip

    # Lower case aliases for HorizontalFlip
    @property
    def horizontalflip(self):
        return self.HorizontalFlip

    @property
    def Id(self):
        return self.com_object.Id

    # Lower case aliases for Id
    @property
    def id(self):
        return self.Id

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
    def Line(self):
        return LineFormat(self.com_object.Line)

    # Lower case aliases for Line
    @property
    def line(self):
        return self.Line

    @property
    def LinkFormat(self):
        return LinkFormat(self.com_object.LinkFormat)

    # Lower case aliases for LinkFormat
    @property
    def linkformat(self):
        return self.LinkFormat

    @property
    def LockAspectRatio(self):
        return self.com_object.LockAspectRatio

    @LockAspectRatio.setter
    def LockAspectRatio(self, value):
        self.com_object.LockAspectRatio = value

    # Lower case aliases for LockAspectRatio
    @property
    def lockaspectratio(self):
        return self.LockAspectRatio

    @lockaspectratio.setter
    def lockaspectratio(self, value):
        self.LockAspectRatio = value

    @property
    def MediaFormat(self):
        return self.com_object.MediaFormat

    # Lower case aliases for MediaFormat
    @property
    def mediaformat(self):
        return self.MediaFormat

    @property
    def MediaType(self):
        return self.com_object.MediaType

    # Lower case aliases for MediaType
    @property
    def mediatype(self):
        return self.MediaType

    @property
    def Model3D(self):
        return Model3DFormat(self.com_object.Model3D)

    # Lower case aliases for Model3D
    @property
    def model3d(self):
        return self.Model3D

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
    def Nodes(self):
        return ShapeNodes(self.com_object.Nodes)

    # Lower case aliases for Nodes
    @property
    def nodes(self):
        return self.Nodes

    @property
    def OLEFormat(self):
        return OLEFormat(self.com_object.OLEFormat)

    # Lower case aliases for OLEFormat
    @property
    def oleformat(self):
        return self.OLEFormat

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ParentGroup(self):
        return Shape(self.com_object.ParentGroup)

    # Lower case aliases for ParentGroup
    @property
    def parentgroup(self):
        return self.ParentGroup

    @property
    def PictureFormat(self):
        return PictureFormat(self.com_object.PictureFormat)

    # Lower case aliases for PictureFormat
    @property
    def pictureformat(self):
        return self.PictureFormat

    @property
    def PlaceholderFormat(self):
        return PlaceholderFormat(self.com_object.PlaceholderFormat)

    # Lower case aliases for PlaceholderFormat
    @property
    def placeholderformat(self):
        return self.PlaceholderFormat

    @property
    def Reflection(self):
        return self.com_object.Reflection

    # Lower case aliases for Reflection
    @property
    def reflection(self):
        return self.Reflection

    @property
    def Rotation(self):
        return self.com_object.Rotation

    @Rotation.setter
    def Rotation(self, value):
        self.com_object.Rotation = value

    # Lower case aliases for Rotation
    @property
    def rotation(self):
        return self.Rotation

    @rotation.setter
    def rotation(self, value):
        self.Rotation = value

    @property
    def Shadow(self):
        return ShadowFormat(self.com_object.Shadow)

    # Lower case aliases for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @property
    def ShapeStyle(self):
        return self.com_object.ShapeStyle

    @ShapeStyle.setter
    def ShapeStyle(self, value):
        self.com_object.ShapeStyle = value

    # Lower case aliases for ShapeStyle
    @property
    def shapestyle(self):
        return self.ShapeStyle

    @shapestyle.setter
    def shapestyle(self, value):
        self.ShapeStyle = value

    @property
    def SmartArt(self):
        return Shape(self.com_object.SmartArt)

    # Lower case aliases for SmartArt
    @property
    def smartart(self):
        return self.SmartArt

    @property
    def SoftEdge(self):
        return self.com_object.SoftEdge

    # Lower case aliases for SoftEdge
    @property
    def softedge(self):
        return self.SoftEdge

    @property
    def Table(self):
        return Table(self.com_object.Table)

    # Lower case aliases for Table
    @property
    def table(self):
        return self.Table

    @property
    def Tags(self):
        return Tags(self.com_object.Tags)

    # Lower case aliases for Tags
    @property
    def tags(self):
        return self.Tags

    @property
    def TextEffect(self):
        return TextEffectFormat(self.com_object.TextEffect)

    # Lower case aliases for TextEffect
    @property
    def texteffect(self):
        return self.TextEffect

    @property
    def TextFrame(self):
        return TextFrame(self.com_object.TextFrame)

    # Lower case aliases for TextFrame
    @property
    def textframe(self):
        return self.TextFrame

    @property
    def TextFrame2(self):
        return TextFrame2(self.com_object.TextFrame2)

    # Lower case aliases for TextFrame2
    @property
    def textframe2(self):
        return self.TextFrame2

    @property
    def ThreeD(self):
        return ThreeDFormat(self.com_object.ThreeD)

    # Lower case aliases for ThreeD
    @property
    def threed(self):
        return self.ThreeD

    @property
    def Title(self):
        return Shape(self.com_object.Title)

    # Lower case aliases for Title
    @property
    def title(self):
        return self.Title

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
    def Type(self):
        return self.com_object.Type

    # Lower case aliases for Type
    @property
    def type(self):
        return self.Type

    @property
    def VerticalFlip(self):
        return self.com_object.VerticalFlip

    # Lower case aliases for VerticalFlip
    @property
    def verticalflip(self):
        return self.VerticalFlip

    @property
    def Vertices(self):
        return self.com_object.Vertices

    # Lower case aliases for Vertices
    @property
    def vertices(self):
        return self.Vertices

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
    def ZOrderPosition(self):
        return self.com_object.ZOrderPosition

    # Lower case aliases for ZOrderPosition
    @property
    def zorderposition(self):
        return self.ZOrderPosition

    def Apply(self):
        self.com_object.Apply()

    # Lower case alias for Apply
    def apply(self):
        return self.Apply()

    def ApplyAnimation(self):
        self.com_object.ApplyAnimation()

    # Lower case alias for ApplyAnimation
    def applyanimation(self):
        return self.ApplyAnimation()

    def ConvertTextToSmartArt(self, Layout=None):
        arguments = com_arguments([unwrap(a) for a in [Layout]])
        self.com_object.ConvertTextToSmartArt(*arguments)

    # Lower case alias for ConvertTextToSmartArt
    def converttexttosmartart(self, Layout=None):
        arguments = [Layout]
        return self.ConvertTextToSmartArt(*arguments)

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

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Duplicate(self):
        return ShapeRange(self.com_object.Duplicate())

    # Lower case alias for Duplicate
    def duplicate(self):
        return self.Duplicate()

    def Flip(self, FlipCmd=None):
        arguments = com_arguments([unwrap(a) for a in [FlipCmd]])
        self.com_object.Flip(*arguments)

    # Lower case alias for Flip
    def flip(self, FlipCmd=None):
        arguments = [FlipCmd]
        return self.Flip(*arguments)

    def IncrementLeft(self, Increment=None):
        arguments = com_arguments([unwrap(a) for a in [Increment]])
        self.com_object.IncrementLeft(*arguments)

    # Lower case alias for IncrementLeft
    def incrementleft(self, Increment=None):
        arguments = [Increment]
        return self.IncrementLeft(*arguments)

    def IncrementRotation(self, Increment=None):
        arguments = com_arguments([unwrap(a) for a in [Increment]])
        self.com_object.IncrementRotation(*arguments)

    # Lower case alias for IncrementRotation
    def incrementrotation(self, Increment=None):
        arguments = [Increment]
        return self.IncrementRotation(*arguments)

    def IncrementTop(self, Increment=None):
        arguments = com_arguments([unwrap(a) for a in [Increment]])
        self.com_object.IncrementTop(*arguments)

    # Lower case alias for IncrementTop
    def incrementtop(self, Increment=None):
        arguments = [Increment]
        return self.IncrementTop(*arguments)

    def PickUp(self):
        self.com_object.PickUp()

    # Lower case alias for PickUp
    def pickup(self):
        return self.PickUp()

    def PickupAnimation(self):
        self.com_object.PickupAnimation()

    # Lower case alias for PickupAnimation
    def pickupanimation(self):
        return self.PickupAnimation()

    def RerouteConnections(self):
        self.com_object.RerouteConnections()

    # Lower case alias for RerouteConnections
    def rerouteconnections(self):
        return self.RerouteConnections()

    def ScaleHeight(self, Factor=None, RelativeToOriginalSize=None, fScale=None):
        arguments = com_arguments([unwrap(a) for a in [Factor, RelativeToOriginalSize, fScale]])
        self.com_object.ScaleHeight(*arguments)

    # Lower case alias for ScaleHeight
    def scaleheight(self, Factor=None, RelativeToOriginalSize=None, fScale=None):
        arguments = [Factor, RelativeToOriginalSize, fScale]
        return self.ScaleHeight(*arguments)

    def ScaleWidth(self, Factor=None, RelativeToOriginalSize=None, fScale=None):
        arguments = com_arguments([unwrap(a) for a in [Factor, RelativeToOriginalSize, fScale]])
        self.com_object.ScaleWidth(*arguments)

    # Lower case alias for ScaleWidth
    def scalewidth(self, Factor=None, RelativeToOriginalSize=None, fScale=None):
        arguments = [Factor, RelativeToOriginalSize, fScale]
        return self.ScaleWidth(*arguments)

    def Select(self, Replace=None):
        arguments = com_arguments([unwrap(a) for a in [Replace]])
        self.com_object.Select(*arguments)

    # Lower case alias for Select
    def select(self, Replace=None):
        arguments = [Replace]
        return self.Select(*arguments)

    def SetShapesDefaultProperties(self):
        self.com_object.SetShapesDefaultProperties()

    # Lower case alias for SetShapesDefaultProperties
    def setshapesdefaultproperties(self):
        return self.SetShapesDefaultProperties()

    def Ungroup(self):
        return ShapeRange(self.com_object.Ungroup())

    # Lower case alias for Ungroup
    def ungroup(self):
        return self.Ungroup()

    def UpgradeMedia(self):
        self.com_object.UpgradeMedia()

    # Lower case alias for UpgradeMedia
    def upgrademedia(self):
        return self.UpgradeMedia()

    def ZOrder(self, ZOrderCmd=None):
        arguments = com_arguments([unwrap(a) for a in [ZOrderCmd]])
        self.com_object.ZOrder(*arguments)

    # Lower case alias for ZOrder
    def zorder(self, ZOrderCmd=None):
        arguments = [ZOrderCmd]
        return self.ZOrder(*arguments)


class ShapeNode:

    def __init__(self, shapenode=None):
        self.com_object= shapenode

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def EditingType(self):
        return self.com_object.EditingType

    # Lower case aliases for EditingType
    @property
    def editingtype(self):
        return self.EditingType

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Points(self):
        return self.com_object.Points

    # Lower case aliases for Points
    @property
    def points(self):
        return self.Points

    @property
    def SegmentType(self):
        return self.com_object.SegmentType

    # Lower case aliases for SegmentType
    @property
    def segmenttype(self):
        return self.SegmentType


class ShapeNodes:

    def __init__(self, shapenodes=None):
        self.com_object= shapenodes

    def __call__(self, item):
        return ShapeNode(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    def Delete(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        self.com_object.Delete(*arguments)

    # Lower case alias for Delete
    def delete(self, Index=None):
        arguments = [Index]
        return self.Delete(*arguments)

    def Insert(self, Index=None, SegmentType=None, EditingType=None, X1=None, Y1=None, X2=None, Y2=None, X3=None, Y3=None):
        arguments = com_arguments([unwrap(a) for a in [Index, SegmentType, EditingType, X1, Y1, X2, Y2, X3, Y3]])
        self.com_object.Insert(*arguments)

    # Lower case alias for Insert
    def insert(self, Index=None, SegmentType=None, EditingType=None, X1=None, Y1=None, X2=None, Y2=None, X3=None, Y3=None):
        arguments = [Index, SegmentType, EditingType, X1, Y1, X2, Y2, X3, Y3]
        return self.Insert(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return ShapeNode(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def SetEditingType(self, Index=None, EditingType=None):
        arguments = com_arguments([unwrap(a) for a in [Index, EditingType]])
        self.com_object.SetEditingType(*arguments)

    # Lower case alias for SetEditingType
    def seteditingtype(self, Index=None, EditingType=None):
        arguments = [Index, EditingType]
        return self.SetEditingType(*arguments)

    def SetPosition(self, Index=None, X1=None, Y1=None):
        arguments = com_arguments([unwrap(a) for a in [Index, X1, Y1]])
        self.com_object.SetPosition(*arguments)

    # Lower case alias for SetPosition
    def setposition(self, Index=None, X1=None, Y1=None):
        arguments = [Index, X1, Y1]
        return self.SetPosition(*arguments)

    def SetSegmentType(self, Index=None, SegmentType=None):
        arguments = com_arguments([unwrap(a) for a in [Index, SegmentType]])
        self.com_object.SetSegmentType(*arguments)

    # Lower case alias for SetSegmentType
    def setsegmenttype(self, Index=None, SegmentType=None):
        arguments = [Index, SegmentType]
        return self.SetSegmentType(*arguments)


class ShapeRange:

    def __init__(self, shaperange=None):
        self.com_object= shaperange

    @property
    def ActionSettings(self):
        return ActionSettings(self.com_object.ActionSettings)

    # Lower case aliases for ActionSettings
    @property
    def actionsettings(self):
        return self.ActionSettings

    @property
    def Adjustments(self):
        return Adjustments(self.com_object.Adjustments)

    # Lower case aliases for Adjustments
    @property
    def adjustments(self):
        return self.Adjustments

    @property
    def AlternativeText(self):
        return self.com_object.AlternativeText

    @AlternativeText.setter
    def AlternativeText(self, value):
        self.com_object.AlternativeText = value

    # Lower case aliases for AlternativeText
    @property
    def alternativetext(self):
        return self.AlternativeText

    @alternativetext.setter
    def alternativetext(self, value):
        self.AlternativeText = value

    @property
    def AnimationSettings(self):
        return AnimationSettings(self.com_object.AnimationSettings)

    # Lower case aliases for AnimationSettings
    @property
    def animationsettings(self):
        return self.AnimationSettings

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AutoShapeType(self):
        return ShapeRange(self.com_object.AutoShapeType)

    @AutoShapeType.setter
    def AutoShapeType(self, value):
        self.com_object.AutoShapeType = value

    # Lower case aliases for AutoShapeType
    @property
    def autoshapetype(self):
        return self.AutoShapeType

    @autoshapetype.setter
    def autoshapetype(self, value):
        self.AutoShapeType = value

    @property
    def BackgroundStyle(self):
        return self.com_object.BackgroundStyle

    @BackgroundStyle.setter
    def BackgroundStyle(self, value):
        self.com_object.BackgroundStyle = value

    # Lower case aliases for BackgroundStyle
    @property
    def backgroundstyle(self):
        return self.BackgroundStyle

    @backgroundstyle.setter
    def backgroundstyle(self, value):
        self.BackgroundStyle = value

    @property
    def BlackWhiteMode(self):
        return self.com_object.BlackWhiteMode

    @BlackWhiteMode.setter
    def BlackWhiteMode(self, value):
        self.com_object.BlackWhiteMode = value

    # Lower case aliases for BlackWhiteMode
    @property
    def blackwhitemode(self):
        return self.BlackWhiteMode

    @blackwhitemode.setter
    def blackwhitemode(self, value):
        self.BlackWhiteMode = value

    @property
    def Callout(self):
        return CalloutFormat(self.com_object.Callout)

    # Lower case aliases for Callout
    @property
    def callout(self):
        return self.Callout

    @property
    def Chart(self):
        return Chart(self.com_object.Chart)

    # Lower case aliases for Chart
    @property
    def chart(self):
        return self.Chart

    @property
    def Child(self):
        return self.com_object.Child

    # Lower case aliases for Child
    @property
    def child(self):
        return self.Child

    @property
    def ConnectionSiteCount(self):
        return self.com_object.ConnectionSiteCount

    # Lower case aliases for ConnectionSiteCount
    @property
    def connectionsitecount(self):
        return self.ConnectionSiteCount

    @property
    def Connector(self):
        return self.com_object.Connector

    # Lower case aliases for Connector
    @property
    def connector(self):
        return self.Connector

    @property
    def ConnectorFormat(self):
        return ConnectorFormat(self.com_object.ConnectorFormat)

    # Lower case aliases for ConnectorFormat
    @property
    def connectorformat(self):
        return self.ConnectorFormat

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def CustomerData(self):
        return CustomerData(self.com_object.CustomerData)

    # Lower case aliases for CustomerData
    @property
    def customerdata(self):
        return self.CustomerData

    @property
    def Decorative(self):
        return self.com_object.Decorative

    @Decorative.setter
    def Decorative(self, value):
        self.com_object.Decorative = value

    # Lower case aliases for Decorative
    @property
    def decorative(self):
        return self.Decorative

    @decorative.setter
    def decorative(self, value):
        self.Decorative = value

    @property
    def Fill(self):
        return FillFormat(self.com_object.Fill)

    # Lower case aliases for Fill
    @property
    def fill(self):
        return self.Fill

    @property
    def Glow(self):
        return self.com_object.Glow

    # Lower case aliases for Glow
    @property
    def glow(self):
        return self.Glow

    @property
    def GraphicStyle(self):
        return self.com_object.GraphicStyle

    @GraphicStyle.setter
    def GraphicStyle(self, value):
        self.com_object.GraphicStyle = value

    # Lower case aliases for GraphicStyle
    @property
    def graphicstyle(self):
        return self.GraphicStyle

    @graphicstyle.setter
    def graphicstyle(self, value):
        self.GraphicStyle = value

    @property
    def GroupItems(self):
        return GroupShapes(self.com_object.GroupItems)

    # Lower case aliases for GroupItems
    @property
    def groupitems(self):
        return self.GroupItems

    @property
    def HasChart(self):
        return self.com_object.HasChart

    # Lower case aliases for HasChart
    @property
    def haschart(self):
        return self.HasChart

    @property
    def HasSmartArt(self):
        return self.com_object.HasSmartArt

    # Lower case aliases for HasSmartArt
    @property
    def hassmartart(self):
        return self.HasSmartArt

    @property
    def HasTable(self):
        return self.com_object.HasTable

    # Lower case aliases for HasTable
    @property
    def hastable(self):
        return self.HasTable

    @property
    def HasTextFrame(self):
        return self.com_object.HasTextFrame

    # Lower case aliases for HasTextFrame
    @property
    def hastextframe(self):
        return self.HasTextFrame

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
    def HorizontalFlip(self):
        return self.com_object.HorizontalFlip

    # Lower case aliases for HorizontalFlip
    @property
    def horizontalflip(self):
        return self.HorizontalFlip

    @property
    def Id(self):
        return self.com_object.Id

    # Lower case aliases for Id
    @property
    def id(self):
        return self.Id

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
    def Line(self):
        return LineFormat(self.com_object.Line)

    # Lower case aliases for Line
    @property
    def line(self):
        return self.Line

    @property
    def LinkFormat(self):
        return LinkFormat(self.com_object.LinkFormat)

    # Lower case aliases for LinkFormat
    @property
    def linkformat(self):
        return self.LinkFormat

    @property
    def LockAspectRatio(self):
        return self.com_object.LockAspectRatio

    @LockAspectRatio.setter
    def LockAspectRatio(self, value):
        self.com_object.LockAspectRatio = value

    # Lower case aliases for LockAspectRatio
    @property
    def lockaspectratio(self):
        return self.LockAspectRatio

    @lockaspectratio.setter
    def lockaspectratio(self, value):
        self.LockAspectRatio = value

    @property
    def MediaFormat(self):
        return MediaFormat(self.com_object.MediaFormat)

    # Lower case aliases for MediaFormat
    @property
    def mediaformat(self):
        return self.MediaFormat

    @property
    def MediaType(self):
        return self.com_object.MediaType

    # Lower case aliases for MediaType
    @property
    def mediatype(self):
        return self.MediaType

    @property
    def Model3D(self):
        return Model3DFormat(self.com_object.Model3D)

    # Lower case aliases for Model3D
    @property
    def model3d(self):
        return self.Model3D

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
    def Nodes(self):
        return ShapeNodes(self.com_object.Nodes)

    # Lower case aliases for Nodes
    @property
    def nodes(self):
        return self.Nodes

    @property
    def OLEFormat(self):
        return OLEFormat(self.com_object.OLEFormat)

    # Lower case aliases for OLEFormat
    @property
    def oleformat(self):
        return self.OLEFormat

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ParentGroup(self):
        return Shape(self.com_object.ParentGroup)

    # Lower case aliases for ParentGroup
    @property
    def parentgroup(self):
        return self.ParentGroup

    @property
    def PictureFormat(self):
        return PictureFormat(self.com_object.PictureFormat)

    # Lower case aliases for PictureFormat
    @property
    def pictureformat(self):
        return self.PictureFormat

    @property
    def PlaceholderFormat(self):
        return PlaceholderFormat(self.com_object.PlaceholderFormat)

    # Lower case aliases for PlaceholderFormat
    @property
    def placeholderformat(self):
        return self.PlaceholderFormat

    @property
    def Reflection(self):
        return self.com_object.Reflection

    # Lower case aliases for Reflection
    @property
    def reflection(self):
        return self.Reflection

    @property
    def Rotation(self):
        return self.com_object.Rotation

    @Rotation.setter
    def Rotation(self, value):
        self.com_object.Rotation = value

    # Lower case aliases for Rotation
    @property
    def rotation(self):
        return self.Rotation

    @rotation.setter
    def rotation(self, value):
        self.Rotation = value

    @property
    def Shadow(self):
        return ShadowFormat(self.com_object.Shadow)

    # Lower case aliases for Shadow
    @property
    def shadow(self):
        return self.Shadow

    @property
    def ShapeStyle(self):
        return self.com_object.ShapeStyle

    # Lower case aliases for ShapeStyle
    @property
    def shapestyle(self):
        return self.ShapeStyle

    @property
    def SmartArt(self):
        return ShapeRange(self.com_object.SmartArt)

    # Lower case aliases for SmartArt
    @property
    def smartart(self):
        return self.SmartArt

    @property
    def SoftEdge(self):
        return self.com_object.SoftEdge

    # Lower case aliases for SoftEdge
    @property
    def softedge(self):
        return self.SoftEdge

    @property
    def Table(self):
        return Table(self.com_object.Table)

    # Lower case aliases for Table
    @property
    def table(self):
        return self.Table

    @property
    def Tags(self):
        return Tags(self.com_object.Tags)

    # Lower case aliases for Tags
    @property
    def tags(self):
        return self.Tags

    @property
    def TextEffect(self):
        return TextEffectFormat(self.com_object.TextEffect)

    # Lower case aliases for TextEffect
    @property
    def texteffect(self):
        return self.TextEffect

    @property
    def TextFrame(self):
        return TextFrame(self.com_object.TextFrame)

    # Lower case aliases for TextFrame
    @property
    def textframe(self):
        return self.TextFrame

    @property
    def TextFrame2(self):
        return TextFrame2(self.com_object.TextFrame2)

    # Lower case aliases for TextFrame2
    @property
    def textframe2(self):
        return self.TextFrame2

    @property
    def ThreeD(self):
        return ThreeDFormat(self.com_object.ThreeD)

    # Lower case aliases for ThreeD
    @property
    def threed(self):
        return self.ThreeD

    @property
    def Title(self):
        return Shape(self.com_object.Title)

    # Lower case aliases for Title
    @property
    def title(self):
        return self.Title

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
    def Type(self):
        return self.com_object.Type

    # Lower case aliases for Type
    @property
    def type(self):
        return self.Type

    @property
    def VerticalFlip(self):
        return self.com_object.VerticalFlip

    # Lower case aliases for VerticalFlip
    @property
    def verticalflip(self):
        return self.VerticalFlip

    @property
    def Vertices(self):
        return self.com_object.Vertices

    # Lower case aliases for Vertices
    @property
    def vertices(self):
        return self.Vertices

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
    def ZOrderPosition(self):
        return self.com_object.ZOrderPosition

    # Lower case aliases for ZOrderPosition
    @property
    def zorderposition(self):
        return self.ZOrderPosition

    def Align(self, AlignCmd=None, RelativeTo=None):
        arguments = com_arguments([unwrap(a) for a in [AlignCmd, RelativeTo]])
        self.com_object.Align(*arguments)

    # Lower case alias for Align
    def align(self, AlignCmd=None, RelativeTo=None):
        arguments = [AlignCmd, RelativeTo]
        return self.Align(*arguments)

    def Apply(self):
        self.com_object.Apply()

    # Lower case alias for Apply
    def apply(self):
        return self.Apply()

    def ApplyAnimation(self):
        self.com_object.ApplyAnimation()

    # Lower case alias for ApplyAnimation
    def applyanimation(self):
        return self.ApplyAnimation()

    def ConvertTextToSmartArt(self, Layout=None):
        arguments = com_arguments([unwrap(a) for a in [Layout]])
        return self.com_object.ConvertTextToSmartArt(*arguments)

    # Lower case alias for ConvertTextToSmartArt
    def converttexttosmartart(self, Layout=None):
        arguments = [Layout]
        return self.ConvertTextToSmartArt(*arguments)

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

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Distribute(self, DistributeCmd=None, RelativeTo=None):
        arguments = com_arguments([unwrap(a) for a in [DistributeCmd, RelativeTo]])
        return self.com_object.Distribute(*arguments)

    # Lower case alias for Distribute
    def distribute(self, DistributeCmd=None, RelativeTo=None):
        arguments = [DistributeCmd, RelativeTo]
        return self.Distribute(*arguments)

    def Duplicate(self):
        return ShapeRange(self.com_object.Duplicate())

    # Lower case alias for Duplicate
    def duplicate(self):
        return self.Duplicate()

    def Flip(self, FlipCmd=None):
        arguments = com_arguments([unwrap(a) for a in [FlipCmd]])
        self.com_object.Flip(*arguments)

    # Lower case alias for Flip
    def flip(self, FlipCmd=None):
        arguments = [FlipCmd]
        return self.Flip(*arguments)

    def Group(self):
        return Shape(self.com_object.Group())

    # Lower case alias for Group
    def group(self):
        return self.Group()

    def IncrementLeft(self, Increment=None):
        arguments = com_arguments([unwrap(a) for a in [Increment]])
        self.com_object.IncrementLeft(*arguments)

    # Lower case alias for IncrementLeft
    def incrementleft(self, Increment=None):
        arguments = [Increment]
        return self.IncrementLeft(*arguments)

    def IncrementRotation(self, Increment=None):
        arguments = com_arguments([unwrap(a) for a in [Increment]])
        self.com_object.IncrementRotation(*arguments)

    # Lower case alias for IncrementRotation
    def incrementrotation(self, Increment=None):
        arguments = [Increment]
        return self.IncrementRotation(*arguments)

    def IncrementTop(self, Increment=None):
        arguments = com_arguments([unwrap(a) for a in [Increment]])
        self.com_object.IncrementTop(*arguments)

    # Lower case alias for IncrementTop
    def incrementtop(self, Increment=None):
        arguments = [Increment]
        return self.IncrementTop(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Shape(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def PickUp(self):
        self.com_object.PickUp()

    # Lower case alias for PickUp
    def pickup(self):
        return self.PickUp()

    def PickupAnimation(self):
        self.com_object.PickupAnimation()

    # Lower case alias for PickupAnimation
    def pickupanimation(self):
        return self.PickupAnimation()

    def Regroup(self):
        return Shape(self.com_object.Regroup())

    # Lower case alias for Regroup
    def regroup(self):
        return self.Regroup()

    def RerouteConnections(self):
        self.com_object.RerouteConnections()

    # Lower case alias for RerouteConnections
    def rerouteconnections(self):
        return self.RerouteConnections()

    def ScaleHeight(self, Factor=None, RelativeToOriginalSize=None, fScale=None):
        arguments = com_arguments([unwrap(a) for a in [Factor, RelativeToOriginalSize, fScale]])
        return self.com_object.ScaleHeight(*arguments)

    # Lower case alias for ScaleHeight
    def scaleheight(self, Factor=None, RelativeToOriginalSize=None, fScale=None):
        arguments = [Factor, RelativeToOriginalSize, fScale]
        return self.ScaleHeight(*arguments)

    def ScaleWidth(self, Factor=None, RelativeToOriginalSize=None, fScale=None):
        arguments = com_arguments([unwrap(a) for a in [Factor, RelativeToOriginalSize, fScale]])
        self.com_object.ScaleWidth(*arguments)

    # Lower case alias for ScaleWidth
    def scalewidth(self, Factor=None, RelativeToOriginalSize=None, fScale=None):
        arguments = [Factor, RelativeToOriginalSize, fScale]
        return self.ScaleWidth(*arguments)

    def Select(self, Replace=None):
        arguments = com_arguments([unwrap(a) for a in [Replace]])
        self.com_object.Select(*arguments)

    # Lower case alias for Select
    def select(self, Replace=None):
        arguments = [Replace]
        return self.Select(*arguments)

    def SetShapesDefaultProperties(self):
        self.com_object.SetShapesDefaultProperties()

    # Lower case alias for SetShapesDefaultProperties
    def setshapesdefaultproperties(self):
        return self.SetShapesDefaultProperties()

    def Ungroup(self):
        return ShapeRange(self.com_object.Ungroup())

    # Lower case alias for Ungroup
    def ungroup(self):
        return self.Ungroup()

    def UpgradeMedia(self):
        self.com_object.UpgradeMedia()

    # Lower case alias for UpgradeMedia
    def upgrademedia(self):
        return self.UpgradeMedia()

    def ZOrder(self, ZOrderCmd=None):
        arguments = com_arguments([unwrap(a) for a in [ZOrderCmd]])
        self.com_object.ZOrder(*arguments)

    # Lower case alias for ZOrder
    def zorder(self, ZOrderCmd=None):
        arguments = [ZOrderCmd]
        return self.ZOrder(*arguments)


class Shapes:

    def __init__(self, shapes=None):
        self.com_object= shapes

    def __call__(self, item):
        return Shape(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def HasTitle(self):
        return self.com_object.HasTitle

    # Lower case aliases for HasTitle
    @property
    def hastitle(self):
        return self.HasTitle

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Placeholders(self):
        return Placeholders(self.com_object.Placeholders)

    # Lower case aliases for Placeholders
    @property
    def placeholders(self):
        return self.Placeholders

    @property
    def Title(self):
        return Shape(self.com_object.Title)

    # Lower case aliases for Title
    @property
    def title(self):
        return self.Title

    def Add3DModel(self, FileName=None, LinkToFile=None, SaveWithDocument=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([unwrap(a) for a in [FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height]])
        return Shape(self.com_object.Add3DModel(*arguments))

    # Lower case alias for Add3DModel
    def add3dmodel(self, FileName=None, LinkToFile=None, SaveWithDocument=None, Left=None, Top=None, Width=None, Height=None):
        arguments = [FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height]
        return self.Add3DModel(*arguments)

    def AddCallout(self, Type=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([unwrap(a) for a in [Type, Left, Top, Width, Height]])
        return Shape(self.com_object.AddCallout(*arguments))

    # Lower case alias for AddCallout
    def addcallout(self, Type=None, Left=None, Top=None, Width=None, Height=None):
        arguments = [Type, Left, Top, Width, Height]
        return self.AddCallout(*arguments)

    def AddConnector(self, Type=None, BeginX=None, BeginY=None, EndX=None, EndY=None):
        arguments = com_arguments([unwrap(a) for a in [Type, BeginX, BeginY, EndX, EndY]])
        return Shape(self.com_object.AddConnector(*arguments))

    # Lower case alias for AddConnector
    def addconnector(self, Type=None, BeginX=None, BeginY=None, EndX=None, EndY=None):
        arguments = [Type, BeginX, BeginY, EndX, EndY]
        return self.AddConnector(*arguments)

    def AddCurve(self, SafeArrayOfPoints=None):
        arguments = com_arguments([unwrap(a) for a in [SafeArrayOfPoints]])
        return Shape(self.com_object.AddCurve(*arguments))

    # Lower case alias for AddCurve
    def addcurve(self, SafeArrayOfPoints=None):
        arguments = [SafeArrayOfPoints]
        return self.AddCurve(*arguments)

    def AddLabel(self, Orientation=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([unwrap(a) for a in [Orientation, Left, Top, Width, Height]])
        return Shape(self.com_object.AddLabel(*arguments))

    # Lower case alias for AddLabel
    def addlabel(self, Orientation=None, Left=None, Top=None, Width=None, Height=None):
        arguments = [Orientation, Left, Top, Width, Height]
        return self.AddLabel(*arguments)

    def AddLine(self, BeginX=None, BeginY=None, EndX=None, EndY=None):
        arguments = com_arguments([unwrap(a) for a in [BeginX, BeginY, EndX, EndY]])
        return Shape(self.com_object.AddLine(*arguments))

    # Lower case alias for AddLine
    def addline(self, BeginX=None, BeginY=None, EndX=None, EndY=None):
        arguments = [BeginX, BeginY, EndX, EndY]
        return self.AddLine(*arguments)

    def AddMediaObject(self, FileName=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([unwrap(a) for a in [FileName, Left, Top, Width, Height]])
        return Shape(self.com_object.AddMediaObject(*arguments))

    # Lower case alias for AddMediaObject
    def addmediaobject(self, FileName=None, Left=None, Top=None, Width=None, Height=None):
        arguments = [FileName, Left, Top, Width, Height]
        return self.AddMediaObject(*arguments)

    def AddMediaObject2(self, FileName=None, LinkToFile=None, SaveWithDocument=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([unwrap(a) for a in [FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height]])
        return Shape(self.com_object.AddMediaObject2(*arguments))

    # Lower case alias for AddMediaObject2
    def addmediaobject2(self, FileName=None, LinkToFile=None, SaveWithDocument=None, Left=None, Top=None, Width=None, Height=None):
        arguments = [FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height]
        return self.AddMediaObject2(*arguments)

    def AddMediaObjectFromEmbedTag(self, EmbedTag=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([unwrap(a) for a in [EmbedTag, Left, Top, Width, Height]])
        return Shape(self.com_object.AddMediaObjectFromEmbedTag(*arguments))

    # Lower case alias for AddMediaObjectFromEmbedTag
    def addmediaobjectfromembedtag(self, EmbedTag=None, Left=None, Top=None, Width=None, Height=None):
        arguments = [EmbedTag, Left, Top, Width, Height]
        return self.AddMediaObjectFromEmbedTag(*arguments)

    def AddOLEObject(self, Left=None, Top=None, Width=None, Height=None, ClassName=None, FileName=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Link=None):
        arguments = com_arguments([unwrap(a) for a in [Left, Top, Width, Height, ClassName, FileName, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Link]])
        return Shape(self.com_object.AddOLEObject(*arguments))

    # Lower case alias for AddOLEObject
    def addoleobject(self, Left=None, Top=None, Width=None, Height=None, ClassName=None, FileName=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Link=None):
        arguments = [Left, Top, Width, Height, ClassName, FileName, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Link]
        return self.AddOLEObject(*arguments)

    def AddPicture(self, FileName=None, LinkToFile=None, SaveWithDocument=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([unwrap(a) for a in [FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height]])
        return Shape(self.com_object.AddPicture(*arguments))

    # Lower case alias for AddPicture
    def addpicture(self, FileName=None, LinkToFile=None, SaveWithDocument=None, Left=None, Top=None, Width=None, Height=None):
        arguments = [FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height]
        return self.AddPicture(*arguments)

    def AddPlaceholder(self, Type=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([unwrap(a) for a in [Type, Left, Top, Width, Height]])
        return Shape(self.com_object.AddPlaceholder(*arguments))

    # Lower case alias for AddPlaceholder
    def addplaceholder(self, Type=None, Left=None, Top=None, Width=None, Height=None):
        arguments = [Type, Left, Top, Width, Height]
        return self.AddPlaceholder(*arguments)

    def AddPolyline(self, SafeArrayOfPoints=None):
        arguments = com_arguments([unwrap(a) for a in [SafeArrayOfPoints]])
        return Shape(self.com_object.AddPolyline(*arguments))

    # Lower case alias for AddPolyline
    def addpolyline(self, SafeArrayOfPoints=None):
        arguments = [SafeArrayOfPoints]
        return self.AddPolyline(*arguments)

    def AddShape(self, Type=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([unwrap(a) for a in [Type, Left, Top, Width, Height]])
        return Shape(self.com_object.AddShape(*arguments))

    # Lower case alias for AddShape
    def addshape(self, Type=None, Left=None, Top=None, Width=None, Height=None):
        arguments = [Type, Left, Top, Width, Height]
        return self.AddShape(*arguments)

    def AddSmartArt(self, Layout=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([unwrap(a) for a in [Layout, Left, Top, Width, Height]])
        return Shape(self.com_object.AddSmartArt(*arguments))

    # Lower case alias for AddSmartArt
    def addsmartart(self, Layout=None, Left=None, Top=None, Width=None, Height=None):
        arguments = [Layout, Left, Top, Width, Height]
        return self.AddSmartArt(*arguments)

    def AddTable(self, NumRows=None, NumColumns=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([unwrap(a) for a in [NumRows, NumColumns, Left, Top, Width, Height]])
        return Shape(self.com_object.AddTable(*arguments))

    # Lower case alias for AddTable
    def addtable(self, NumRows=None, NumColumns=None, Left=None, Top=None, Width=None, Height=None):
        arguments = [NumRows, NumColumns, Left, Top, Width, Height]
        return self.AddTable(*arguments)

    def AddTextbox(self, Orientation=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([unwrap(a) for a in [Orientation, Left, Top, Width, Height]])
        return Shape(self.com_object.AddTextbox(*arguments))

    # Lower case alias for AddTextbox
    def addtextbox(self, Orientation=None, Left=None, Top=None, Width=None, Height=None):
        arguments = [Orientation, Left, Top, Width, Height]
        return self.AddTextbox(*arguments)

    def AddTextEffect(self, PresetTextEffect=None, Text=None, FontName=None, FontSize=None, FontBold=None, FontItalic=None, Left=None, Top=None):
        arguments = com_arguments([unwrap(a) for a in [PresetTextEffect, Text, FontName, FontSize, FontBold, FontItalic, Left, Top]])
        return Shape(self.com_object.AddTextEffect(*arguments))

    # Lower case alias for AddTextEffect
    def addtexteffect(self, PresetTextEffect=None, Text=None, FontName=None, FontSize=None, FontBold=None, FontItalic=None, Left=None, Top=None):
        arguments = [PresetTextEffect, Text, FontName, FontSize, FontBold, FontItalic, Left, Top]
        return self.AddTextEffect(*arguments)

    def AddTitle(self):
        return Shape(self.com_object.AddTitle())

    # Lower case alias for AddTitle
    def addtitle(self):
        return self.AddTitle()

    def BuildFreeform(self, EditingType=None, X1=None, Y1=None):
        arguments = com_arguments([unwrap(a) for a in [EditingType, X1, Y1]])
        return FreeformBuilder(self.com_object.BuildFreeform(*arguments))

    # Lower case alias for BuildFreeform
    def buildfreeform(self, EditingType=None, X1=None, Y1=None):
        arguments = [EditingType, X1, Y1]
        return self.BuildFreeform(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Shape(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Paste(self):
        return ShapeRange(self.com_object.Paste())

    # Lower case alias for Paste
    def paste(self):
        return self.Paste()

    def PasteSpecial(self, DataType=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Link=None):
        arguments = com_arguments([unwrap(a) for a in [DataType, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Link]])
        return ShapeRange(self.com_object.PasteSpecial(*arguments))

    # Lower case alias for PasteSpecial
    def pastespecial(self, DataType=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Link=None):
        arguments = [DataType, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Link]
        return self.PasteSpecial(*arguments)

    def Range(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return ShapeRange(self.com_object.Range(*arguments))

    # Lower case alias for Range
    def range(self, Index=None):
        arguments = [Index]
        return self.Range(*arguments)

    def SelectAll(self):
        self.com_object.SelectAll()

    # Lower case alias for SelectAll
    def selectall(self):
        return self.SelectAll()


class Slide:

    def __init__(self, slide=None):
        self.com_object= slide

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Background(self):
        return ShapeRange(self.com_object.Background)

    # Lower case aliases for Background
    @property
    def background(self):
        return self.Background

    @property
    def BackgroundStyle(self):
        return self.com_object.BackgroundStyle

    @BackgroundStyle.setter
    def BackgroundStyle(self, value):
        self.com_object.BackgroundStyle = value

    # Lower case aliases for BackgroundStyle
    @property
    def backgroundstyle(self):
        return self.BackgroundStyle

    @backgroundstyle.setter
    def backgroundstyle(self, value):
        self.BackgroundStyle = value

    @property
    def ColorScheme(self):
        return ColorScheme(self.com_object.ColorScheme)

    @ColorScheme.setter
    def ColorScheme(self, value):
        self.com_object.ColorScheme = value

    # Lower case aliases for ColorScheme
    @property
    def colorscheme(self):
        return self.ColorScheme

    @colorscheme.setter
    def colorscheme(self, value):
        self.ColorScheme = value

    @property
    def Comments(self):
        return Comments(self.com_object.Comments)

    # Lower case aliases for Comments
    @property
    def comments(self):
        return self.Comments

    @property
    def CustomerData(self):
        return CustomerData(self.com_object.CustomerData)

    # Lower case aliases for CustomerData
    @property
    def customerdata(self):
        return self.CustomerData

    @property
    def CustomLayout(self):
        return CustomLayout(self.com_object.CustomLayout)

    # Lower case aliases for CustomLayout
    @property
    def customlayout(self):
        return self.CustomLayout

    @property
    def Design(self):
        return Design(self.com_object.Design)

    # Lower case aliases for Design
    @property
    def design(self):
        return self.Design

    @property
    def DisplayMasterShapes(self):
        return self.com_object.DisplayMasterShapes

    @DisplayMasterShapes.setter
    def DisplayMasterShapes(self, value):
        self.com_object.DisplayMasterShapes = value

    # Lower case aliases for DisplayMasterShapes
    @property
    def displaymastershapes(self):
        return self.DisplayMasterShapes

    @displaymastershapes.setter
    def displaymastershapes(self, value):
        self.DisplayMasterShapes = value

    @property
    def FollowMasterBackground(self):
        return self.com_object.FollowMasterBackground

    @FollowMasterBackground.setter
    def FollowMasterBackground(self, value):
        self.com_object.FollowMasterBackground = value

    # Lower case aliases for FollowMasterBackground
    @property
    def followmasterbackground(self):
        return self.FollowMasterBackground

    @followmasterbackground.setter
    def followmasterbackground(self, value):
        self.FollowMasterBackground = value

    @property
    def HasNotesPage(self):
        return self.com_object.HasNotesPage

    # Lower case aliases for HasNotesPage
    @property
    def hasnotespage(self):
        return self.HasNotesPage

    @property
    def HeadersFooters(self):
        return HeadersFooters(self.com_object.HeadersFooters)

    # Lower case aliases for HeadersFooters
    @property
    def headersfooters(self):
        return self.HeadersFooters

    @property
    def Hyperlinks(self):
        return Hyperlinks(self.com_object.Hyperlinks)

    # Lower case aliases for Hyperlinks
    @property
    def hyperlinks(self):
        return self.Hyperlinks

    @property
    def Layout(self):
        return PpSlideLayout(self.com_object.Layout)

    @Layout.setter
    def Layout(self, value):
        self.com_object.Layout = value

    # Lower case aliases for Layout
    @property
    def layout(self):
        return self.Layout

    @layout.setter
    def layout(self, value):
        self.Layout = value

    @property
    def Master(self):
        return Master(self.com_object.Master)

    # Lower case aliases for Master
    @property
    def master(self):
        return self.Master

    @property
    def Name(self):
        return self.com_object.Name

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def NotesPage(self):
        return SlideRange(self.com_object.NotesPage)

    # Lower case aliases for NotesPage
    @property
    def notespage(self):
        return self.NotesPage

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PrintSteps(self):
        return self.com_object.PrintSteps

    # Lower case aliases for PrintSteps
    @property
    def printsteps(self):
        return self.PrintSteps

    @property
    def sectionIndex(self):
        return Slide(self.com_object.sectionIndex)

    # Lower case aliases for sectionIndex
    @property
    def sectionindex(self):
        return self.sectionIndex

    @property
    def Shapes(self):
        return Shapes(self.com_object.Shapes)

    # Lower case aliases for Shapes
    @property
    def shapes(self):
        return self.Shapes

    @property
    def SlideID(self):
        return self.com_object.SlideID

    # Lower case aliases for SlideID
    @property
    def slideid(self):
        return self.SlideID

    @property
    def SlideIndex(self):
        return Slides(self.com_object.SlideIndex)

    # Lower case aliases for SlideIndex
    @property
    def slideindex(self):
        return self.SlideIndex

    @property
    def SlideNumber(self):
        return self.com_object.SlideNumber

    # Lower case aliases for SlideNumber
    @property
    def slidenumber(self):
        return self.SlideNumber

    @property
    def SlideShowTransition(self):
        return SlideShowTransition(self.com_object.SlideShowTransition)

    # Lower case aliases for SlideShowTransition
    @property
    def slideshowtransition(self):
        return self.SlideShowTransition

    @property
    def Tags(self):
        return Tags(self.com_object.Tags)

    # Lower case aliases for Tags
    @property
    def tags(self):
        return self.Tags

    @property
    def ThemeColorScheme(self):
        return self.com_object.ThemeColorScheme

    # Lower case aliases for ThemeColorScheme
    @property
    def themecolorscheme(self):
        return self.ThemeColorScheme

    @property
    def TimeLine(self):
        return TimeLine(self.com_object.TimeLine)

    # Lower case aliases for TimeLine
    @property
    def timeline(self):
        return self.TimeLine

    def ApplyTemplate(self, FileName=None):
        arguments = com_arguments([unwrap(a) for a in [FileName]])
        self.com_object.ApplyTemplate(*arguments)

    # Lower case alias for ApplyTemplate
    def applytemplate(self, FileName=None):
        arguments = [FileName]
        return self.ApplyTemplate(*arguments)

    def ApplyTheme(self, themeName=None):
        arguments = com_arguments([unwrap(a) for a in [themeName]])
        self.com_object.ApplyTheme(*arguments)

    # Lower case alias for ApplyTheme
    def applytheme(self, themeName=None):
        arguments = [themeName]
        return self.ApplyTheme(*arguments)

    def ApplyThemeColorScheme(self, themeColorSchemeName=None):
        arguments = com_arguments([unwrap(a) for a in [themeColorSchemeName]])
        self.com_object.ApplyThemeColorScheme(*arguments)

    # Lower case alias for ApplyThemeColorScheme
    def applythemecolorscheme(self, themeColorSchemeName=None):
        arguments = [themeColorSchemeName]
        return self.ApplyThemeColorScheme(*arguments)

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

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Duplicate(self):
        return SlideRange(self.com_object.Duplicate())

    # Lower case alias for Duplicate
    def duplicate(self):
        return self.Duplicate()

    def Export(self, FileName=None, FilterName=None, ScaleWidth=None, ScaleHeight=None):
        arguments = com_arguments([unwrap(a) for a in [FileName, FilterName, ScaleWidth, ScaleHeight]])
        self.com_object.Export(*arguments)

    # Lower case alias for Export
    def export(self, FileName=None, FilterName=None, ScaleWidth=None, ScaleHeight=None):
        arguments = [FileName, FilterName, ScaleWidth, ScaleHeight]
        return self.Export(*arguments)

    def MoveTo(self, toPos=None):
        arguments = com_arguments([unwrap(a) for a in [toPos]])
        self.com_object.MoveTo(*arguments)

    # Lower case alias for MoveTo
    def moveto(self, toPos=None):
        arguments = [toPos]
        return self.MoveTo(*arguments)

    def MoveToSectionStart(self, toSection=None):
        arguments = com_arguments([unwrap(a) for a in [toSection]])
        self.com_object.MoveToSectionStart(*arguments)

    # Lower case alias for MoveToSectionStart
    def movetosectionstart(self, toSection=None):
        arguments = [toSection]
        return self.MoveToSectionStart(*arguments)

    def PublishSlides(self, SlideLibraryUrl=None, Overwrite=None, UseSlideOrder=None):
        arguments = com_arguments([unwrap(a) for a in [SlideLibraryUrl, Overwrite, UseSlideOrder]])
        return self.com_object.PublishSlides(*arguments)

    # Lower case alias for PublishSlides
    def publishslides(self, SlideLibraryUrl=None, Overwrite=None, UseSlideOrder=None):
        arguments = [SlideLibraryUrl, Overwrite, UseSlideOrder]
        return self.PublishSlides(*arguments)

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class SlideRange:

    def __init__(self, sliderange=None):
        self.com_object= sliderange

    def __call__(self, item):
        return SlideRange(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Background(self):
        return ShapeRange(self.com_object.Background)

    # Lower case aliases for Background
    @property
    def background(self):
        return self.Background

    @property
    def BackgroundStyle(self):
        return self.com_object.BackgroundStyle

    @BackgroundStyle.setter
    def BackgroundStyle(self, value):
        self.com_object.BackgroundStyle = value

    # Lower case aliases for BackgroundStyle
    @property
    def backgroundstyle(self):
        return self.BackgroundStyle

    @backgroundstyle.setter
    def backgroundstyle(self, value):
        self.BackgroundStyle = value

    @property
    def ColorScheme(self):
        return ColorScheme(self.com_object.ColorScheme)

    @ColorScheme.setter
    def ColorScheme(self, value):
        self.com_object.ColorScheme = value

    # Lower case aliases for ColorScheme
    @property
    def colorscheme(self):
        return self.ColorScheme

    @colorscheme.setter
    def colorscheme(self, value):
        self.ColorScheme = value

    @property
    def Comments(self):
        return Comments(self.com_object.Comments)

    # Lower case aliases for Comments
    @property
    def comments(self):
        return self.Comments

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def CustomerData(self):
        return CustomerData(self.com_object.CustomerData)

    # Lower case aliases for CustomerData
    @property
    def customerdata(self):
        return self.CustomerData

    @property
    def CustomLayout(self):
        return CustomLayout(self.com_object.CustomLayout)

    # Lower case aliases for CustomLayout
    @property
    def customlayout(self):
        return self.CustomLayout

    @property
    def Design(self):
        return Design(self.com_object.Design)

    # Lower case aliases for Design
    @property
    def design(self):
        return self.Design

    @property
    def DisplayMasterShapes(self):
        return self.com_object.DisplayMasterShapes

    @DisplayMasterShapes.setter
    def DisplayMasterShapes(self, value):
        self.com_object.DisplayMasterShapes = value

    # Lower case aliases for DisplayMasterShapes
    @property
    def displaymastershapes(self):
        return self.DisplayMasterShapes

    @displaymastershapes.setter
    def displaymastershapes(self, value):
        self.DisplayMasterShapes = value

    @property
    def FollowMasterBackground(self):
        return self.com_object.FollowMasterBackground

    @FollowMasterBackground.setter
    def FollowMasterBackground(self, value):
        self.com_object.FollowMasterBackground = value

    # Lower case aliases for FollowMasterBackground
    @property
    def followmasterbackground(self):
        return self.FollowMasterBackground

    @followmasterbackground.setter
    def followmasterbackground(self, value):
        self.FollowMasterBackground = value

    @property
    def HasNotesPage(self):
        return self.com_object.HasNotesPage

    # Lower case aliases for HasNotesPage
    @property
    def hasnotespage(self):
        return self.HasNotesPage

    @property
    def HeadersFooters(self):
        return HeadersFooters(self.com_object.HeadersFooters)

    # Lower case aliases for HeadersFooters
    @property
    def headersfooters(self):
        return self.HeadersFooters

    @property
    def Hyperlinks(self):
        return Hyperlinks(self.com_object.Hyperlinks)

    # Lower case aliases for Hyperlinks
    @property
    def hyperlinks(self):
        return self.Hyperlinks

    @property
    def Layout(self):
        return PpSlideLayout(self.com_object.Layout)

    @Layout.setter
    def Layout(self, value):
        self.com_object.Layout = value

    # Lower case aliases for Layout
    @property
    def layout(self):
        return self.Layout

    @layout.setter
    def layout(self, value):
        self.Layout = value

    @property
    def Master(self):
        return Master(self.com_object.Master)

    # Lower case aliases for Master
    @property
    def master(self):
        return self.Master

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
    def NotesPage(self):
        return SlideRange(self.com_object.NotesPage)

    # Lower case aliases for NotesPage
    @property
    def notespage(self):
        return self.NotesPage

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PrintSteps(self):
        return self.com_object.PrintSteps

    # Lower case aliases for PrintSteps
    @property
    def printsteps(self):
        return self.PrintSteps

    @property
    def sectionIndex(self):
        return SlideRange(self.com_object.sectionIndex)

    # Lower case aliases for sectionIndex
    @property
    def sectionindex(self):
        return self.sectionIndex

    @property
    def Shapes(self):
        return Shapes(self.com_object.Shapes)

    # Lower case aliases for Shapes
    @property
    def shapes(self):
        return self.Shapes

    @property
    def SlideID(self):
        return self.com_object.SlideID

    # Lower case aliases for SlideID
    @property
    def slideid(self):
        return self.SlideID

    @property
    def SlideIndex(self):
        return Slides(self.com_object.SlideIndex)

    # Lower case aliases for SlideIndex
    @property
    def slideindex(self):
        return self.SlideIndex

    @property
    def SlideNumber(self):
        return self.com_object.SlideNumber

    # Lower case aliases for SlideNumber
    @property
    def slidenumber(self):
        return self.SlideNumber

    @property
    def SlideShowTransition(self):
        return SlideShowTransition(self.com_object.SlideShowTransition)

    # Lower case aliases for SlideShowTransition
    @property
    def slideshowtransition(self):
        return self.SlideShowTransition

    @property
    def Tags(self):
        return Tags(self.com_object.Tags)

    # Lower case aliases for Tags
    @property
    def tags(self):
        return self.Tags

    @property
    def ThemeColorScheme(self):
        return self.com_object.ThemeColorScheme

    # Lower case aliases for ThemeColorScheme
    @property
    def themecolorscheme(self):
        return self.ThemeColorScheme

    @property
    def TimeLine(self):
        return TimeLine(self.com_object.TimeLine)

    # Lower case aliases for TimeLine
    @property
    def timeline(self):
        return self.TimeLine

    def ApplyTemplate(self, FileName=None):
        arguments = com_arguments([unwrap(a) for a in [FileName]])
        self.com_object.ApplyTemplate(*arguments)

    # Lower case alias for ApplyTemplate
    def applytemplate(self, FileName=None):
        arguments = [FileName]
        return self.ApplyTemplate(*arguments)

    def ApplyTheme(self, themeName=None):
        arguments = com_arguments([unwrap(a) for a in [themeName]])
        self.com_object.ApplyTheme(*arguments)

    # Lower case alias for ApplyTheme
    def applytheme(self, themeName=None):
        arguments = [themeName]
        return self.ApplyTheme(*arguments)

    def ApplyThemeColorScheme(self, themeColorSchemeName=None):
        arguments = com_arguments([unwrap(a) for a in [themeColorSchemeName]])
        self.com_object.ApplyThemeColorScheme(*arguments)

    # Lower case alias for ApplyThemeColorScheme
    def applythemecolorscheme(self, themeColorSchemeName=None):
        arguments = [themeColorSchemeName]
        return self.ApplyThemeColorScheme(*arguments)

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

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Duplicate(self):
        return SlideRange(self.com_object.Duplicate())

    # Lower case alias for Duplicate
    def duplicate(self):
        return self.Duplicate()

    def Export(self, FileName=None, FilterName=None, ScaleWidth=None, ScaleHeight=None):
        arguments = com_arguments([unwrap(a) for a in [FileName, FilterName, ScaleWidth, ScaleHeight]])
        self.com_object.Export(*arguments)

    # Lower case alias for Export
    def export(self, FileName=None, FilterName=None, ScaleWidth=None, ScaleHeight=None):
        arguments = [FileName, FilterName, ScaleWidth, ScaleHeight]
        return self.Export(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Slide(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def MoveTo(self, toPos=None):
        arguments = com_arguments([unwrap(a) for a in [toPos]])
        self.com_object.MoveTo(*arguments)

    # Lower case alias for MoveTo
    def moveto(self, toPos=None):
        arguments = [toPos]
        return self.MoveTo(*arguments)

    def MoveToSectionStart(self, toSection=None):
        arguments = com_arguments([unwrap(a) for a in [toSection]])
        self.com_object.MoveToSectionStart(*arguments)

    # Lower case alias for MoveToSectionStart
    def movetosectionstart(self, toSection=None):
        arguments = [toSection]
        return self.MoveToSectionStart(*arguments)

    def PublishSlides(self, SlideLibraryUrl=None, Overwrite=None):
        arguments = com_arguments([unwrap(a) for a in [SlideLibraryUrl, Overwrite]])
        self.com_object.PublishSlides(*arguments)

    # Lower case alias for PublishSlides
    def publishslides(self, SlideLibraryUrl=None, Overwrite=None):
        arguments = [SlideLibraryUrl, Overwrite]
        return self.PublishSlides(*arguments)

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class Slides:

    def __init__(self, slides=None):
        self.com_object= slides

    def __call__(self, item):
        return Slide(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def AddSlide(self, Index=None, pCustomLayout=None):
        arguments = com_arguments([unwrap(a) for a in [Index, pCustomLayout]])
        return Slide(self.com_object.AddSlide(*arguments))

    # Lower case alias for AddSlide
    def addslide(self, Index=None, pCustomLayout=None):
        arguments = [Index, pCustomLayout]
        return self.AddSlide(*arguments)

    def FindBySlideID(self, SlideID=None):
        arguments = com_arguments([unwrap(a) for a in [SlideID]])
        return Slide(self.com_object.FindBySlideID(*arguments))

    # Lower case alias for FindBySlideID
    def findbyslideid(self, SlideID=None):
        arguments = [SlideID]
        return self.FindBySlideID(*arguments)

    def InsertFromFile(self, FileName=None, Index=None, SlideStart=None, SlideEnd=None):
        arguments = com_arguments([unwrap(a) for a in [FileName, Index, SlideStart, SlideEnd]])
        return self.com_object.InsertFromFile(*arguments)

    # Lower case alias for InsertFromFile
    def insertfromfile(self, FileName=None, Index=None, SlideStart=None, SlideEnd=None):
        arguments = [FileName, Index, SlideStart, SlideEnd]
        return self.InsertFromFile(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Slide(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Paste(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return SlideRange(self.com_object.Paste(*arguments))

    # Lower case alias for Paste
    def paste(self, Index=None):
        arguments = [Index]
        return self.Paste(*arguments)

    def Range(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return SlideRange(self.com_object.Range(*arguments))

    # Lower case alias for Range
    def range(self, Index=None):
        arguments = [Index]
        return self.Range(*arguments)


class SlideShowSettings:

    def __init__(self, slideshowsettings=None):
        self.com_object= slideshowsettings

    @property
    def AdvanceMode(self):
        return self.com_object.AdvanceMode

    @AdvanceMode.setter
    def AdvanceMode(self, value):
        self.com_object.AdvanceMode = value

    # Lower case aliases for AdvanceMode
    @property
    def advancemode(self):
        return self.AdvanceMode

    @advancemode.setter
    def advancemode(self, value):
        self.AdvanceMode = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def EndingSlide(self):
        return self.com_object.EndingSlide

    @EndingSlide.setter
    def EndingSlide(self, value):
        self.com_object.EndingSlide = value

    # Lower case aliases for EndingSlide
    @property
    def endingslide(self):
        return self.EndingSlide

    @endingslide.setter
    def endingslide(self, value):
        self.EndingSlide = value

    @property
    def LoopUntilStopped(self):
        return self.com_object.LoopUntilStopped

    @LoopUntilStopped.setter
    def LoopUntilStopped(self, value):
        self.com_object.LoopUntilStopped = value

    # Lower case aliases for LoopUntilStopped
    @property
    def loopuntilstopped(self):
        return self.LoopUntilStopped

    @loopuntilstopped.setter
    def loopuntilstopped(self, value):
        self.LoopUntilStopped = value

    @property
    def NamedSlideShows(self):
        return NamedSlideShows(self.com_object.NamedSlideShows)

    # Lower case aliases for NamedSlideShows
    @property
    def namedslideshows(self):
        return self.NamedSlideShows

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PointerColor(self):
        return ColorFormat(self.com_object.PointerColor)

    # Lower case aliases for PointerColor
    @property
    def pointercolor(self):
        return self.PointerColor

    @property
    def RangeType(self):
        return self.com_object.RangeType

    @RangeType.setter
    def RangeType(self, value):
        self.com_object.RangeType = value

    # Lower case aliases for RangeType
    @property
    def rangetype(self):
        return self.RangeType

    @rangetype.setter
    def rangetype(self, value):
        self.RangeType = value

    @property
    def ShowMediaControls(self):
        return self.com_object.ShowMediaControls

    @ShowMediaControls.setter
    def ShowMediaControls(self, value):
        self.com_object.ShowMediaControls = value

    # Lower case aliases for ShowMediaControls
    @property
    def showmediacontrols(self):
        return self.ShowMediaControls

    @showmediacontrols.setter
    def showmediacontrols(self, value):
        self.ShowMediaControls = value

    @property
    def ShowPresenterView(self):
        return SlideShowSettings(self.com_object.ShowPresenterView)

    @ShowPresenterView.setter
    def ShowPresenterView(self, value):
        self.com_object.ShowPresenterView = value

    # Lower case aliases for ShowPresenterView
    @property
    def showpresenterview(self):
        return self.ShowPresenterView

    @showpresenterview.setter
    def showpresenterview(self, value):
        self.ShowPresenterView = value

    @property
    def ShowScrollbar(self):
        return self.com_object.ShowScrollbar

    @ShowScrollbar.setter
    def ShowScrollbar(self, value):
        self.com_object.ShowScrollbar = value

    # Lower case aliases for ShowScrollbar
    @property
    def showscrollbar(self):
        return self.ShowScrollbar

    @showscrollbar.setter
    def showscrollbar(self, value):
        self.ShowScrollbar = value

    @property
    def ShowType(self):
        return self.com_object.ShowType

    @ShowType.setter
    def ShowType(self, value):
        self.com_object.ShowType = value

    # Lower case aliases for ShowType
    @property
    def showtype(self):
        return self.ShowType

    @showtype.setter
    def showtype(self, value):
        self.ShowType = value

    @property
    def ShowWithAnimation(self):
        return self.com_object.ShowWithAnimation

    @ShowWithAnimation.setter
    def ShowWithAnimation(self, value):
        self.com_object.ShowWithAnimation = value

    # Lower case aliases for ShowWithAnimation
    @property
    def showwithanimation(self):
        return self.ShowWithAnimation

    @showwithanimation.setter
    def showwithanimation(self, value):
        self.ShowWithAnimation = value

    @property
    def ShowWithNarration(self):
        return self.com_object.ShowWithNarration

    @ShowWithNarration.setter
    def ShowWithNarration(self, value):
        self.com_object.ShowWithNarration = value

    # Lower case aliases for ShowWithNarration
    @property
    def showwithnarration(self):
        return self.ShowWithNarration

    @showwithnarration.setter
    def showwithnarration(self, value):
        self.ShowWithNarration = value

    @property
    def SlideShowName(self):
        return self.com_object.SlideShowName

    @SlideShowName.setter
    def SlideShowName(self, value):
        self.com_object.SlideShowName = value

    # Lower case aliases for SlideShowName
    @property
    def slideshowname(self):
        return self.SlideShowName

    @slideshowname.setter
    def slideshowname(self, value):
        self.SlideShowName = value

    @property
    def StartingSlide(self):
        return self.com_object.StartingSlide

    @StartingSlide.setter
    def StartingSlide(self, value):
        self.com_object.StartingSlide = value

    # Lower case aliases for StartingSlide
    @property
    def startingslide(self):
        return self.StartingSlide

    @startingslide.setter
    def startingslide(self, value):
        self.StartingSlide = value

    def Run(self):
        return SlideShowWindow(self.com_object.Run())

    # Lower case alias for Run
    def run(self):
        return self.Run()


class SlideShowTransition:

    def __init__(self, slideshowtransition=None):
        self.com_object= slideshowtransition

    @property
    def AdvanceOnClick(self):
        return self.com_object.AdvanceOnClick

    @AdvanceOnClick.setter
    def AdvanceOnClick(self, value):
        self.com_object.AdvanceOnClick = value

    # Lower case aliases for AdvanceOnClick
    @property
    def advanceonclick(self):
        return self.AdvanceOnClick

    @advanceonclick.setter
    def advanceonclick(self, value):
        self.AdvanceOnClick = value

    @property
    def AdvanceOnTime(self):
        return self.com_object.AdvanceOnTime

    @AdvanceOnTime.setter
    def AdvanceOnTime(self, value):
        self.com_object.AdvanceOnTime = value

    # Lower case aliases for AdvanceOnTime
    @property
    def advanceontime(self):
        return self.AdvanceOnTime

    @advanceontime.setter
    def advanceontime(self, value):
        self.AdvanceOnTime = value

    @property
    def AdvanceTime(self):
        return self.com_object.AdvanceTime

    @AdvanceTime.setter
    def AdvanceTime(self, value):
        self.com_object.AdvanceTime = value

    # Lower case aliases for AdvanceTime
    @property
    def advancetime(self):
        return self.AdvanceTime

    @advancetime.setter
    def advancetime(self, value):
        self.AdvanceTime = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Duration(self):
        return self.com_object.Duration

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
    def EntryEffect(self):
        return self.com_object.EntryEffect

    @EntryEffect.setter
    def EntryEffect(self, value):
        self.com_object.EntryEffect = value

    # Lower case aliases for EntryEffect
    @property
    def entryeffect(self):
        return self.EntryEffect

    @entryeffect.setter
    def entryeffect(self, value):
        self.EntryEffect = value

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
    def LoopSoundUntilNext(self):
        return self.com_object.LoopSoundUntilNext

    @LoopSoundUntilNext.setter
    def LoopSoundUntilNext(self, value):
        self.com_object.LoopSoundUntilNext = value

    # Lower case aliases for LoopSoundUntilNext
    @property
    def loopsounduntilnext(self):
        return self.LoopSoundUntilNext

    @loopsounduntilnext.setter
    def loopsounduntilnext(self, value):
        self.LoopSoundUntilNext = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SoundEffect(self):
        return SoundEffect(self.com_object.SoundEffect)

    # Lower case aliases for SoundEffect
    @property
    def soundeffect(self):
        return self.SoundEffect

    @property
    def Speed(self):
        return self.com_object.Speed

    @Speed.setter
    def Speed(self, value):
        self.com_object.Speed = value

    # Lower case aliases for Speed
    @property
    def speed(self):
        return self.Speed

    @speed.setter
    def speed(self, value):
        self.Speed = value


class SlideShowView:

    def __init__(self, slideshowview=None):
        self.com_object= slideshowview

    @property
    def AcceleratorsEnabled(self):
        return self.com_object.AcceleratorsEnabled

    @AcceleratorsEnabled.setter
    def AcceleratorsEnabled(self, value):
        self.com_object.AcceleratorsEnabled = value

    # Lower case aliases for AcceleratorsEnabled
    @property
    def acceleratorsenabled(self):
        return self.AcceleratorsEnabled

    @acceleratorsenabled.setter
    def acceleratorsenabled(self, value):
        self.AcceleratorsEnabled = value

    @property
    def AdvanceMode(self):
        return self.com_object.AdvanceMode

    # Lower case aliases for AdvanceMode
    @property
    def advancemode(self):
        return self.AdvanceMode

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def CurrentShowPosition(self):
        return self.com_object.CurrentShowPosition

    # Lower case aliases for CurrentShowPosition
    @property
    def currentshowposition(self):
        return self.CurrentShowPosition

    @property
    def IsNamedShow(self):
        return self.com_object.IsNamedShow

    # Lower case aliases for IsNamedShow
    @property
    def isnamedshow(self):
        return self.IsNamedShow

    @property
    def LastSlideViewed(self):
        return Slide(self.com_object.LastSlideViewed)

    # Lower case aliases for LastSlideViewed
    @property
    def lastslideviewed(self):
        return self.LastSlideViewed

    @property
    def MediaControlsHeight(self):
        return self.com_object.MediaControlsHeight

    # Lower case aliases for MediaControlsHeight
    @property
    def mediacontrolsheight(self):
        return self.MediaControlsHeight

    @property
    def MediaControlsLeft(self):
        return Slide(self.com_object.MediaControlsLeft)

    # Lower case aliases for MediaControlsLeft
    @property
    def mediacontrolsleft(self):
        return self.MediaControlsLeft

    @property
    def MediaControlsTop(self):
        return Slide(self.com_object.MediaControlsTop)

    # Lower case aliases for MediaControlsTop
    @property
    def mediacontrolstop(self):
        return self.MediaControlsTop

    @property
    def MediaControlsVisible(self):
        return self.com_object.MediaControlsVisible

    # Lower case aliases for MediaControlsVisible
    @property
    def mediacontrolsvisible(self):
        return self.MediaControlsVisible

    @property
    def MediaControlsWidth(self):
        return self.com_object.MediaControlsWidth

    # Lower case aliases for MediaControlsWidth
    @property
    def mediacontrolswidth(self):
        return self.MediaControlsWidth

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PointerColor(self):
        return ColorFormat(self.com_object.PointerColor)

    # Lower case aliases for PointerColor
    @property
    def pointercolor(self):
        return self.PointerColor

    @property
    def PointerType(self):
        return self.com_object.PointerType

    @PointerType.setter
    def PointerType(self, value):
        self.com_object.PointerType = value

    # Lower case aliases for PointerType
    @property
    def pointertype(self):
        return self.PointerType

    @pointertype.setter
    def pointertype(self, value):
        self.PointerType = value

    @property
    def PresentationElapsedTime(self):
        return self.com_object.PresentationElapsedTime

    # Lower case aliases for PresentationElapsedTime
    @property
    def presentationelapsedtime(self):
        return self.PresentationElapsedTime

    @property
    def Slide(self):
        return Slide(self.com_object.Slide)

    # Lower case aliases for Slide
    @property
    def slide(self):
        return self.Slide

    @property
    def SlideElapsedTime(self):
        return self.com_object.SlideElapsedTime

    @SlideElapsedTime.setter
    def SlideElapsedTime(self, value):
        self.com_object.SlideElapsedTime = value

    # Lower case aliases for SlideElapsedTime
    @property
    def slideelapsedtime(self):
        return self.SlideElapsedTime

    @slideelapsedtime.setter
    def slideelapsedtime(self, value):
        self.SlideElapsedTime = value

    @property
    def SlideShowName(self):
        return self.com_object.SlideShowName

    # Lower case aliases for SlideShowName
    @property
    def slideshowname(self):
        return self.SlideShowName

    @property
    def State(self):
        return self.com_object.State

    @State.setter
    def State(self, value):
        self.com_object.State = value

    # Lower case aliases for State
    @property
    def state(self):
        return self.State

    @state.setter
    def state(self, value):
        self.State = value

    @property
    def Zoom(self):
        return self.com_object.Zoom

    # Lower case aliases for Zoom
    @property
    def zoom(self):
        return self.Zoom

    def DrawLine(self, BeginX=None, BeginY=None, EndX=None, EndY=None):
        arguments = com_arguments([unwrap(a) for a in [BeginX, BeginY, EndX, EndY]])
        self.com_object.DrawLine(*arguments)

    # Lower case alias for DrawLine
    def drawline(self, BeginX=None, BeginY=None, EndX=None, EndY=None):
        arguments = [BeginX, BeginY, EndX, EndY]
        return self.DrawLine(*arguments)

    def EndNamedShow(self):
        self.com_object.EndNamedShow()

    # Lower case alias for EndNamedShow
    def endnamedshow(self):
        return self.EndNamedShow()

    def EraseDrawing(self):
        self.com_object.EraseDrawing()

    # Lower case alias for EraseDrawing
    def erasedrawing(self):
        return self.EraseDrawing()

    def Exit(self):
        self.com_object.Exit()

    # Lower case alias for Exit
    def exit(self):
        return self.Exit()

    def First(self):
        return self.com_object.First()

    # Lower case alias for First
    def first(self):
        return self.First()

    def FirstAnimationIsAutomatic(self):
        return self.com_object.FirstAnimationIsAutomatic()

    # Lower case alias for FirstAnimationIsAutomatic
    def firstanimationisautomatic(self):
        return self.FirstAnimationIsAutomatic()

    def GetClickCount(self):
        return self.com_object.GetClickCount()

    # Lower case alias for GetClickCount
    def getclickcount(self):
        return self.GetClickCount()

    def GetClickIndex(self):
        return self.com_object.GetClickIndex()

    # Lower case alias for GetClickIndex
    def getclickindex(self):
        return self.GetClickIndex()

    def GotoClick(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        self.com_object.GotoClick(*arguments)

    # Lower case alias for GotoClick
    def gotoclick(self, Index=None):
        arguments = [Index]
        return self.GotoClick(*arguments)

    def GotoNamedShow(self, SlideShowName=None):
        arguments = com_arguments([unwrap(a) for a in [SlideShowName]])
        self.com_object.GotoNamedShow(*arguments)

    # Lower case alias for GotoNamedShow
    def gotonamedshow(self, SlideShowName=None):
        arguments = [SlideShowName]
        return self.GotoNamedShow(*arguments)

    def GotoSlide(self, Index=None, ResetSlide=None):
        arguments = com_arguments([unwrap(a) for a in [Index, ResetSlide]])
        self.com_object.GotoSlide(*arguments)

    # Lower case alias for GotoSlide
    def gotoslide(self, Index=None, ResetSlide=None):
        arguments = [Index, ResetSlide]
        return self.GotoSlide(*arguments)

    def Last(self):
        self.com_object.Last()

    # Lower case alias for Last
    def last(self):
        return self.Last()

    def Next(self):
        self.com_object.Next()

    # Lower case alias for Next
    def next(self):
        return self.Next()

    def Player(self, ShapeId=None):
        arguments = com_arguments([unwrap(a) for a in [ShapeId]])
        return Player(self.com_object.Player(*arguments))

    # Lower case alias for Player
    def player(self, ShapeId=None):
        arguments = [ShapeId]
        return self.Player(*arguments)

    def Previous(self):
        self.com_object.Previous()

    # Lower case alias for Previous
    def previous(self):
        return self.Previous()

    def ResetSlideTime(self):
        self.com_object.ResetSlideTime()

    # Lower case alias for ResetSlideTime
    def resetslidetime(self):
        return self.ResetSlideTime()


class SlideShowWindow:

    def __init__(self, slideshowwindow=None):
        self.com_object= slideshowwindow

    @property
    def Active(self):
        return self.com_object.Active

    # Lower case aliases for Active
    @property
    def active(self):
        return self.Active

    @property
    def Application(self):
        return Application(self.com_object.Application)

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
    def IsFullScreen(self):
        return self.com_object.IsFullScreen

    # Lower case aliases for IsFullScreen
    @property
    def isfullscreen(self):
        return self.IsFullScreen

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
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Presentation(self):
        return Presentation(self.com_object.Presentation)

    # Lower case aliases for Presentation
    @property
    def presentation(self):
        return self.Presentation

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
    def View(self):
        return SlideShowView(self.com_object.View)

    # Lower case aliases for View
    @property
    def view(self):
        return self.View

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

    def Activate(self):
        self.com_object.Activate()

    # Lower case alias for Activate
    def activate(self):
        return self.Activate()


class SlideShowWindows:

    def __init__(self, slideshowwindows=None):
        self.com_object= slideshowwindows

    def __call__(self, item):
        return SlideShowWindow(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return SlideShowWindow(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class SoundEffect:

    def __init__(self, soundeffect=None):
        self.com_object= soundeffect

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def ImportFromFile(self, FullName=None):
        arguments = com_arguments([unwrap(a) for a in [FullName]])
        self.com_object.ImportFromFile(*arguments)

    # Lower case alias for ImportFromFile
    def importfromfile(self, FullName=None):
        arguments = [FullName]
        return self.ImportFromFile(*arguments)

    def Play(self):
        self.com_object.Play()

    # Lower case alias for Play
    def play(self):
        return self.Play()


class Table:

    def __init__(self, table=None):
        self.com_object= table

    @property
    def AlternativeText(self):
        return self.com_object.AlternativeText

    @AlternativeText.setter
    def AlternativeText(self, value):
        self.com_object.AlternativeText = value

    # Lower case aliases for AlternativeText
    @property
    def alternativetext(self):
        return self.AlternativeText

    @alternativetext.setter
    def alternativetext(self, value):
        self.AlternativeText = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Background(self):
        return TableBackground(self.com_object.Background)

    # Lower case aliases for Background
    @property
    def background(self):
        return self.Background

    @property
    def Columns(self):
        return Columns(self.com_object.Columns)

    # Lower case aliases for Columns
    @property
    def columns(self):
        return self.Columns

    @property
    def FirstCol(self):
        return self.com_object.FirstCol

    @FirstCol.setter
    def FirstCol(self, value):
        self.com_object.FirstCol = value

    # Lower case aliases for FirstCol
    @property
    def firstcol(self):
        return self.FirstCol

    @firstcol.setter
    def firstcol(self, value):
        self.FirstCol = value

    @property
    def FirstRow(self):
        return self.com_object.FirstRow

    @FirstRow.setter
    def FirstRow(self, value):
        self.com_object.FirstRow = value

    # Lower case aliases for FirstRow
    @property
    def firstrow(self):
        return self.FirstRow

    @firstrow.setter
    def firstrow(self, value):
        self.FirstRow = value

    @property
    def HorizBanding(self):
        return self.com_object.HorizBanding

    @HorizBanding.setter
    def HorizBanding(self, value):
        self.com_object.HorizBanding = value

    # Lower case aliases for HorizBanding
    @property
    def horizbanding(self):
        return self.HorizBanding

    @horizbanding.setter
    def horizbanding(self, value):
        self.HorizBanding = value

    @property
    def LastCol(self):
        return self.com_object.LastCol

    @LastCol.setter
    def LastCol(self, value):
        self.com_object.LastCol = value

    # Lower case aliases for LastCol
    @property
    def lastcol(self):
        return self.LastCol

    @lastcol.setter
    def lastcol(self, value):
        self.LastCol = value

    @property
    def LastRow(self):
        return self.com_object.LastRow

    @LastRow.setter
    def LastRow(self, value):
        self.com_object.LastRow = value

    # Lower case aliases for LastRow
    @property
    def lastrow(self):
        return self.LastRow

    @lastrow.setter
    def lastrow(self, value):
        self.LastRow = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Rows(self):
        return Rows(self.com_object.Rows)

    # Lower case aliases for Rows
    @property
    def rows(self):
        return self.Rows

    @property
    def Style(self):
        return TableStyle(self.com_object.Style)

    # Lower case aliases for Style
    @property
    def style(self):
        return self.Style

    @property
    def TableDirection(self):
        return self.com_object.TableDirection

    @TableDirection.setter
    def TableDirection(self, value):
        self.com_object.TableDirection = value

    # Lower case aliases for TableDirection
    @property
    def tabledirection(self):
        return self.TableDirection

    @tabledirection.setter
    def tabledirection(self, value):
        self.TableDirection = value

    @property
    def Title(self):
        return Table(self.com_object.Title)

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
    def VertBanding(self):
        return self.com_object.VertBanding

    @VertBanding.setter
    def VertBanding(self, value):
        self.com_object.VertBanding = value

    # Lower case aliases for VertBanding
    @property
    def vertbanding(self):
        return self.VertBanding

    @vertbanding.setter
    def vertbanding(self, value):
        self.VertBanding = value

    def ApplyStyle(self, StyleID=None, SaveFormatting=None):
        arguments = com_arguments([unwrap(a) for a in [StyleID, SaveFormatting]])
        self.com_object.ApplyStyle(*arguments)

    # Lower case alias for ApplyStyle
    def applystyle(self, StyleID=None, SaveFormatting=None):
        arguments = [StyleID, SaveFormatting]
        return self.ApplyStyle(*arguments)

    def Cell(self, Row=None, Column=None):
        arguments = com_arguments([unwrap(a) for a in [Row, Column]])
        return Cell(self.com_object.Cell(*arguments))

    # Lower case alias for Cell
    def cell(self, Row=None, Column=None):
        arguments = [Row, Column]
        return self.Cell(*arguments)

    def ScaleProportionally(self, scale=None):
        arguments = com_arguments([unwrap(a) for a in [scale]])
        self.com_object.ScaleProportionally(*arguments)

    # Lower case alias for ScaleProportionally
    def scaleproportionally(self, scale=None):
        arguments = [scale]
        return self.ScaleProportionally(*arguments)


class TableBackground:

    def __init__(self, tablebackground=None):
        self.com_object= tablebackground

    @property
    def Fill(self):
        return FillFormat(self.com_object.Fill)

    # Lower case aliases for Fill
    @property
    def fill(self):
        return self.Fill

    @property
    def Picture(self):
        return PictureFormat(self.com_object.Picture)

    # Lower case aliases for Picture
    @property
    def picture(self):
        return self.Picture

    @property
    def Reflection(self):
        return self.com_object.Reflection

    # Lower case aliases for Reflection
    @property
    def reflection(self):
        return self.Reflection

    @property
    def Shadow(self):
        return ShadowFormat(self.com_object.Shadow)

    # Lower case aliases for Shadow
    @property
    def shadow(self):
        return self.Shadow


class TableStyle:

    def __init__(self, tablestyle=None):
        self.com_object= tablestyle

    @property
    def Id(self):
        return self.com_object.Id

    # Lower case aliases for Id
    @property
    def id(self):
        return self.Id

    @property
    def Name(self):
        return self.com_object.Name

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name


class TabStop:

    def __init__(self, tabstop=None):
        self.com_object= tabstop

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

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

    def Clear(self):
        self.com_object.Clear()

    # Lower case alias for Clear
    def clear(self):
        return self.Clear()


class TabStops:

    def __init__(self, tabstops=None):
        self.com_object= tabstops

    def __call__(self, item):
        return TabStop(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def DefaultSpacing(self):
        return self.com_object.DefaultSpacing

    @DefaultSpacing.setter
    def DefaultSpacing(self, value):
        self.com_object.DefaultSpacing = value

    # Lower case aliases for DefaultSpacing
    @property
    def defaultspacing(self):
        return self.DefaultSpacing

    @defaultspacing.setter
    def defaultspacing(self, value):
        self.DefaultSpacing = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    def Add(self, Type=None, Position=None):
        arguments = com_arguments([unwrap(a) for a in [Type, Position]])
        return TabStop(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Type=None, Position=None):
        arguments = [Type, Position]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return TabStop(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class Tags:

    def __init__(self, tags=None):
        self.com_object= tags

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Add(self, Name=None, Value=None):
        arguments = com_arguments([unwrap(a) for a in [Name, Value]])
        self.com_object.Add(*arguments)

    # Lower case alias for Add
    def add(self, Name=None, Value=None):
        arguments = [Name, Value]
        return self.Add(*arguments)

    def Delete(self, Name=None):
        arguments = com_arguments([unwrap(a) for a in [Name]])
        self.com_object.Delete(*arguments)

    # Lower case alias for Delete
    def delete(self, Name=None):
        arguments = [Name]
        return self.Delete(*arguments)

    def Item(self, Name=None):
        arguments = com_arguments([unwrap(a) for a in [Name]])
        return self.com_object.Item(*arguments)

    # Lower case alias for Item
    def item(self, Name=None):
        arguments = [Name]
        return self.Item(*arguments)

    def Name(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return self.com_object.Name(*arguments)

    # Lower case alias for Name
    def name(self, Index=None):
        arguments = [Index]
        return self.Name(*arguments)

    def Value(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return self.com_object.Value(*arguments)

    # Lower case alias for Value
    def value(self, Index=None):
        arguments = [Index]
        return self.Value(*arguments)


class TextEffectFormat:

    def __init__(self, texteffectformat=None):
        self.com_object= texteffectformat

    @property
    def Alignment(self):
        return self.com_object.Alignment

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
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def FontBold(self):
        return self.com_object.FontBold

    @FontBold.setter
    def FontBold(self, value):
        self.com_object.FontBold = value

    # Lower case aliases for FontBold
    @property
    def fontbold(self):
        return self.FontBold

    @fontbold.setter
    def fontbold(self, value):
        self.FontBold = value

    @property
    def FontItalic(self):
        return self.com_object.FontItalic

    @FontItalic.setter
    def FontItalic(self, value):
        self.com_object.FontItalic = value

    # Lower case aliases for FontItalic
    @property
    def fontitalic(self):
        return self.FontItalic

    @fontitalic.setter
    def fontitalic(self, value):
        self.FontItalic = value

    @property
    def FontName(self):
        return self.com_object.FontName

    @FontName.setter
    def FontName(self, value):
        self.com_object.FontName = value

    # Lower case aliases for FontName
    @property
    def fontname(self):
        return self.FontName

    @fontname.setter
    def fontname(self, value):
        self.FontName = value

    @property
    def FontSize(self):
        return self.com_object.FontSize

    @FontSize.setter
    def FontSize(self, value):
        self.com_object.FontSize = value

    # Lower case aliases for FontSize
    @property
    def fontsize(self):
        return self.FontSize

    @fontsize.setter
    def fontsize(self, value):
        self.FontSize = value

    @property
    def KernedPairs(self):
        return self.com_object.KernedPairs

    @KernedPairs.setter
    def KernedPairs(self, value):
        self.com_object.KernedPairs = value

    # Lower case aliases for KernedPairs
    @property
    def kernedpairs(self):
        return self.KernedPairs

    @kernedpairs.setter
    def kernedpairs(self, value):
        self.KernedPairs = value

    @property
    def NormalizedHeight(self):
        return self.com_object.NormalizedHeight

    @NormalizedHeight.setter
    def NormalizedHeight(self, value):
        self.com_object.NormalizedHeight = value

    # Lower case aliases for NormalizedHeight
    @property
    def normalizedheight(self):
        return self.NormalizedHeight

    @normalizedheight.setter
    def normalizedheight(self, value):
        self.NormalizedHeight = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PresetShape(self):
        return self.com_object.PresetShape

    @PresetShape.setter
    def PresetShape(self, value):
        self.com_object.PresetShape = value

    # Lower case aliases for PresetShape
    @property
    def presetshape(self):
        return self.PresetShape

    @presetshape.setter
    def presetshape(self, value):
        self.PresetShape = value

    @property
    def PresetTextEffect(self):
        return self.com_object.PresetTextEffect

    @PresetTextEffect.setter
    def PresetTextEffect(self, value):
        self.com_object.PresetTextEffect = value

    # Lower case aliases for PresetTextEffect
    @property
    def presettexteffect(self):
        return self.PresetTextEffect

    @presettexteffect.setter
    def presettexteffect(self, value):
        self.PresetTextEffect = value

    @property
    def RotatedChars(self):
        return self.com_object.RotatedChars

    @RotatedChars.setter
    def RotatedChars(self, value):
        self.com_object.RotatedChars = value

    # Lower case aliases for RotatedChars
    @property
    def rotatedchars(self):
        return self.RotatedChars

    @rotatedchars.setter
    def rotatedchars(self, value):
        self.RotatedChars = value

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
    def Tracking(self):
        return self.com_object.Tracking

    @Tracking.setter
    def Tracking(self, value):
        self.com_object.Tracking = value

    # Lower case aliases for Tracking
    @property
    def tracking(self):
        return self.Tracking

    @tracking.setter
    def tracking(self, value):
        self.Tracking = value

    def ToggleVerticalText(self):
        self.com_object.ToggleVerticalText()

    # Lower case alias for ToggleVerticalText
    def toggleverticaltext(self):
        return self.ToggleVerticalText()


class TextFrame:

    def __init__(self, textframe=None):
        self.com_object= textframe

    @property
    def Application(self):
        return Application(self.com_object.Application)

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
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def HasText(self):
        return self.com_object.HasText

    # Lower case aliases for HasText
    @property
    def hastext(self):
        return self.HasText

    @property
    def HorizontalAnchor(self):
        return self.com_object.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.com_object.HorizontalAnchor = value

    # Lower case aliases for HorizontalAnchor
    @property
    def horizontalanchor(self):
        return self.HorizontalAnchor

    @horizontalanchor.setter
    def horizontalanchor(self, value):
        self.HorizontalAnchor = value

    @property
    def MarginBottom(self):
        return self.com_object.MarginBottom

    @MarginBottom.setter
    def MarginBottom(self, value):
        self.com_object.MarginBottom = value

    # Lower case aliases for MarginBottom
    @property
    def marginbottom(self):
        return self.MarginBottom

    @marginbottom.setter
    def marginbottom(self, value):
        self.MarginBottom = value

    @property
    def MarginLeft(self):
        return self.com_object.MarginLeft

    @MarginLeft.setter
    def MarginLeft(self, value):
        self.com_object.MarginLeft = value

    # Lower case aliases for MarginLeft
    @property
    def marginleft(self):
        return self.MarginLeft

    @marginleft.setter
    def marginleft(self, value):
        self.MarginLeft = value

    @property
    def MarginRight(self):
        return self.com_object.MarginRight

    @MarginRight.setter
    def MarginRight(self, value):
        self.com_object.MarginRight = value

    # Lower case aliases for MarginRight
    @property
    def marginright(self):
        return self.MarginRight

    @marginright.setter
    def marginright(self, value):
        self.MarginRight = value

    @property
    def MarginTop(self):
        return self.com_object.MarginTop

    @MarginTop.setter
    def MarginTop(self, value):
        self.com_object.MarginTop = value

    # Lower case aliases for MarginTop
    @property
    def margintop(self):
        return self.MarginTop

    @margintop.setter
    def margintop(self, value):
        self.MarginTop = value

    @property
    def Orientation(self):
        return self.com_object.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.com_object.Orientation = value

    # Lower case aliases for Orientation
    @property
    def orientation(self):
        return self.Orientation

    @orientation.setter
    def orientation(self, value):
        self.Orientation = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Ruler(self):
        return Ruler(self.com_object.Ruler)

    # Lower case aliases for Ruler
    @property
    def ruler(self):
        return self.Ruler

    @property
    def TextRange(self):
        return TextRange(self.com_object.TextRange)

    # Lower case aliases for TextRange
    @property
    def textrange(self):
        return self.TextRange

    @property
    def VerticalAnchor(self):
        return self.com_object.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.com_object.VerticalAnchor = value

    # Lower case aliases for VerticalAnchor
    @property
    def verticalanchor(self):
        return self.VerticalAnchor

    @verticalanchor.setter
    def verticalanchor(self, value):
        self.VerticalAnchor = value

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

    def DeleteText(self):
        self.com_object.DeleteText()

    # Lower case alias for DeleteText
    def deletetext(self):
        return self.DeleteText()


class TextFrame2:

    def __init__(self, textframe2=None):
        self.com_object= textframe2

    @property
    def Application(self):
        return Application(self.com_object.Application)

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
    def Column(self):
        return Column(self.com_object.Column)

    # Lower case aliases for Column
    @property
    def column(self):
        return self.Column

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def HasText(self):
        return self.com_object.HasText

    # Lower case aliases for HasText
    @property
    def hastext(self):
        return self.HasText

    @property
    def HorizontalAnchor(self):
        return self.com_object.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.com_object.HorizontalAnchor = value

    # Lower case aliases for HorizontalAnchor
    @property
    def horizontalanchor(self):
        return self.HorizontalAnchor

    @horizontalanchor.setter
    def horizontalanchor(self, value):
        self.HorizontalAnchor = value

    @property
    def MarginBottom(self):
        return self.com_object.MarginBottom

    @MarginBottom.setter
    def MarginBottom(self, value):
        self.com_object.MarginBottom = value

    # Lower case aliases for MarginBottom
    @property
    def marginbottom(self):
        return self.MarginBottom

    @marginbottom.setter
    def marginbottom(self, value):
        self.MarginBottom = value

    @property
    def MarginLeft(self):
        return self.com_object.MarginLeft

    @MarginLeft.setter
    def MarginLeft(self, value):
        self.com_object.MarginLeft = value

    # Lower case aliases for MarginLeft
    @property
    def marginleft(self):
        return self.MarginLeft

    @marginleft.setter
    def marginleft(self, value):
        self.MarginLeft = value

    @property
    def MarginRight(self):
        return self.com_object.MarginRight

    @MarginRight.setter
    def MarginRight(self, value):
        self.com_object.MarginRight = value

    # Lower case aliases for MarginRight
    @property
    def marginright(self):
        return self.MarginRight

    @marginright.setter
    def marginright(self, value):
        self.MarginRight = value

    @property
    def MarginTop(self):
        return self.com_object.MarginTop

    @MarginTop.setter
    def MarginTop(self, value):
        self.com_object.MarginTop = value

    # Lower case aliases for MarginTop
    @property
    def margintop(self):
        return self.MarginTop

    @margintop.setter
    def margintop(self, value):
        self.MarginTop = value

    @property
    def NoTextRotation(self):
        return self.com_object.NoTextRotation

    @NoTextRotation.setter
    def NoTextRotation(self, value):
        self.com_object.NoTextRotation = value

    # Lower case aliases for NoTextRotation
    @property
    def notextrotation(self):
        return self.NoTextRotation

    @notextrotation.setter
    def notextrotation(self, value):
        self.NoTextRotation = value

    @property
    def Orientation(self):
        return self.com_object.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.com_object.Orientation = value

    # Lower case aliases for Orientation
    @property
    def orientation(self):
        return self.Orientation

    @orientation.setter
    def orientation(self, value):
        self.Orientation = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PathFormat(self):
        return self.com_object.PathFormat

    @PathFormat.setter
    def PathFormat(self, value):
        self.com_object.PathFormat = value

    # Lower case aliases for PathFormat
    @property
    def pathformat(self):
        return self.PathFormat

    @pathformat.setter
    def pathformat(self, value):
        self.PathFormat = value

    @property
    def Ruler(self):
        return self.com_object.Ruler

    # Lower case aliases for Ruler
    @property
    def ruler(self):
        return self.Ruler

    @property
    def TextRange(self):
        return self.com_object.TextRange

    # Lower case aliases for TextRange
    @property
    def textrange(self):
        return self.TextRange

    @property
    def ThreeD(self):
        return ThreeDFormat(self.com_object.ThreeD)

    # Lower case aliases for ThreeD
    @property
    def threed(self):
        return self.ThreeD

    @property
    def VerticalAnchor(self):
        return self.com_object.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.com_object.VerticalAnchor = value

    # Lower case aliases for VerticalAnchor
    @property
    def verticalanchor(self):
        return self.VerticalAnchor

    @verticalanchor.setter
    def verticalanchor(self, value):
        self.VerticalAnchor = value

    @property
    def WarpFormat(self):
        return self.com_object.WarpFormat

    @WarpFormat.setter
    def WarpFormat(self, value):
        self.com_object.WarpFormat = value

    # Lower case aliases for WarpFormat
    @property
    def warpformat(self):
        return self.WarpFormat

    @warpformat.setter
    def warpformat(self, value):
        self.WarpFormat = value

    @property
    def WordArtFormat(self):
        return self.com_object.WordArtFormat

    @WordArtFormat.setter
    def WordArtFormat(self, value):
        self.com_object.WordArtFormat = value

    # Lower case aliases for WordArtFormat
    @property
    def wordartformat(self):
        return self.WordArtFormat

    @wordartformat.setter
    def wordartformat(self, value):
        self.WordArtFormat = value

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

    def DeleteText(self):
        return self.com_object.DeleteText()

    # Lower case alias for DeleteText
    def deletetext(self):
        return self.DeleteText()


class TextRange:

    def __init__(self, textrange=None):
        self.com_object= textrange

    @property
    def ActionSettings(self):
        return ActionSettings(self.com_object.ActionSettings)

    # Lower case aliases for ActionSettings
    @property
    def actionsettings(self):
        return self.ActionSettings

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def BoundHeight(self):
        return self.com_object.BoundHeight

    # Lower case aliases for BoundHeight
    @property
    def boundheight(self):
        return self.BoundHeight

    @property
    def BoundLeft(self):
        return self.com_object.BoundLeft

    # Lower case aliases for BoundLeft
    @property
    def boundleft(self):
        return self.BoundLeft

    @property
    def BoundTop(self):
        return self.com_object.BoundTop

    # Lower case aliases for BoundTop
    @property
    def boundtop(self):
        return self.BoundTop

    @property
    def BoundWidth(self):
        return self.com_object.BoundWidth

    # Lower case aliases for BoundWidth
    @property
    def boundwidth(self):
        return self.BoundWidth

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Font(self):
        return Font(self.com_object.Font)

    # Lower case aliases for Font
    @property
    def font(self):
        return self.Font

    @property
    def IndentLevel(self):
        return self.com_object.IndentLevel

    @IndentLevel.setter
    def IndentLevel(self, value):
        self.com_object.IndentLevel = value

    # Lower case aliases for IndentLevel
    @property
    def indentlevel(self):
        return self.IndentLevel

    @indentlevel.setter
    def indentlevel(self, value):
        self.IndentLevel = value

    @property
    def LanguageID(self):
        return self.com_object.LanguageID

    @LanguageID.setter
    def LanguageID(self, value):
        self.com_object.LanguageID = value

    # Lower case aliases for LanguageID
    @property
    def languageid(self):
        return self.LanguageID

    @languageid.setter
    def languageid(self, value):
        self.LanguageID = value

    @property
    def Length(self):
        return self.com_object.Length

    # Lower case aliases for Length
    @property
    def length(self):
        return self.Length

    @property
    def ParagraphFormat(self):
        return ParagraphFormat(self.com_object.ParagraphFormat)

    # Lower case aliases for ParagraphFormat
    @property
    def paragraphformat(self):
        return self.ParagraphFormat

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Start(self):
        return self.com_object.Start

    # Lower case aliases for Start
    @property
    def start(self):
        return self.Start

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

    def AddPeriods(self):
        self.com_object.AddPeriods()

    # Lower case alias for AddPeriods
    def addperiods(self):
        return self.AddPeriods()

    def ChangeCase(self, Type=None):
        arguments = com_arguments([unwrap(a) for a in [Type]])
        self.com_object.ChangeCase(*arguments)

    # Lower case alias for ChangeCase
    def changecase(self, Type=None):
        arguments = [Type]
        return self.ChangeCase(*arguments)

    def Characters(self, Start=None, Length=None):
        arguments = com_arguments([unwrap(a) for a in [Start, Length]])
        return TextRange(self.com_object.Characters(*arguments))

    # Lower case alias for Characters
    def characters(self, Start=None, Length=None):
        arguments = [Start, Length]
        return self.Characters(*arguments)

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

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Find(self, FindWhat=None, After=None, MatchCase=None, WholeWords=None):
        arguments = com_arguments([unwrap(a) for a in [FindWhat, After, MatchCase, WholeWords]])
        return TextRange(self.com_object.Find(*arguments))

    # Lower case alias for Find
    def find(self, FindWhat=None, After=None, MatchCase=None, WholeWords=None):
        arguments = [FindWhat, After, MatchCase, WholeWords]
        return self.Find(*arguments)

    def InsertAfter(self, NewText=None):
        arguments = com_arguments([unwrap(a) for a in [NewText]])
        self.com_object.InsertAfter(*arguments)

    # Lower case alias for InsertAfter
    def insertafter(self, NewText=None):
        arguments = [NewText]
        return self.InsertAfter(*arguments)

    def InsertBefore(self, NewText=None):
        arguments = com_arguments([unwrap(a) for a in [NewText]])
        self.com_object.InsertBefore(*arguments)

    # Lower case alias for InsertBefore
    def insertbefore(self, NewText=None):
        arguments = [NewText]
        return self.InsertBefore(*arguments)

    def InsertDateTime(self, DateTimeFormat=None, InsertAsField=None):
        arguments = com_arguments([unwrap(a) for a in [DateTimeFormat, InsertAsField]])
        return TextRange(self.com_object.InsertDateTime(*arguments))

    # Lower case alias for InsertDateTime
    def insertdatetime(self, DateTimeFormat=None, InsertAsField=None):
        arguments = [DateTimeFormat, InsertAsField]
        return self.InsertDateTime(*arguments)

    def InsertSlideNumber(self):
        return TextRange(self.com_object.InsertSlideNumber())

    # Lower case alias for InsertSlideNumber
    def insertslidenumber(self):
        return self.InsertSlideNumber()

    def InsertSymbol(self, FontName=None, CharNumber=None, UniCode=None):
        arguments = com_arguments([unwrap(a) for a in [FontName, CharNumber, UniCode]])
        return TextRange(self.com_object.InsertSymbol(*arguments))

    # Lower case alias for InsertSymbol
    def insertsymbol(self, FontName=None, CharNumber=None, UniCode=None):
        arguments = [FontName, CharNumber, UniCode]
        return self.InsertSymbol(*arguments)

    def Lines(self, Start=None, Length=None):
        arguments = com_arguments([unwrap(a) for a in [Start, Length]])
        return TextRange(self.com_object.Lines(*arguments))

    # Lower case alias for Lines
    def lines(self, Start=None, Length=None):
        arguments = [Start, Length]
        return self.Lines(*arguments)

    def LtrRun(self):
        self.com_object.LtrRun()

    # Lower case alias for LtrRun
    def ltrrun(self):
        return self.LtrRun()

    def Paragraphs(self, Start=None, Length=None):
        arguments = com_arguments([unwrap(a) for a in [Start, Length]])
        return TextRange(self.com_object.Paragraphs(*arguments))

    # Lower case alias for Paragraphs
    def paragraphs(self, Start=None, Length=None):
        arguments = [Start, Length]
        return self.Paragraphs(*arguments)

    def Paste(self):
        return TextRange(self.com_object.Paste())

    # Lower case alias for Paste
    def paste(self):
        return self.Paste()

    def PasteSpecial(self, DataType=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Link=None):
        arguments = com_arguments([unwrap(a) for a in [DataType, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Link]])
        return TextRange(self.com_object.PasteSpecial(*arguments))

    # Lower case alias for PasteSpecial
    def pastespecial(self, DataType=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Link=None):
        arguments = [DataType, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Link]
        return self.PasteSpecial(*arguments)

    def RemovePeriods(self):
        self.com_object.RemovePeriods()

    # Lower case alias for RemovePeriods
    def removeperiods(self):
        return self.RemovePeriods()

    def Replace(self, FindWhat=None, ReplaceWhat=None, After=None, MatchCase=None, WholeWords=None):
        arguments = com_arguments([unwrap(a) for a in [FindWhat, ReplaceWhat, After, MatchCase, WholeWords]])
        return TextRange(self.com_object.Replace(*arguments))

    # Lower case alias for Replace
    def replace(self, FindWhat=None, ReplaceWhat=None, After=None, MatchCase=None, WholeWords=None):
        arguments = [FindWhat, ReplaceWhat, After, MatchCase, WholeWords]
        return self.Replace(*arguments)

    def RotatedBounds(self, X1=None, Y1=None, X2=None, Y2=None, X3=None, Y3=None, X4=None, Y4=None):
        arguments = com_arguments([unwrap(a) for a in [X1, Y1, X2, Y2, X3, Y3, X4, Y4]])
        self.com_object.RotatedBounds(*arguments)

    # Lower case alias for RotatedBounds
    def rotatedbounds(self, X1=None, Y1=None, X2=None, Y2=None, X3=None, Y3=None, X4=None, Y4=None):
        arguments = [X1, Y1, X2, Y2, X3, Y3, X4, Y4]
        return self.RotatedBounds(*arguments)

    def RtlRun(self):
        self.com_object.RtlRun()

    # Lower case alias for RtlRun
    def rtlrun(self):
        return self.RtlRun()

    def Runs(self, Start=None, Length=None):
        arguments = com_arguments([unwrap(a) for a in [Start, Length]])
        return TextRange(self.com_object.Runs(*arguments))

    # Lower case alias for Runs
    def runs(self, Start=None, Length=None):
        arguments = [Start, Length]
        return self.Runs(*arguments)

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()

    def Sentences(self, Start=None, Length=None):
        arguments = com_arguments([unwrap(a) for a in [Start, Length]])
        return TextRange(self.com_object.Sentences(*arguments))

    # Lower case alias for Sentences
    def sentences(self, Start=None, Length=None):
        arguments = [Start, Length]
        return self.Sentences(*arguments)

    def TrimText(self):
        return TextRange(self.com_object.TrimText())

    # Lower case alias for TrimText
    def trimtext(self):
        return self.TrimText()

    def Words(self, Start=None, Length=None):
        arguments = com_arguments([unwrap(a) for a in [Start, Length]])
        return TextRange(self.com_object.Words(*arguments))

    # Lower case alias for Words
    def words(self, Start=None, Length=None):
        arguments = [Start, Length]
        return self.Words(*arguments)


class TextStyle:

    def __init__(self, textstyle=None):
        self.com_object= textstyle

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Levels(self):
        return TextStyleLevels(self.com_object.Levels)

    # Lower case aliases for Levels
    @property
    def levels(self):
        return self.Levels

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Ruler(self):
        return Ruler(self.com_object.Ruler)

    # Lower case aliases for Ruler
    @property
    def ruler(self):
        return self.Ruler

    @property
    def TextFrame(self):
        return TextFrame(self.com_object.TextFrame)

    # Lower case aliases for TextFrame
    @property
    def textframe(self):
        return self.TextFrame


class TextStyleLevel:

    def __init__(self, textstylelevel=None):
        self.com_object= textstylelevel

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Font(self):
        return Font(self.com_object.Font)

    # Lower case aliases for Font
    @property
    def font(self):
        return self.Font

    @property
    def ParagraphFormat(self):
        return ParagraphFormat(self.com_object.ParagraphFormat)

    # Lower case aliases for ParagraphFormat
    @property
    def paragraphformat(self):
        return self.ParagraphFormat

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent


class TextStyleLevels:

    def __init__(self, textstylelevels=None):
        self.com_object= textstylelevels

    def __call__(self, item):
        return TextStyleLevel(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return TextStyleLevel(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class TextStyles:

    def __init__(self, textstyles=None):
        self.com_object= textstyles

    def __call__(self, item):
        return TextStyle(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

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

    def Item(self, Type=None):
        arguments = com_arguments([unwrap(a) for a in [Type]])
        return TextStyle(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Type=None):
        arguments = [Type]
        return self.Item(*arguments)


class ThreeDFormat:

    def __init__(self, threedformat=None):
        self.com_object= threedformat

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def BevelBottomDepth(self):
        return ThreeDFormat(self.com_object.BevelBottomDepth)

    @BevelBottomDepth.setter
    def BevelBottomDepth(self, value):
        self.com_object.BevelBottomDepth = value

    # Lower case aliases for BevelBottomDepth
    @property
    def bevelbottomdepth(self):
        return self.BevelBottomDepth

    @bevelbottomdepth.setter
    def bevelbottomdepth(self, value):
        self.BevelBottomDepth = value

    @property
    def BevelBottomInset(self):
        return ThreeDFormat(self.com_object.BevelBottomInset)

    @BevelBottomInset.setter
    def BevelBottomInset(self, value):
        self.com_object.BevelBottomInset = value

    # Lower case aliases for BevelBottomInset
    @property
    def bevelbottominset(self):
        return self.BevelBottomInset

    @bevelbottominset.setter
    def bevelbottominset(self, value):
        self.BevelBottomInset = value

    @property
    def BevelBottomType(self):
        return self.com_object.BevelBottomType

    @BevelBottomType.setter
    def BevelBottomType(self, value):
        self.com_object.BevelBottomType = value

    # Lower case aliases for BevelBottomType
    @property
    def bevelbottomtype(self):
        return self.BevelBottomType

    @bevelbottomtype.setter
    def bevelbottomtype(self, value):
        self.BevelBottomType = value

    @property
    def BevelTopDepth(self):
        return ThreeDFormat(self.com_object.BevelTopDepth)

    @BevelTopDepth.setter
    def BevelTopDepth(self, value):
        self.com_object.BevelTopDepth = value

    # Lower case aliases for BevelTopDepth
    @property
    def beveltopdepth(self):
        return self.BevelTopDepth

    @beveltopdepth.setter
    def beveltopdepth(self, value):
        self.BevelTopDepth = value

    @property
    def BevelTopInset(self):
        return ThreeDFormat(self.com_object.BevelTopInset)

    @BevelTopInset.setter
    def BevelTopInset(self, value):
        self.com_object.BevelTopInset = value

    # Lower case aliases for BevelTopInset
    @property
    def beveltopinset(self):
        return self.BevelTopInset

    @beveltopinset.setter
    def beveltopinset(self, value):
        self.BevelTopInset = value

    @property
    def BevelTopType(self):
        return self.com_object.BevelTopType

    @BevelTopType.setter
    def BevelTopType(self, value):
        self.com_object.BevelTopType = value

    # Lower case aliases for BevelTopType
    @property
    def beveltoptype(self):
        return self.BevelTopType

    @beveltoptype.setter
    def beveltoptype(self, value):
        self.BevelTopType = value

    @property
    def ContourColor(self):
        return ColorFormat(self.com_object.ContourColor)

    # Lower case aliases for ContourColor
    @property
    def contourcolor(self):
        return self.ContourColor

    @property
    def ContourWidth(self):
        return ThreeDFormat(self.com_object.ContourWidth)

    @ContourWidth.setter
    def ContourWidth(self, value):
        self.com_object.ContourWidth = value

    # Lower case aliases for ContourWidth
    @property
    def contourwidth(self):
        return self.ContourWidth

    @contourwidth.setter
    def contourwidth(self, value):
        self.ContourWidth = value

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Depth(self):
        return self.com_object.Depth

    @Depth.setter
    def Depth(self, value):
        self.com_object.Depth = value

    # Lower case aliases for Depth
    @property
    def depth(self):
        return self.Depth

    @depth.setter
    def depth(self, value):
        self.Depth = value

    @property
    def ExtrusionColor(self):
        return ColorFormat(self.com_object.ExtrusionColor)

    # Lower case aliases for ExtrusionColor
    @property
    def extrusioncolor(self):
        return self.ExtrusionColor

    @property
    def ExtrusionColorType(self):
        return self.com_object.ExtrusionColorType

    @ExtrusionColorType.setter
    def ExtrusionColorType(self, value):
        self.com_object.ExtrusionColorType = value

    # Lower case aliases for ExtrusionColorType
    @property
    def extrusioncolortype(self):
        return self.ExtrusionColorType

    @extrusioncolortype.setter
    def extrusioncolortype(self, value):
        self.ExtrusionColorType = value

    @property
    def FieldOfView(self):
        return ThreeDFormat(self.com_object.FieldOfView)

    @FieldOfView.setter
    def FieldOfView(self, value):
        self.com_object.FieldOfView = value

    # Lower case aliases for FieldOfView
    @property
    def fieldofview(self):
        return self.FieldOfView

    @fieldofview.setter
    def fieldofview(self, value):
        self.FieldOfView = value

    @property
    def LightAngle(self):
        return self.com_object.LightAngle

    @LightAngle.setter
    def LightAngle(self, value):
        self.com_object.LightAngle = value

    # Lower case aliases for LightAngle
    @property
    def lightangle(self):
        return self.LightAngle

    @lightangle.setter
    def lightangle(self, value):
        self.LightAngle = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Perspective(self):
        return self.com_object.Perspective

    @Perspective.setter
    def Perspective(self, value):
        self.com_object.Perspective = value

    # Lower case aliases for Perspective
    @property
    def perspective(self):
        return self.Perspective

    @perspective.setter
    def perspective(self, value):
        self.Perspective = value

    @property
    def PresetCamera(self):
        return ThreeDFormat(self.com_object.PresetCamera)

    # Lower case aliases for PresetCamera
    @property
    def presetcamera(self):
        return self.PresetCamera

    @property
    def PresetExtrusionDirection(self):
        return self.com_object.PresetExtrusionDirection

    # Lower case aliases for PresetExtrusionDirection
    @property
    def presetextrusiondirection(self):
        return self.PresetExtrusionDirection

    @property
    def PresetLighting(self):
        return ThreeDFormat(self.com_object.PresetLighting)

    @PresetLighting.setter
    def PresetLighting(self, value):
        self.com_object.PresetLighting = value

    # Lower case aliases for PresetLighting
    @property
    def presetlighting(self):
        return self.PresetLighting

    @presetlighting.setter
    def presetlighting(self, value):
        self.PresetLighting = value

    @property
    def PresetLightingDirection(self):
        return self.com_object.PresetLightingDirection

    @PresetLightingDirection.setter
    def PresetLightingDirection(self, value):
        self.com_object.PresetLightingDirection = value

    # Lower case aliases for PresetLightingDirection
    @property
    def presetlightingdirection(self):
        return self.PresetLightingDirection

    @presetlightingdirection.setter
    def presetlightingdirection(self, value):
        self.PresetLightingDirection = value

    @property
    def PresetLightingSoftness(self):
        return self.com_object.PresetLightingSoftness

    @PresetLightingSoftness.setter
    def PresetLightingSoftness(self, value):
        self.com_object.PresetLightingSoftness = value

    # Lower case aliases for PresetLightingSoftness
    @property
    def presetlightingsoftness(self):
        return self.PresetLightingSoftness

    @presetlightingsoftness.setter
    def presetlightingsoftness(self, value):
        self.PresetLightingSoftness = value

    @property
    def PresetMaterial(self):
        return self.com_object.PresetMaterial

    @PresetMaterial.setter
    def PresetMaterial(self, value):
        self.com_object.PresetMaterial = value

    # Lower case aliases for PresetMaterial
    @property
    def presetmaterial(self):
        return self.PresetMaterial

    @presetmaterial.setter
    def presetmaterial(self, value):
        self.PresetMaterial = value

    @property
    def PresetThreeDFormat(self):
        return self.com_object.PresetThreeDFormat

    # Lower case aliases for PresetThreeDFormat
    @property
    def presetthreedformat(self):
        return self.PresetThreeDFormat

    @property
    def ProjectText(self):
        return self.com_object.ProjectText

    @ProjectText.setter
    def ProjectText(self, value):
        self.com_object.ProjectText = value

    # Lower case aliases for ProjectText
    @property
    def projecttext(self):
        return self.ProjectText

    @projecttext.setter
    def projecttext(self, value):
        self.ProjectText = value

    @property
    def RotationX(self):
        return self.com_object.RotationX

    @RotationX.setter
    def RotationX(self, value):
        self.com_object.RotationX = value

    # Lower case aliases for RotationX
    @property
    def rotationx(self):
        return self.RotationX

    @rotationx.setter
    def rotationx(self, value):
        self.RotationX = value

    @property
    def RotationY(self):
        return self.com_object.RotationY

    @RotationY.setter
    def RotationY(self, value):
        self.com_object.RotationY = value

    # Lower case aliases for RotationY
    @property
    def rotationy(self):
        return self.RotationY

    @rotationy.setter
    def rotationy(self, value):
        self.RotationY = value

    @property
    def RotationZ(self):
        return ThreeDFormat(self.com_object.RotationZ)

    @RotationZ.setter
    def RotationZ(self, value):
        self.com_object.RotationZ = value

    # Lower case aliases for RotationZ
    @property
    def rotationz(self):
        return self.RotationZ

    @rotationz.setter
    def rotationz(self, value):
        self.RotationZ = value

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

    @property
    def Z(self):
        return ThreeDFormat(self.com_object.Z)

    @Z.setter
    def Z(self, value):
        self.com_object.Z = value

    # Lower case aliases for Z
    @property
    def z(self):
        return self.Z

    @z.setter
    def z(self, value):
        self.Z = value

    def IncrementRotationHorizontal(self, Increment=None):
        arguments = com_arguments([unwrap(a) for a in [Increment]])
        self.com_object.IncrementRotationHorizontal(*arguments)

    # Lower case alias for IncrementRotationHorizontal
    def incrementrotationhorizontal(self, Increment=None):
        arguments = [Increment]
        return self.IncrementRotationHorizontal(*arguments)

    def IncrementRotationVertical(self, Increment=None):
        arguments = com_arguments([unwrap(a) for a in [Increment]])
        self.com_object.IncrementRotationVertical(*arguments)

    # Lower case alias for IncrementRotationVertical
    def incrementrotationvertical(self, Increment=None):
        arguments = [Increment]
        return self.IncrementRotationVertical(*arguments)

    def IncrementRotationX(self, Increment=None):
        arguments = com_arguments([unwrap(a) for a in [Increment]])
        self.com_object.IncrementRotationX(*arguments)

    # Lower case alias for IncrementRotationX
    def incrementrotationx(self, Increment=None):
        arguments = [Increment]
        return self.IncrementRotationX(*arguments)

    def IncrementRotationY(self, Increment=None):
        arguments = com_arguments([unwrap(a) for a in [Increment]])
        self.com_object.IncrementRotationY(*arguments)

    # Lower case alias for IncrementRotationY
    def incrementrotationy(self, Increment=None):
        arguments = [Increment]
        return self.IncrementRotationY(*arguments)

    def IncrementRotationZ(self, Increment=None):
        arguments = com_arguments([unwrap(a) for a in [Increment]])
        self.com_object.IncrementRotationZ(*arguments)

    # Lower case alias for IncrementRotationZ
    def incrementrotationz(self, Increment=None):
        arguments = [Increment]
        return self.IncrementRotationZ(*arguments)

    def ResetRotation(self):
        self.com_object.ResetRotation()

    # Lower case alias for ResetRotation
    def resetrotation(self):
        return self.ResetRotation()

    def SetExtrusionDirection(self, PresetExtrusionDirection=None):
        arguments = com_arguments([unwrap(a) for a in [PresetExtrusionDirection]])
        self.com_object.SetExtrusionDirection(*arguments)

    # Lower case alias for SetExtrusionDirection
    def setextrusiondirection(self, PresetExtrusionDirection=None):
        arguments = [PresetExtrusionDirection]
        return self.SetExtrusionDirection(*arguments)

    def SetPresetCamera(self, PresetCamera=None):
        arguments = com_arguments([unwrap(a) for a in [PresetCamera]])
        self.com_object.SetPresetCamera(*arguments)

    # Lower case alias for SetPresetCamera
    def setpresetcamera(self, PresetCamera=None):
        arguments = [PresetCamera]
        return self.SetPresetCamera(*arguments)

    def SetThreeDFormat(self, PresetThreeDFormat=None):
        arguments = com_arguments([unwrap(a) for a in [PresetThreeDFormat]])
        self.com_object.SetThreeDFormat(*arguments)

    # Lower case alias for SetThreeDFormat
    def setthreedformat(self, PresetThreeDFormat=None):
        arguments = [PresetThreeDFormat]
        return self.SetThreeDFormat(*arguments)


class TickLabels:

    def __init__(self, ticklabels=None):
        self.com_object= ticklabels

    @property
    def Alignment(self):
        return self.com_object.Alignment

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
    def Application(self):
        return self.com_object.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Depth(self):
        return self.com_object.Depth

    # Lower case aliases for Depth
    @property
    def depth(self):
        return self.Depth

    @property
    def Font(self):
        return ChartFont(self.com_object.Font)

    # Lower case aliases for Font
    @property
    def font(self):
        return self.Font

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

    @property
    def MultiLevel(self):
        return self.com_object.MultiLevel

    @MultiLevel.setter
    def MultiLevel(self, value):
        self.com_object.MultiLevel = value

    # Lower case aliases for MultiLevel
    @property
    def multilevel(self):
        return self.MultiLevel

    @multilevel.setter
    def multilevel(self, value):
        self.MultiLevel = value

    @property
    def Name(self):
        return self.com_object.Name

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def NumberFormat(self):
        return self.com_object.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value):
        self.com_object.NumberFormat = value

    # Lower case aliases for NumberFormat
    @property
    def numberformat(self):
        return self.NumberFormat

    @numberformat.setter
    def numberformat(self, value):
        self.NumberFormat = value

    @property
    def NumberFormatLinked(self):
        return self.com_object.NumberFormatLinked

    @NumberFormatLinked.setter
    def NumberFormatLinked(self, value):
        self.com_object.NumberFormatLinked = value

    # Lower case aliases for NumberFormatLinked
    @property
    def numberformatlinked(self):
        return self.NumberFormatLinked

    @numberformatlinked.setter
    def numberformatlinked(self, value):
        self.NumberFormatLinked = value

    @property
    def NumberFormatLocal(self):
        return self.com_object.NumberFormatLocal

    @NumberFormatLocal.setter
    def NumberFormatLocal(self, value):
        self.com_object.NumberFormatLocal = value

    # Lower case aliases for NumberFormatLocal
    @property
    def numberformatlocal(self):
        return self.NumberFormatLocal

    @numberformatlocal.setter
    def numberformatlocal(self, value):
        self.NumberFormatLocal = value

    @property
    def Offset(self):
        return self.com_object.Offset

    @Offset.setter
    def Offset(self, value):
        self.com_object.Offset = value

    # Lower case aliases for Offset
    @property
    def offset(self):
        return self.Offset

    @offset.setter
    def offset(self, value):
        self.Offset = value

    @property
    def Orientation(self):
        return self.com_object.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.com_object.Orientation = value

    # Lower case aliases for Orientation
    @property
    def orientation(self):
        return self.Orientation

    @orientation.setter
    def orientation(self, value):
        self.Orientation = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ReadingOrder(self):
        return XlReadingOrder(self.com_object.ReadingOrder)

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.com_object.ReadingOrder = value

    # Lower case aliases for ReadingOrder
    @property
    def readingorder(self):
        return self.ReadingOrder

    @readingorder.setter
    def readingorder(self, value):
        self.ReadingOrder = value

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class TimeLine:

    def __init__(self, timeline=None):
        self.com_object= timeline

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def InteractiveSequences(self):
        return Sequences(self.com_object.InteractiveSequences)

    # Lower case aliases for InteractiveSequences
    @property
    def interactivesequences(self):
        return self.InteractiveSequences

    @property
    def MainSequence(self):
        return Sequence(self.com_object.MainSequence)

    # Lower case aliases for MainSequence
    @property
    def mainsequence(self):
        return self.MainSequence

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent


class Timing:

    def __init__(self, timing=None):
        self.com_object= timing

    @property
    def Accelerate(self):
        return self.com_object.Accelerate

    @Accelerate.setter
    def Accelerate(self, value):
        self.com_object.Accelerate = value

    # Lower case aliases for Accelerate
    @property
    def accelerate(self):
        return self.Accelerate

    @accelerate.setter
    def accelerate(self, value):
        self.Accelerate = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AutoReverse(self):
        return self.com_object.AutoReverse

    @AutoReverse.setter
    def AutoReverse(self, value):
        self.com_object.AutoReverse = value

    # Lower case aliases for AutoReverse
    @property
    def autoreverse(self):
        return self.AutoReverse

    @autoreverse.setter
    def autoreverse(self, value):
        self.AutoReverse = value

    @property
    def BounceEnd(self):
        return self.com_object.BounceEnd

    @BounceEnd.setter
    def BounceEnd(self, value):
        self.com_object.BounceEnd = value

    # Lower case aliases for BounceEnd
    @property
    def bounceend(self):
        return self.BounceEnd

    @bounceend.setter
    def bounceend(self, value):
        self.BounceEnd = value

    @property
    def BounceEndIntensity(self):
        return self.com_object.BounceEndIntensity

    @BounceEndIntensity.setter
    def BounceEndIntensity(self, value):
        self.com_object.BounceEndIntensity = value

    # Lower case aliases for BounceEndIntensity
    @property
    def bounceendintensity(self):
        return self.BounceEndIntensity

    @bounceendintensity.setter
    def bounceendintensity(self, value):
        self.BounceEndIntensity = value

    @property
    def Decelerate(self):
        return self.com_object.Decelerate

    @Decelerate.setter
    def Decelerate(self, value):
        self.com_object.Decelerate = value

    # Lower case aliases for Decelerate
    @property
    def decelerate(self):
        return self.Decelerate

    @decelerate.setter
    def decelerate(self, value):
        self.Decelerate = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RepeatCount(self):
        return self.com_object.RepeatCount

    @RepeatCount.setter
    def RepeatCount(self, value):
        self.com_object.RepeatCount = value

    # Lower case aliases for RepeatCount
    @property
    def repeatcount(self):
        return self.RepeatCount

    @repeatcount.setter
    def repeatcount(self, value):
        self.RepeatCount = value

    @property
    def RepeatDuration(self):
        return self.com_object.RepeatDuration

    @RepeatDuration.setter
    def RepeatDuration(self, value):
        self.com_object.RepeatDuration = value

    # Lower case aliases for RepeatDuration
    @property
    def repeatduration(self):
        return self.RepeatDuration

    @repeatduration.setter
    def repeatduration(self, value):
        self.RepeatDuration = value

    @property
    def Restart(self):
        return self.com_object.Restart

    @Restart.setter
    def Restart(self, value):
        self.com_object.Restart = value

    # Lower case aliases for Restart
    @property
    def restart(self):
        return self.Restart

    @restart.setter
    def restart(self, value):
        self.Restart = value

    @property
    def RewindAtEnd(self):
        return self.com_object.RewindAtEnd

    @RewindAtEnd.setter
    def RewindAtEnd(self, value):
        self.com_object.RewindAtEnd = value

    # Lower case aliases for RewindAtEnd
    @property
    def rewindatend(self):
        return self.RewindAtEnd

    @rewindatend.setter
    def rewindatend(self, value):
        self.RewindAtEnd = value

    @property
    def SmoothEnd(self):
        return self.com_object.SmoothEnd

    @SmoothEnd.setter
    def SmoothEnd(self, value):
        self.com_object.SmoothEnd = value

    # Lower case aliases for SmoothEnd
    @property
    def smoothend(self):
        return self.SmoothEnd

    @smoothend.setter
    def smoothend(self, value):
        self.SmoothEnd = value

    @property
    def SmoothStart(self):
        return self.com_object.SmoothStart

    @SmoothStart.setter
    def SmoothStart(self, value):
        self.com_object.SmoothStart = value

    # Lower case aliases for SmoothStart
    @property
    def smoothstart(self):
        return self.SmoothStart

    @smoothstart.setter
    def smoothstart(self, value):
        self.SmoothStart = value

    @property
    def Speed(self):
        return self.com_object.Speed

    @Speed.setter
    def Speed(self, value):
        self.com_object.Speed = value

    # Lower case aliases for Speed
    @property
    def speed(self):
        return self.Speed

    @speed.setter
    def speed(self, value):
        self.Speed = value

    @property
    def triggerBookmark(self):
        return self.com_object.triggerBookmark

    @triggerBookmark.setter
    def triggerBookmark(self, value):
        self.com_object.triggerBookmark = value

    # Lower case aliases for triggerBookmark
    @property
    def triggerbookmark(self):
        return self.triggerBookmark

    @triggerbookmark.setter
    def triggerbookmark(self, value):
        self.triggerBookmark = value

    @property
    def TriggerDelayTime(self):
        return self.com_object.TriggerDelayTime

    @TriggerDelayTime.setter
    def TriggerDelayTime(self, value):
        self.com_object.TriggerDelayTime = value

    # Lower case aliases for TriggerDelayTime
    @property
    def triggerdelaytime(self):
        return self.TriggerDelayTime

    @triggerdelaytime.setter
    def triggerdelaytime(self, value):
        self.TriggerDelayTime = value

    @property
    def TriggerShape(self):
        return self.com_object.TriggerShape

    @TriggerShape.setter
    def TriggerShape(self, value):
        self.com_object.TriggerShape = value

    # Lower case aliases for TriggerShape
    @property
    def triggershape(self):
        return self.TriggerShape

    @triggershape.setter
    def triggershape(self, value):
        self.TriggerShape = value

    @property
    def TriggerType(self):
        return self.com_object.TriggerType

    @TriggerType.setter
    def TriggerType(self, value):
        self.com_object.TriggerType = value

    # Lower case aliases for TriggerType
    @property
    def triggertype(self):
        return self.TriggerType

    @triggertype.setter
    def triggertype(self, value):
        self.TriggerType = value


class Trendline:

    def __init__(self, trendline=None):
        self.com_object= trendline

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Backward2(self):
        return self.com_object.Backward2

    @Backward2.setter
    def Backward2(self, value):
        self.com_object.Backward2 = value

    # Lower case aliases for Backward2
    @property
    def backward2(self):
        return self.Backward2

    @backward2.setter
    def backward2(self, value):
        self.Backward2 = value

    @property
    def Border(self):
        return ChartBorder(self.com_object.Border)

    # Lower case aliases for Border
    @property
    def border(self):
        return self.Border

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def DataLabel(self):
        return DataLabel(self.com_object.DataLabel)

    # Lower case aliases for DataLabel
    @property
    def datalabel(self):
        return self.DataLabel

    @property
    def DisplayEquation(self):
        return self.com_object.DisplayEquation

    @DisplayEquation.setter
    def DisplayEquation(self, value):
        self.com_object.DisplayEquation = value

    # Lower case aliases for DisplayEquation
    @property
    def displayequation(self):
        return self.DisplayEquation

    @displayequation.setter
    def displayequation(self, value):
        self.DisplayEquation = value

    @property
    def DisplayRSquared(self):
        return self.com_object.DisplayRSquared

    @DisplayRSquared.setter
    def DisplayRSquared(self, value):
        self.com_object.DisplayRSquared = value

    # Lower case aliases for DisplayRSquared
    @property
    def displayrsquared(self):
        return self.DisplayRSquared

    @displayrsquared.setter
    def displayrsquared(self, value):
        self.DisplayRSquared = value

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

    @property
    def Forward2(self):
        return self.com_object.Forward2

    @Forward2.setter
    def Forward2(self, value):
        self.com_object.Forward2 = value

    # Lower case aliases for Forward2
    @property
    def forward2(self):
        return self.Forward2

    @forward2.setter
    def forward2(self, value):
        self.Forward2 = value

    @property
    def Index(self):
        return self.com_object.Index

    # Lower case aliases for Index
    @property
    def index(self):
        return self.Index

    @property
    def Intercept(self):
        return self.com_object.Intercept

    @Intercept.setter
    def Intercept(self, value):
        self.com_object.Intercept = value

    # Lower case aliases for Intercept
    @property
    def intercept(self):
        return self.Intercept

    @intercept.setter
    def intercept(self, value):
        self.Intercept = value

    @property
    def InterceptIsAuto(self):
        return self.com_object.InterceptIsAuto

    @InterceptIsAuto.setter
    def InterceptIsAuto(self, value):
        self.com_object.InterceptIsAuto = value

    # Lower case aliases for InterceptIsAuto
    @property
    def interceptisauto(self):
        return self.InterceptIsAuto

    @interceptisauto.setter
    def interceptisauto(self, value):
        self.InterceptIsAuto = value

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
    def NameIsAuto(self):
        return self.com_object.NameIsAuto

    @NameIsAuto.setter
    def NameIsAuto(self, value):
        self.com_object.NameIsAuto = value

    # Lower case aliases for NameIsAuto
    @property
    def nameisauto(self):
        return self.NameIsAuto

    @nameisauto.setter
    def nameisauto(self, value):
        self.NameIsAuto = value

    @property
    def Order(self):
        return self.com_object.Order

    @Order.setter
    def Order(self, value):
        self.com_object.Order = value

    # Lower case aliases for Order
    @property
    def order(self):
        return self.Order

    @order.setter
    def order(self, value):
        self.Order = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Period(self):
        return self.com_object.Period

    @Period.setter
    def Period(self, value):
        self.com_object.Period = value

    # Lower case aliases for Period
    @property
    def period(self):
        return self.Period

    @period.setter
    def period(self, value):
        self.Period = value

    @property
    def Type(self):
        return XlTrendlineType(self.com_object.Type)

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

    def ClearFormats(self):
        self.com_object.ClearFormats()

    # Lower case alias for ClearFormats
    def clearformats(self):
        return self.ClearFormats()

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class Trendlines:

    def __init__(self, trendlines=None):
        self.com_object= trendlines

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    def Add(self, Type=None, Order=None, Period=None, Forward=None, Backward=None, Intercept=None, DisplayEquation=None, DisplayRSquared=None, Name=None):
        arguments = com_arguments([unwrap(a) for a in [Type, Order, Period, Forward, Backward, Intercept, DisplayEquation, DisplayRSquared, Name]])
        return Trendline(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Type=None, Order=None, Period=None, Forward=None, Backward=None, Intercept=None, DisplayEquation=None, DisplayRSquared=None, Name=None):
        arguments = [Type, Order, Period, Forward, Backward, Intercept, DisplayEquation, DisplayRSquared, Name]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Trendline(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class UpBars:

    def __init__(self, upbars=None):
        self.com_object= upbars

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Fill(self):
        return FillFormat(self.com_object.Fill)

    # Lower case aliases for Fill
    @property
    def fill(self):
        return self.Fill

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

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

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


class View:

    def __init__(self, view=None):
        self.com_object= view

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def DisplaySlideMiniature(self):
        return self.com_object.DisplaySlideMiniature

    @DisplaySlideMiniature.setter
    def DisplaySlideMiniature(self, value):
        self.com_object.DisplaySlideMiniature = value

    # Lower case aliases for DisplaySlideMiniature
    @property
    def displayslideminiature(self):
        return self.DisplaySlideMiniature

    @displayslideminiature.setter
    def displayslideminiature(self, value):
        self.DisplaySlideMiniature = value

    @property
    def MediaControlsHeight(self):
        return self.com_object.MediaControlsHeight

    # Lower case aliases for MediaControlsHeight
    @property
    def mediacontrolsheight(self):
        return self.MediaControlsHeight

    @property
    def MediaControlsLeft(self):
        return self.com_object.MediaControlsLeft

    # Lower case aliases for MediaControlsLeft
    @property
    def mediacontrolsleft(self):
        return self.MediaControlsLeft

    @property
    def MediaControlsTop(self):
        return self.com_object.MediaControlsTop

    # Lower case aliases for MediaControlsTop
    @property
    def mediacontrolstop(self):
        return self.MediaControlsTop

    @property
    def MediaControlsVisible(self):
        return self.com_object.MediaControlsVisible

    # Lower case aliases for MediaControlsVisible
    @property
    def mediacontrolsvisible(self):
        return self.MediaControlsVisible

    @property
    def MediaControlsWidth(self):
        return self.com_object.MediaControlsWidth

    # Lower case aliases for MediaControlsWidth
    @property
    def mediacontrolswidth(self):
        return self.MediaControlsWidth

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PrintOptions(self):
        return PrintOptions(self.com_object.PrintOptions)

    # Lower case aliases for PrintOptions
    @property
    def printoptions(self):
        return self.PrintOptions

    @property
    def Slide(self):
        return Slide(self.com_object.Slide)

    @Slide.setter
    def Slide(self, value):
        self.com_object.Slide = value

    # Lower case aliases for Slide
    @property
    def slide(self):
        return self.Slide

    @slide.setter
    def slide(self, value):
        self.Slide = value

    @property
    def Type(self):
        return self.com_object.Type

    # Lower case aliases for Type
    @property
    def type(self):
        return self.Type

    @property
    def Zoom(self):
        return self.com_object.Zoom

    @Zoom.setter
    def Zoom(self, value):
        self.com_object.Zoom = value

    # Lower case aliases for Zoom
    @property
    def zoom(self):
        return self.Zoom

    @zoom.setter
    def zoom(self, value):
        self.Zoom = value

    @property
    def ZoomToFit(self):
        return self.com_object.ZoomToFit

    @ZoomToFit.setter
    def ZoomToFit(self, value):
        self.com_object.ZoomToFit = value

    # Lower case aliases for ZoomToFit
    @property
    def zoomtofit(self):
        return self.ZoomToFit

    @zoomtofit.setter
    def zoomtofit(self, value):
        self.ZoomToFit = value

    def GotoSlide(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        self.com_object.GotoSlide(*arguments)

    # Lower case alias for GotoSlide
    def gotoslide(self, Index=None):
        arguments = [Index]
        return self.GotoSlide(*arguments)

    def Paste(self):
        self.com_object.Paste()

    # Lower case alias for Paste
    def paste(self):
        return self.Paste()

    def PasteSpecial(self, DataType=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Link=None):
        arguments = com_arguments([unwrap(a) for a in [DataType, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Link]])
        self.com_object.PasteSpecial(*arguments)

    # Lower case alias for PasteSpecial
    def pastespecial(self, DataType=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None, Link=None):
        arguments = [DataType, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Link]
        return self.PasteSpecial(*arguments)

    def Player(self, ShapeId=None):
        arguments = com_arguments([unwrap(a) for a in [ShapeId]])
        return Player(self.com_object.Player(*arguments))

    # Lower case alias for Player
    def player(self, ShapeId=None):
        arguments = [ShapeId]
        return self.Player(*arguments)

    def PrintOut(self, From=None, To=None, PrintToFile=None, Copies=None, Collate=None):
        arguments = com_arguments([unwrap(a) for a in [From, To, PrintToFile, Copies, Collate]])
        self.com_object.PrintOut(*arguments)

    # Lower case alias for PrintOut
    def printout(self, From=None, To=None, PrintToFile=None, Copies=None, Collate=None):
        arguments = [From, To, PrintToFile, Copies, Collate]
        return self.PrintOut(*arguments)


class Walls:

    def __init__(self, walls=None):
        self.com_object= walls

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    # Lower case aliases for Format
    @property
    def format(self):
        return self.Format

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
    def PictureType(self):
        return self.com_object.PictureType

    @PictureType.setter
    def PictureType(self, value):
        self.com_object.PictureType = value

    # Lower case aliases for PictureType
    @property
    def picturetype(self):
        return self.PictureType

    @picturetype.setter
    def picturetype(self, value):
        self.PictureType = value

    @property
    def PictureUnit(self):
        return self.com_object.PictureUnit

    @PictureUnit.setter
    def PictureUnit(self, value):
        self.com_object.PictureUnit = value

    # Lower case aliases for PictureUnit
    @property
    def pictureunit(self):
        return self.PictureUnit

    @pictureunit.setter
    def pictureunit(self, value):
        self.PictureUnit = value

    @property
    def Thickness(self):
        return self.com_object.Thickness

    @Thickness.setter
    def Thickness(self, value):
        self.com_object.Thickness = value

    # Lower case aliases for Thickness
    @property
    def thickness(self):
        return self.Thickness

    @thickness.setter
    def thickness(self, value):
        self.Thickness = value

    def ClearFormats(self):
        self.com_object.ClearFormats()

    # Lower case alias for ClearFormats
    def clearformats(self):
        return self.ClearFormats()

    def Paste(self):
        self.com_object.Paste()

    # Lower case alias for Paste
    def paste(self):
        return self.Paste()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()


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
