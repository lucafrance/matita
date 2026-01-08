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

    @property
    def action(self):
        """Lower case alias for Action"""
        return self.Action

    @action.setter
    def action(self, value):
        """Lower case alias for Action.setter"""
        self.Action = value

    @property
    def ActionVerb(self):
        return self.com_object.ActionVerb

    @ActionVerb.setter
    def ActionVerb(self, value):
        self.com_object.ActionVerb = value

    @property
    def actionverb(self):
        """Lower case alias for ActionVerb"""
        return self.ActionVerb

    @actionverb.setter
    def actionverb(self, value):
        """Lower case alias for ActionVerb.setter"""
        self.ActionVerb = value

    @property
    def AnimateAction(self):
        return self.com_object.AnimateAction

    @AnimateAction.setter
    def AnimateAction(self, value):
        self.com_object.AnimateAction = value

    @property
    def animateaction(self):
        """Lower case alias for AnimateAction"""
        return self.AnimateAction

    @animateaction.setter
    def animateaction(self, value):
        """Lower case alias for AnimateAction.setter"""
        self.AnimateAction = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Hyperlink(self):
        return Hyperlink(self.com_object.Hyperlink)

    @property
    def hyperlink(self):
        """Lower case alias for Hyperlink"""
        return self.Hyperlink

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Run(self):
        return self.com_object.Run

    @Run.setter
    def Run(self, value):
        self.com_object.Run = value

    @property
    def run(self):
        """Lower case alias for Run"""
        return self.Run

    @run.setter
    def run(self, value):
        """Lower case alias for Run.setter"""
        self.Run = value

    @property
    def ShowAndReturn(self):
        return self.com_object.ShowAndReturn

    @ShowAndReturn.setter
    def ShowAndReturn(self, value):
        self.com_object.ShowAndReturn = value

    @property
    def showandreturn(self):
        """Lower case alias for ShowAndReturn"""
        return self.ShowAndReturn

    @showandreturn.setter
    def showandreturn(self, value):
        """Lower case alias for ShowAndReturn.setter"""
        self.ShowAndReturn = value

    @property
    def SlideShowName(self):
        return self.com_object.SlideShowName

    @SlideShowName.setter
    def SlideShowName(self, value):
        self.com_object.SlideShowName = value

    @property
    def slideshowname(self):
        """Lower case alias for SlideShowName"""
        return self.SlideShowName

    @slideshowname.setter
    def slideshowname(self, value):
        """Lower case alias for SlideShowName.setter"""
        self.SlideShowName = value

    @property
    def SoundEffect(self):
        return SoundEffect(self.com_object.SoundEffect)

    @property
    def soundeffect(self):
        """Lower case alias for SoundEffect"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def AutoLoad(self):
        return self.com_object.AutoLoad

    @AutoLoad.setter
    def AutoLoad(self, value):
        self.com_object.AutoLoad = value

    @property
    def autoload(self):
        """Lower case alias for AutoLoad"""
        return self.AutoLoad

    @autoload.setter
    def autoload(self, value):
        """Lower case alias for AutoLoad.setter"""
        self.AutoLoad = value

    @property
    def FullName(self):
        return self.com_object.FullName

    @property
    def fullname(self):
        """Lower case alias for FullName"""
        return self.FullName

    @property
    def Loaded(self):
        return self.com_object.Loaded

    @Loaded.setter
    def Loaded(self, value):
        self.com_object.Loaded = value

    @property
    def loaded(self):
        """Lower case alias for Loaded"""
        return self.Loaded

    @loaded.setter
    def loaded(self, value):
        """Lower case alias for Loaded.setter"""
        self.Loaded = value

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
    def Path(self):
        return AddIn(self.com_object.Path)

    @property
    def path(self):
        """Lower case alias for Path"""
        return self.Path

    @property
    def Registered(self):
        return self.com_object.Registered

    @Registered.setter
    def Registered(self, value):
        self.com_object.Registered = value

    @property
    def registered(self):
        """Lower case alias for Registered"""
        return self.Registered

    @registered.setter
    def registered(self, value):
        """Lower case alias for Registered.setter"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Count(self):
        return self.com_object.Count

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Item(self):
        return self.com_object.Item

    @Item.setter
    def Item(self, value):
        self.com_object.Item = value

    @property
    def item(self):
        """Lower case alias for Item"""
        return self.Item

    @item.setter
    def item(self, value):
        """Lower case alias for Item.setter"""
        self.Item = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
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

    @property
    def accumulate(self):
        """Lower case alias for Accumulate"""
        return self.Accumulate

    @accumulate.setter
    def accumulate(self, value):
        """Lower case alias for Accumulate.setter"""
        self.Accumulate = value

    @property
    def Additive(self):
        return self.com_object.Additive

    @Additive.setter
    def Additive(self, value):
        self.com_object.Additive = value

    @property
    def additive(self):
        """Lower case alias for Additive"""
        return self.Additive

    @additive.setter
    def additive(self, value):
        """Lower case alias for Additive.setter"""
        self.Additive = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def ColorEffect(self):
        return ColorEffect(self.com_object.ColorEffect)

    @property
    def coloreffect(self):
        """Lower case alias for ColorEffect"""
        return self.ColorEffect

    @property
    def CommandEffect(self):
        return CommandEffect(self.com_object.CommandEffect)

    @property
    def commandeffect(self):
        """Lower case alias for CommandEffect"""
        return self.CommandEffect

    @property
    def FilterEffect(self):
        return FilterEffect(self.com_object.FilterEffect)

    @property
    def filtereffect(self):
        """Lower case alias for FilterEffect"""
        return self.FilterEffect

    @property
    def MotionEffect(self):
        return MotionEffect(self.com_object.MotionEffect)

    @property
    def motioneffect(self):
        """Lower case alias for MotionEffect"""
        return self.MotionEffect

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PropertyEffect(self):
        return PropertyEffect(self.com_object.PropertyEffect)

    @property
    def propertyeffect(self):
        """Lower case alias for PropertyEffect"""
        return self.PropertyEffect

    @property
    def RotationEffect(self):
        return RotationEffect(self.com_object.RotationEffect)

    @property
    def rotationeffect(self):
        """Lower case alias for RotationEffect"""
        return self.RotationEffect

    @property
    def ScaleEffect(self):
        return ScaleEffect(self.com_object.ScaleEffect)

    @property
    def scaleeffect(self):
        """Lower case alias for ScaleEffect"""
        return self.ScaleEffect

    @property
    def SetEffect(self):
        return SetEffect(self.com_object.SetEffect)

    @property
    def seteffect(self):
        """Lower case alias for SetEffect"""
        return self.SetEffect

    @property
    def Timing(self):
        return Timing(self.com_object.Timing)

    @property
    def timing(self):
        """Lower case alias for Timing"""
        return self.Timing

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


class AnimationBehaviors:

    def __init__(self, animationbehaviors=None):
        self.com_object= animationbehaviors

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def Smooth(self):
        return self.com_object.Smooth

    @Smooth.setter
    def Smooth(self, value):
        self.com_object.Smooth = value

    @property
    def smooth(self):
        """Lower case alias for Smooth"""
        return self.Smooth

    @smooth.setter
    def smooth(self, value):
        """Lower case alias for Smooth.setter"""
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

    @property
    def advancemode(self):
        """Lower case alias for AdvanceMode"""
        return self.AdvanceMode

    @advancemode.setter
    def advancemode(self, value):
        """Lower case alias for AdvanceMode.setter"""
        self.AdvanceMode = value

    @property
    def AdvanceTime(self):
        return self.com_object.AdvanceTime

    @AdvanceTime.setter
    def AdvanceTime(self, value):
        self.com_object.AdvanceTime = value

    @property
    def advancetime(self):
        """Lower case alias for AdvanceTime"""
        return self.AdvanceTime

    @advancetime.setter
    def advancetime(self, value):
        """Lower case alias for AdvanceTime.setter"""
        self.AdvanceTime = value

    @property
    def AfterEffect(self):
        return PpAfterEffect(self.com_object.AfterEffect)

    @AfterEffect.setter
    def AfterEffect(self, value):
        self.com_object.AfterEffect = value

    @property
    def aftereffect(self):
        """Lower case alias for AfterEffect"""
        return self.AfterEffect

    @aftereffect.setter
    def aftereffect(self, value):
        """Lower case alias for AfterEffect.setter"""
        self.AfterEffect = value

    @property
    def Animate(self):
        return self.com_object.Animate

    @Animate.setter
    def Animate(self, value):
        self.com_object.Animate = value

    @property
    def animate(self):
        """Lower case alias for Animate"""
        return self.Animate

    @animate.setter
    def animate(self, value):
        """Lower case alias for Animate.setter"""
        self.Animate = value

    @property
    def AnimateBackground(self):
        return self.com_object.AnimateBackground

    @AnimateBackground.setter
    def AnimateBackground(self, value):
        self.com_object.AnimateBackground = value

    @property
    def animatebackground(self):
        """Lower case alias for AnimateBackground"""
        return self.AnimateBackground

    @animatebackground.setter
    def animatebackground(self, value):
        """Lower case alias for AnimateBackground.setter"""
        self.AnimateBackground = value

    @property
    def AnimateTextInReverse(self):
        return self.com_object.AnimateTextInReverse

    @AnimateTextInReverse.setter
    def AnimateTextInReverse(self, value):
        self.com_object.AnimateTextInReverse = value

    @property
    def animatetextinreverse(self):
        """Lower case alias for AnimateTextInReverse"""
        return self.AnimateTextInReverse

    @animatetextinreverse.setter
    def animatetextinreverse(self, value):
        """Lower case alias for AnimateTextInReverse.setter"""
        self.AnimateTextInReverse = value

    @property
    def AnimationOrder(self):
        return self.com_object.AnimationOrder

    @AnimationOrder.setter
    def AnimationOrder(self, value):
        self.com_object.AnimationOrder = value

    @property
    def animationorder(self):
        """Lower case alias for AnimationOrder"""
        return self.AnimationOrder

    @animationorder.setter
    def animationorder(self, value):
        """Lower case alias for AnimationOrder.setter"""
        self.AnimationOrder = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def ChartUnitEffect(self):
        return self.com_object.ChartUnitEffect

    @ChartUnitEffect.setter
    def ChartUnitEffect(self, value):
        self.com_object.ChartUnitEffect = value

    @property
    def chartuniteffect(self):
        """Lower case alias for ChartUnitEffect"""
        return self.ChartUnitEffect

    @chartuniteffect.setter
    def chartuniteffect(self, value):
        """Lower case alias for ChartUnitEffect.setter"""
        self.ChartUnitEffect = value

    @property
    def DimColor(self):
        return ColorFormat(self.com_object.DimColor)

    @DimColor.setter
    def DimColor(self, value):
        self.com_object.DimColor = value

    @property
    def dimcolor(self):
        """Lower case alias for DimColor"""
        return self.DimColor

    @dimcolor.setter
    def dimcolor(self, value):
        """Lower case alias for DimColor.setter"""
        self.DimColor = value

    @property
    def EntryEffect(self):
        return self.com_object.EntryEffect

    @EntryEffect.setter
    def EntryEffect(self, value):
        self.com_object.EntryEffect = value

    @property
    def entryeffect(self):
        """Lower case alias for EntryEffect"""
        return self.EntryEffect

    @entryeffect.setter
    def entryeffect(self, value):
        """Lower case alias for EntryEffect.setter"""
        self.EntryEffect = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PlaySettings(self):
        return PlaySettings(self.com_object.PlaySettings)

    @property
    def playsettings(self):
        """Lower case alias for PlaySettings"""
        return self.PlaySettings

    @property
    def SoundEffect(self):
        return SoundEffect(self.com_object.SoundEffect)

    @property
    def soundeffect(self):
        """Lower case alias for SoundEffect"""
        return self.SoundEffect

    @property
    def TextLevelEffect(self):
        return self.com_object.TextLevelEffect

    @TextLevelEffect.setter
    def TextLevelEffect(self, value):
        self.com_object.TextLevelEffect = value

    @property
    def textleveleffect(self):
        """Lower case alias for TextLevelEffect"""
        return self.TextLevelEffect

    @textleveleffect.setter
    def textleveleffect(self, value):
        """Lower case alias for TextLevelEffect.setter"""
        self.TextLevelEffect = value

    @property
    def TextUnitEffect(self):
        return self.com_object.TextUnitEffect

    @TextUnitEffect.setter
    def TextUnitEffect(self, value):
        self.com_object.TextUnitEffect = value

    @property
    def textuniteffect(self):
        """Lower case alias for TextUnitEffect"""
        return self.TextUnitEffect

    @textuniteffect.setter
    def textuniteffect(self, value):
        """Lower case alias for TextUnitEffect.setter"""
        self.TextUnitEffect = value


class Application:

    def __init__(self, application=None):
        if application is None:
            self.com_object = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
        else:
            self.com_object = application

    @property
    def Active(self):
        return self.com_object.Active

    @property
    def active(self):
        """Lower case alias for Active"""
        return self.Active

    @property
    def ActiveEncryptionSession(self):
        return self.com_object.ActiveEncryptionSession

    @property
    def activeencryptionsession(self):
        """Lower case alias for ActiveEncryptionSession"""
        return self.ActiveEncryptionSession

    @property
    def ActivePresentation(self):
        return Presentation(self.com_object.ActivePresentation)

    @property
    def activepresentation(self):
        """Lower case alias for ActivePresentation"""
        return self.ActivePresentation

    @property
    def ActivePrinter(self):
        return self.com_object.ActivePrinter

    @property
    def activeprinter(self):
        """Lower case alias for ActivePrinter"""
        return self.ActivePrinter

    @property
    def ActiveProtectedViewWindow(self):
        return ProtectedViewWindow(self.com_object.ActiveProtectedViewWindow)

    @property
    def activeprotectedviewwindow(self):
        """Lower case alias for ActiveProtectedViewWindow"""
        return self.ActiveProtectedViewWindow

    @property
    def ActiveWindow(self):
        return DocumentWindow(self.com_object.ActiveWindow)

    @property
    def activewindow(self):
        """Lower case alias for ActiveWindow"""
        return self.ActiveWindow

    @property
    def AddIns(self):
        return AddIns(self.com_object.AddIns)

    @property
    def addins(self):
        """Lower case alias for AddIns"""
        return self.AddIns

    @property
    def Assistance(self):
        return self.com_object.Assistance

    @property
    def assistance(self):
        """Lower case alias for Assistance"""
        return self.Assistance

    @property
    def AutoCorrect(self):
        return AutoCorrect(self.com_object.AutoCorrect)

    @property
    def autocorrect(self):
        """Lower case alias for AutoCorrect"""
        return self.AutoCorrect

    @property
    def AutomationSecurity(self):
        return self.com_object.AutomationSecurity

    @AutomationSecurity.setter
    def AutomationSecurity(self, value):
        self.com_object.AutomationSecurity = value

    @property
    def automationsecurity(self):
        """Lower case alias for AutomationSecurity"""
        return self.AutomationSecurity

    @automationsecurity.setter
    def automationsecurity(self, value):
        """Lower case alias for AutomationSecurity.setter"""
        self.AutomationSecurity = value

    @property
    def Build(self):
        return self.com_object.Build

    @property
    def build(self):
        """Lower case alias for Build"""
        return self.Build

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
    def chartdatapointtrack(self):
        return self.com_object.chartdatapointtrack

    @chartdatapointtrack.setter
    def chartdatapointtrack(self, value):
        self.com_object.chartdatapointtrack = value

    @property
    def COMAddIns(self):
        return self.com_object.COMAddIns

    @property
    def comaddins(self):
        """Lower case alias for COMAddIns"""
        return self.COMAddIns

    @property
    def CommandBars(self):
        return self.com_object.CommandBars

    @property
    def commandbars(self):
        """Lower case alias for CommandBars"""
        return self.CommandBars

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def DisplayAlerts(self):
        return self.com_object.DisplayAlerts

    @DisplayAlerts.setter
    def DisplayAlerts(self, value):
        self.com_object.DisplayAlerts = value

    @property
    def displayalerts(self):
        """Lower case alias for DisplayAlerts"""
        return self.DisplayAlerts

    @displayalerts.setter
    def displayalerts(self, value):
        """Lower case alias for DisplayAlerts.setter"""
        self.DisplayAlerts = value

    @property
    def DisplayDocumentInformationPanel(self):
        return self.com_object.DisplayDocumentInformationPanel

    @DisplayDocumentInformationPanel.setter
    def DisplayDocumentInformationPanel(self, value):
        self.com_object.DisplayDocumentInformationPanel = value

    @property
    def displaydocumentinformationpanel(self):
        """Lower case alias for DisplayDocumentInformationPanel"""
        return self.DisplayDocumentInformationPanel

    @displaydocumentinformationpanel.setter
    def displaydocumentinformationpanel(self, value):
        """Lower case alias for DisplayDocumentInformationPanel.setter"""
        self.DisplayDocumentInformationPanel = value

    @property
    def DisplayGridLines(self):
        return self.com_object.DisplayGridLines

    @DisplayGridLines.setter
    def DisplayGridLines(self, value):
        self.com_object.DisplayGridLines = value

    @property
    def displaygridlines(self):
        """Lower case alias for DisplayGridLines"""
        return self.DisplayGridLines

    @displaygridlines.setter
    def displaygridlines(self, value):
        """Lower case alias for DisplayGridLines.setter"""
        self.DisplayGridLines = value

    @property
    def displayguides(self):
        return self.com_object.displayguides

    @property
    def FeatureInstall(self):
        return self.com_object.FeatureInstall

    @FeatureInstall.setter
    def FeatureInstall(self, value):
        self.com_object.FeatureInstall = value

    @property
    def featureinstall(self):
        """Lower case alias for FeatureInstall"""
        return self.FeatureInstall

    @featureinstall.setter
    def featureinstall(self, value):
        """Lower case alias for FeatureInstall.setter"""
        self.FeatureInstall = value

    def FileConverters(self, Index1=None, Index2=None):
        arguments = com_arguments([unwrap(a) for a in [Index1, Index2]])
        if hasattr(self.com_object, "GetFileConverters"):
            return self.com_object.GetFileConverters(*arguments)
        else:
            return self.com_object.FileConverters(*arguments)

    def fileconverters(self, Index1=None, Index2=None):
        """Lower case alias for FileConverters"""
        arguments = [Index1, Index2]
        return self.FileConverters(*arguments)

    def FileDialog(self, Type=None):
        arguments = com_arguments([unwrap(a) for a in [Type]])
        if hasattr(self.com_object, "GetFileDialog"):
            return self.com_object.GetFileDialog(*arguments)
        else:
            return self.com_object.FileDialog(*arguments)

    def filedialog(self, Type=None):
        """Lower case alias for FileDialog"""
        arguments = [Type]
        return self.FileDialog(*arguments)

    @property
    def FileValidation(self):
        return self.com_object.FileValidation

    @FileValidation.setter
    def FileValidation(self, value):
        self.com_object.FileValidation = value

    @property
    def filevalidation(self):
        """Lower case alias for FileValidation"""
        return self.FileValidation

    @filevalidation.setter
    def filevalidation(self, value):
        """Lower case alias for FileValidation.setter"""
        self.FileValidation = value

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
    def IsSandboxed(self):
        return self.com_object.IsSandboxed

    @property
    def issandboxed(self):
        """Lower case alias for IsSandboxed"""
        return self.IsSandboxed

    @property
    def LanguageSettings(self):
        return self.com_object.LanguageSettings

    @property
    def languagesettings(self):
        """Lower case alias for LanguageSettings"""
        return self.LanguageSettings

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
    def Name(self):
        return self.com_object.Name

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def NewPresentation(self):
        return self.com_object.NewPresentation

    @property
    def newpresentation(self):
        """Lower case alias for NewPresentation"""
        return self.NewPresentation

    @property
    def OperatingSystem(self):
        return self.com_object.OperatingSystem

    @property
    def operatingsystem(self):
        """Lower case alias for OperatingSystem"""
        return self.OperatingSystem

    @property
    def Options(self):
        return Options(self.com_object.Options)

    @property
    def options(self):
        """Lower case alias for Options"""
        return self.Options

    @property
    def Path(self):
        return Application(self.com_object.Path)

    @property
    def path(self):
        """Lower case alias for Path"""
        return self.Path

    @property
    def Presentations(self):
        return Presentations(self.com_object.Presentations)

    @property
    def presentations(self):
        """Lower case alias for Presentations"""
        return self.Presentations

    @property
    def ProductCode(self):
        return self.com_object.ProductCode

    @property
    def productcode(self):
        """Lower case alias for ProductCode"""
        return self.ProductCode

    @property
    def ProtectedViewWindows(self):
        return ProtectedViewWindows(self.com_object.ProtectedViewWindows)

    @property
    def protectedviewwindows(self):
        """Lower case alias for ProtectedViewWindows"""
        return self.ProtectedViewWindows

    @property
    def SensitivityLabelPolicy(self):
        return self.com_object.SensitivityLabelPolicy

    @property
    def sensitivitylabelpolicy(self):
        """Lower case alias for SensitivityLabelPolicy"""
        return self.SensitivityLabelPolicy

    @property
    def ShowStartupDialog(self):
        return self.com_object.ShowStartupDialog

    @ShowStartupDialog.setter
    def ShowStartupDialog(self, value):
        self.com_object.ShowStartupDialog = value

    @property
    def showstartupdialog(self):
        """Lower case alias for ShowStartupDialog"""
        return self.ShowStartupDialog

    @showstartupdialog.setter
    def showstartupdialog(self, value):
        """Lower case alias for ShowStartupDialog.setter"""
        self.ShowStartupDialog = value

    @property
    def ShowWindowsInTaskbar(self):
        return self.com_object.ShowWindowsInTaskbar

    @ShowWindowsInTaskbar.setter
    def ShowWindowsInTaskbar(self, value):
        self.com_object.ShowWindowsInTaskbar = value

    @property
    def showwindowsintaskbar(self):
        """Lower case alias for ShowWindowsInTaskbar"""
        return self.ShowWindowsInTaskbar

    @showwindowsintaskbar.setter
    def showwindowsintaskbar(self, value):
        """Lower case alias for ShowWindowsInTaskbar.setter"""
        self.ShowWindowsInTaskbar = value

    @property
    def SlideShowWindows(self):
        return SlideShowWindows(self.com_object.SlideShowWindows)

    @property
    def slideshowwindows(self):
        """Lower case alias for SlideShowWindows"""
        return self.SlideShowWindows

    @property
    def SmartArtColors(self):
        return Application(self.com_object.SmartArtColors)

    @property
    def smartartcolors(self):
        """Lower case alias for SmartArtColors"""
        return self.SmartArtColors

    @property
    def SmartArtLayouts(self):
        return Application(self.com_object.SmartArtLayouts)

    @property
    def smartartlayouts(self):
        """Lower case alias for SmartArtLayouts"""
        return self.SmartArtLayouts

    @property
    def SmartArtQuickStyles(self):
        return Application(self.com_object.SmartArtQuickStyles)

    @property
    def smartartquickstyles(self):
        """Lower case alias for SmartArtQuickStyles"""
        return self.SmartArtQuickStyles

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
    def VBE(self):
        return self.com_object.VBE

    @property
    def vbe(self):
        """Lower case alias for VBE"""
        return self.VBE

    @property
    def Version(self):
        return self.com_object.Version

    @property
    def version(self):
        """Lower case alias for Version"""
        return self.Version

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
    def Windows(self):
        return DocumentWindows(self.com_object.Windows)

    @property
    def windows(self):
        """Lower case alias for Windows"""
        return self.Windows

    @property
    def WindowState(self):
        return self.com_object.WindowState

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

    @property
    def displayautocorrectoptions(self):
        """Lower case alias for DisplayAutoCorrectOptions"""
        return self.DisplayAutoCorrectOptions

    @displayautocorrectoptions.setter
    def displayautocorrectoptions(self, value):
        """Lower case alias for DisplayAutoCorrectOptions.setter"""
        self.DisplayAutoCorrectOptions = value

    @property
    def DisplayAutoLayoutOptions(self):
        return self.com_object.DisplayAutoLayoutOptions

    @DisplayAutoLayoutOptions.setter
    def DisplayAutoLayoutOptions(self, value):
        self.com_object.DisplayAutoLayoutOptions = value

    @property
    def displayautolayoutoptions(self):
        """Lower case alias for DisplayAutoLayoutOptions"""
        return self.DisplayAutoLayoutOptions

    @displayautolayoutoptions.setter
    def displayautolayoutoptions(self, value):
        """Lower case alias for DisplayAutoLayoutOptions.setter"""
        self.DisplayAutoLayoutOptions = value


class Axes:

    def __init__(self, axes=None):
        self.com_object= axes

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Count(self):
        return self.com_object.Count

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def AxisBetweenCategories(self):
        return self.com_object.AxisBetweenCategories

    @AxisBetweenCategories.setter
    def AxisBetweenCategories(self, value):
        self.com_object.AxisBetweenCategories = value

    @property
    def axisbetweencategories(self):
        """Lower case alias for AxisBetweenCategories"""
        return self.AxisBetweenCategories

    @axisbetweencategories.setter
    def axisbetweencategories(self, value):
        """Lower case alias for AxisBetweenCategories.setter"""
        self.AxisBetweenCategories = value

    @property
    def AxisGroup(self):
        return XlAxisGroup(self.com_object.AxisGroup)

    @property
    def axisgroup(self):
        """Lower case alias for AxisGroup"""
        return self.AxisGroup

    @property
    def AxisTitle(self):
        return AxisTitle(self.com_object.AxisTitle)

    @property
    def axistitle(self):
        """Lower case alias for AxisTitle"""
        return self.AxisTitle

    @property
    def BaseUnit(self):
        return XlTimeUnit(self.com_object.BaseUnit)

    @BaseUnit.setter
    def BaseUnit(self, value):
        self.com_object.BaseUnit = value

    @property
    def baseunit(self):
        """Lower case alias for BaseUnit"""
        return self.BaseUnit

    @baseunit.setter
    def baseunit(self, value):
        """Lower case alias for BaseUnit.setter"""
        self.BaseUnit = value

    @property
    def BaseUnitIsAuto(self):
        return self.com_object.BaseUnitIsAuto

    @BaseUnitIsAuto.setter
    def BaseUnitIsAuto(self, value):
        self.com_object.BaseUnitIsAuto = value

    @property
    def baseunitisauto(self):
        """Lower case alias for BaseUnitIsAuto"""
        return self.BaseUnitIsAuto

    @baseunitisauto.setter
    def baseunitisauto(self, value):
        """Lower case alias for BaseUnitIsAuto.setter"""
        self.BaseUnitIsAuto = value

    @property
    def Border(self):
        return ChartBorder(self.com_object.Border)

    @property
    def border(self):
        """Lower case alias for Border"""
        return self.Border

    @property
    def CategoryNames(self):
        return self.com_object.CategoryNames

    @CategoryNames.setter
    def CategoryNames(self, value):
        self.com_object.CategoryNames = value

    @property
    def categorynames(self):
        """Lower case alias for CategoryNames"""
        return self.CategoryNames

    @categorynames.setter
    def categorynames(self, value):
        """Lower case alias for CategoryNames.setter"""
        self.CategoryNames = value

    @property
    def CategoryType(self):
        return XlCategoryType(self.com_object.CategoryType)

    @CategoryType.setter
    def CategoryType(self, value):
        self.com_object.CategoryType = value

    @property
    def categorytype(self):
        """Lower case alias for CategoryType"""
        return self.CategoryType

    @categorytype.setter
    def categorytype(self, value):
        """Lower case alias for CategoryType.setter"""
        self.CategoryType = value

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Crosses(self):
        return self.com_object.Crosses

    @Crosses.setter
    def Crosses(self, value):
        self.com_object.Crosses = value

    @property
    def crosses(self):
        """Lower case alias for Crosses"""
        return self.Crosses

    @crosses.setter
    def crosses(self, value):
        """Lower case alias for Crosses.setter"""
        self.Crosses = value

    @property
    def CrossesAt(self):
        return self.com_object.CrossesAt

    @CrossesAt.setter
    def CrossesAt(self, value):
        self.com_object.CrossesAt = value

    @property
    def crossesat(self):
        """Lower case alias for CrossesAt"""
        return self.CrossesAt

    @crossesat.setter
    def crossesat(self, value):
        """Lower case alias for CrossesAt.setter"""
        self.CrossesAt = value

    @property
    def DisplayUnit(self):
        return XlDisplayUnit(self.com_object.DisplayUnit)

    @DisplayUnit.setter
    def DisplayUnit(self, value):
        self.com_object.DisplayUnit = value

    @property
    def displayunit(self):
        """Lower case alias for DisplayUnit"""
        return self.DisplayUnit

    @displayunit.setter
    def displayunit(self, value):
        """Lower case alias for DisplayUnit.setter"""
        self.DisplayUnit = value

    @property
    def DisplayUnitCustom(self):
        return self.com_object.DisplayUnitCustom

    @DisplayUnitCustom.setter
    def DisplayUnitCustom(self, value):
        self.com_object.DisplayUnitCustom = value

    @property
    def displayunitcustom(self):
        """Lower case alias for DisplayUnitCustom"""
        return self.DisplayUnitCustom

    @displayunitcustom.setter
    def displayunitcustom(self, value):
        """Lower case alias for DisplayUnitCustom.setter"""
        self.DisplayUnitCustom = value

    @property
    def DisplayUnitLabel(self):
        return DisplayUnitLabel(self.com_object.DisplayUnitLabel)

    @property
    def displayunitlabel(self):
        """Lower case alias for DisplayUnitLabel"""
        return self.DisplayUnitLabel

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

    @property
    def HasDisplayUnitLabel(self):
        return self.com_object.HasDisplayUnitLabel

    @HasDisplayUnitLabel.setter
    def HasDisplayUnitLabel(self, value):
        self.com_object.HasDisplayUnitLabel = value

    @property
    def hasdisplayunitlabel(self):
        """Lower case alias for HasDisplayUnitLabel"""
        return self.HasDisplayUnitLabel

    @hasdisplayunitlabel.setter
    def hasdisplayunitlabel(self, value):
        """Lower case alias for HasDisplayUnitLabel.setter"""
        self.HasDisplayUnitLabel = value

    @property
    def HasMajorGridlines(self):
        return self.com_object.HasMajorGridlines

    @HasMajorGridlines.setter
    def HasMajorGridlines(self, value):
        self.com_object.HasMajorGridlines = value

    @property
    def hasmajorgridlines(self):
        """Lower case alias for HasMajorGridlines"""
        return self.HasMajorGridlines

    @hasmajorgridlines.setter
    def hasmajorgridlines(self, value):
        """Lower case alias for HasMajorGridlines.setter"""
        self.HasMajorGridlines = value

    @property
    def HasMinorGridlines(self):
        return self.com_object.HasMinorGridlines

    @HasMinorGridlines.setter
    def HasMinorGridlines(self, value):
        self.com_object.HasMinorGridlines = value

    @property
    def hasminorgridlines(self):
        """Lower case alias for HasMinorGridlines"""
        return self.HasMinorGridlines

    @hasminorgridlines.setter
    def hasminorgridlines(self, value):
        """Lower case alias for HasMinorGridlines.setter"""
        self.HasMinorGridlines = value

    @property
    def HasTitle(self):
        return self.com_object.HasTitle

    @HasTitle.setter
    def HasTitle(self, value):
        self.com_object.HasTitle = value

    @property
    def hastitle(self):
        """Lower case alias for HasTitle"""
        return self.HasTitle

    @hastitle.setter
    def hastitle(self, value):
        """Lower case alias for HasTitle.setter"""
        self.HasTitle = value

    @property
    def Height(self):
        return self.com_object.Height

    @property
    def height(self):
        """Lower case alias for Height"""
        return self.Height

    @property
    def Left(self):
        return self.com_object.Left

    @property
    def left(self):
        """Lower case alias for Left"""
        return self.Left

    @property
    def LogBase(self):
        return self.com_object.LogBase

    @LogBase.setter
    def LogBase(self, value):
        self.com_object.LogBase = value

    @property
    def logbase(self):
        """Lower case alias for LogBase"""
        return self.LogBase

    @logbase.setter
    def logbase(self, value):
        """Lower case alias for LogBase.setter"""
        self.LogBase = value

    @property
    def MajorGridlines(self):
        return Gridlines(self.com_object.MajorGridlines)

    @property
    def majorgridlines(self):
        """Lower case alias for MajorGridlines"""
        return self.MajorGridlines

    @property
    def MajorTickMark(self):
        return XlTickMark(self.com_object.MajorTickMark)

    @MajorTickMark.setter
    def MajorTickMark(self, value):
        self.com_object.MajorTickMark = value

    @property
    def majortickmark(self):
        """Lower case alias for MajorTickMark"""
        return self.MajorTickMark

    @majortickmark.setter
    def majortickmark(self, value):
        """Lower case alias for MajorTickMark.setter"""
        self.MajorTickMark = value

    @property
    def MajorUnit(self):
        return self.com_object.MajorUnit

    @MajorUnit.setter
    def MajorUnit(self, value):
        self.com_object.MajorUnit = value

    @property
    def majorunit(self):
        """Lower case alias for MajorUnit"""
        return self.MajorUnit

    @majorunit.setter
    def majorunit(self, value):
        """Lower case alias for MajorUnit.setter"""
        self.MajorUnit = value

    @property
    def MajorUnitIsAuto(self):
        return self.com_object.MajorUnitIsAuto

    @MajorUnitIsAuto.setter
    def MajorUnitIsAuto(self, value):
        self.com_object.MajorUnitIsAuto = value

    @property
    def majorunitisauto(self):
        """Lower case alias for MajorUnitIsAuto"""
        return self.MajorUnitIsAuto

    @majorunitisauto.setter
    def majorunitisauto(self, value):
        """Lower case alias for MajorUnitIsAuto.setter"""
        self.MajorUnitIsAuto = value

    @property
    def MajorUnitScale(self):
        return self.com_object.MajorUnitScale

    @MajorUnitScale.setter
    def MajorUnitScale(self, value):
        self.com_object.MajorUnitScale = value

    @property
    def majorunitscale(self):
        """Lower case alias for MajorUnitScale"""
        return self.MajorUnitScale

    @majorunitscale.setter
    def majorunitscale(self, value):
        """Lower case alias for MajorUnitScale.setter"""
        self.MajorUnitScale = value

    @property
    def MaximumScale(self):
        return self.com_object.MaximumScale

    @MaximumScale.setter
    def MaximumScale(self, value):
        self.com_object.MaximumScale = value

    @property
    def maximumscale(self):
        """Lower case alias for MaximumScale"""
        return self.MaximumScale

    @maximumscale.setter
    def maximumscale(self, value):
        """Lower case alias for MaximumScale.setter"""
        self.MaximumScale = value

    @property
    def MaximumScaleIsAuto(self):
        return self.com_object.MaximumScaleIsAuto

    @MaximumScaleIsAuto.setter
    def MaximumScaleIsAuto(self, value):
        self.com_object.MaximumScaleIsAuto = value

    @property
    def maximumscaleisauto(self):
        """Lower case alias for MaximumScaleIsAuto"""
        return self.MaximumScaleIsAuto

    @maximumscaleisauto.setter
    def maximumscaleisauto(self, value):
        """Lower case alias for MaximumScaleIsAuto.setter"""
        self.MaximumScaleIsAuto = value

    @property
    def MinimumScale(self):
        return self.com_object.MinimumScale

    @MinimumScale.setter
    def MinimumScale(self, value):
        self.com_object.MinimumScale = value

    @property
    def minimumscale(self):
        """Lower case alias for MinimumScale"""
        return self.MinimumScale

    @minimumscale.setter
    def minimumscale(self, value):
        """Lower case alias for MinimumScale.setter"""
        self.MinimumScale = value

    @property
    def MinimumScaleIsAuto(self):
        return self.com_object.MinimumScaleIsAuto

    @MinimumScaleIsAuto.setter
    def MinimumScaleIsAuto(self, value):
        self.com_object.MinimumScaleIsAuto = value

    @property
    def minimumscaleisauto(self):
        """Lower case alias for MinimumScaleIsAuto"""
        return self.MinimumScaleIsAuto

    @minimumscaleisauto.setter
    def minimumscaleisauto(self, value):
        """Lower case alias for MinimumScaleIsAuto.setter"""
        self.MinimumScaleIsAuto = value

    @property
    def MinorGridlines(self):
        return Gridlines(self.com_object.MinorGridlines)

    @property
    def minorgridlines(self):
        """Lower case alias for MinorGridlines"""
        return self.MinorGridlines

    @property
    def MinorTickMark(self):
        return XlTickMark(self.com_object.MinorTickMark)

    @MinorTickMark.setter
    def MinorTickMark(self, value):
        self.com_object.MinorTickMark = value

    @property
    def minortickmark(self):
        """Lower case alias for MinorTickMark"""
        return self.MinorTickMark

    @minortickmark.setter
    def minortickmark(self, value):
        """Lower case alias for MinorTickMark.setter"""
        self.MinorTickMark = value

    @property
    def MinorUnit(self):
        return self.com_object.MinorUnit

    @MinorUnit.setter
    def MinorUnit(self, value):
        self.com_object.MinorUnit = value

    @property
    def minorunit(self):
        """Lower case alias for MinorUnit"""
        return self.MinorUnit

    @minorunit.setter
    def minorunit(self, value):
        """Lower case alias for MinorUnit.setter"""
        self.MinorUnit = value

    @property
    def MinorUnitIsAuto(self):
        return self.com_object.MinorUnitIsAuto

    @MinorUnitIsAuto.setter
    def MinorUnitIsAuto(self, value):
        self.com_object.MinorUnitIsAuto = value

    @property
    def minorunitisauto(self):
        """Lower case alias for MinorUnitIsAuto"""
        return self.MinorUnitIsAuto

    @minorunitisauto.setter
    def minorunitisauto(self, value):
        """Lower case alias for MinorUnitIsAuto.setter"""
        self.MinorUnitIsAuto = value

    @property
    def MinorUnitScale(self):
        return self.com_object.MinorUnitScale

    @MinorUnitScale.setter
    def MinorUnitScale(self, value):
        self.com_object.MinorUnitScale = value

    @property
    def minorunitscale(self):
        """Lower case alias for MinorUnitScale"""
        return self.MinorUnitScale

    @minorunitscale.setter
    def minorunitscale(self, value):
        """Lower case alias for MinorUnitScale.setter"""
        self.MinorUnitScale = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def ReversePlotOrder(self):
        return self.com_object.ReversePlotOrder

    @ReversePlotOrder.setter
    def ReversePlotOrder(self, value):
        self.com_object.ReversePlotOrder = value

    @property
    def reverseplotorder(self):
        """Lower case alias for ReversePlotOrder"""
        return self.ReversePlotOrder

    @reverseplotorder.setter
    def reverseplotorder(self, value):
        """Lower case alias for ReversePlotOrder.setter"""
        self.ReversePlotOrder = value

    @property
    def ScaleType(self):
        return XlScaleType(self.com_object.ScaleType)

    @ScaleType.setter
    def ScaleType(self, value):
        self.com_object.ScaleType = value

    @property
    def scaletype(self):
        """Lower case alias for ScaleType"""
        return self.ScaleType

    @scaletype.setter
    def scaletype(self, value):
        """Lower case alias for ScaleType.setter"""
        self.ScaleType = value

    @property
    def TickLabelPosition(self):
        return self.com_object.TickLabelPosition

    @TickLabelPosition.setter
    def TickLabelPosition(self, value):
        self.com_object.TickLabelPosition = value

    @property
    def ticklabelposition(self):
        """Lower case alias for TickLabelPosition"""
        return self.TickLabelPosition

    @ticklabelposition.setter
    def ticklabelposition(self, value):
        """Lower case alias for TickLabelPosition.setter"""
        self.TickLabelPosition = value

    @property
    def TickLabels(self):
        return TickLabels(self.com_object.TickLabels)

    @property
    def ticklabels(self):
        """Lower case alias for TickLabels"""
        return self.TickLabels

    @property
    def TickLabelSpacing(self):
        return self.com_object.TickLabelSpacing

    @TickLabelSpacing.setter
    def TickLabelSpacing(self, value):
        self.com_object.TickLabelSpacing = value

    @property
    def ticklabelspacing(self):
        """Lower case alias for TickLabelSpacing"""
        return self.TickLabelSpacing

    @ticklabelspacing.setter
    def ticklabelspacing(self, value):
        """Lower case alias for TickLabelSpacing.setter"""
        self.TickLabelSpacing = value

    @property
    def TickLabelSpacingIsAuto(self):
        return self.com_object.TickLabelSpacingIsAuto

    @TickLabelSpacingIsAuto.setter
    def TickLabelSpacingIsAuto(self, value):
        self.com_object.TickLabelSpacingIsAuto = value

    @property
    def ticklabelspacingisauto(self):
        """Lower case alias for TickLabelSpacingIsAuto"""
        return self.TickLabelSpacingIsAuto

    @ticklabelspacingisauto.setter
    def ticklabelspacingisauto(self, value):
        """Lower case alias for TickLabelSpacingIsAuto.setter"""
        self.TickLabelSpacingIsAuto = value

    @property
    def TickMarkSpacing(self):
        return self.com_object.TickMarkSpacing

    @TickMarkSpacing.setter
    def TickMarkSpacing(self, value):
        self.com_object.TickMarkSpacing = value

    @property
    def tickmarkspacing(self):
        """Lower case alias for TickMarkSpacing"""
        return self.TickMarkSpacing

    @tickmarkspacing.setter
    def tickmarkspacing(self, value):
        """Lower case alias for TickMarkSpacing.setter"""
        self.TickMarkSpacing = value

    @property
    def Top(self):
        return self.com_object.Top

    @property
    def top(self):
        """Lower case alias for Top"""
        return self.Top

    @property
    def Type(self):
        return XlAxisType(self.com_object.Type)

    @property
    def type(self):
        """Lower case alias for Type"""
        return self.Type

    @property
    def Width(self):
        return self.com_object.Width

    @property
    def width(self):
        """Lower case alias for Width"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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

    def Characters(self, Start=None, Length=None):
        arguments = com_arguments([unwrap(a) for a in [Start, Length]])
        if hasattr(self.com_object, "GetCharacters"):
            return ChartCharacters(self.com_object.GetCharacters(*arguments))
        else:
            return ChartCharacters(self.com_object.Characters(*arguments))

    def characters(self, Start=None, Length=None):
        """Lower case alias for Characters"""
        arguments = [Start, Length]
        return self.Characters(*arguments)

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

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
    def FormulaLocal(self):
        return self.com_object.FormulaLocal

    @FormulaLocal.setter
    def FormulaLocal(self, value):
        self.com_object.FormulaLocal = value

    @property
    def formulalocal(self):
        """Lower case alias for FormulaLocal"""
        return self.FormulaLocal

    @formulalocal.setter
    def formulalocal(self, value):
        """Lower case alias for FormulaLocal.setter"""
        self.FormulaLocal = value

    @property
    def FormulaR1C1(self):
        return self.com_object.FormulaR1C1

    @FormulaR1C1.setter
    def FormulaR1C1(self, value):
        self.com_object.FormulaR1C1 = value

    @property
    def formular1c1(self):
        """Lower case alias for FormulaR1C1"""
        return self.FormulaR1C1

    @formular1c1.setter
    def formular1c1(self, value):
        """Lower case alias for FormulaR1C1.setter"""
        self.FormulaR1C1 = value

    @property
    def FormulaR1C1Local(self):
        return self.com_object.FormulaR1C1Local

    @FormulaR1C1Local.setter
    def FormulaR1C1Local(self, value):
        self.com_object.FormulaR1C1Local = value

    @property
    def formular1c1local(self):
        """Lower case alias for FormulaR1C1Local"""
        return self.FormulaR1C1Local

    @formular1c1local.setter
    def formular1c1local(self, value):
        """Lower case alias for FormulaR1C1Local.setter"""
        self.FormulaR1C1Local = value

    @property
    def Height(self):
        return self.com_object.Height

    @property
    def height(self):
        """Lower case alias for Height"""
        return self.Height

    @property
    def HorizontalAlignment(self):
        return self.com_object.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.com_object.HorizontalAlignment = value

    @property
    def horizontalalignment(self):
        """Lower case alias for HorizontalAlignment"""
        return self.HorizontalAlignment

    @horizontalalignment.setter
    def horizontalalignment(self, value):
        """Lower case alias for HorizontalAlignment.setter"""
        self.HorizontalAlignment = value

    @property
    def IncludeInLayout(self):
        return self.com_object.IncludeInLayout

    @IncludeInLayout.setter
    def IncludeInLayout(self, value):
        self.com_object.IncludeInLayout = value

    @property
    def includeinlayout(self):
        """Lower case alias for IncludeInLayout"""
        return self.IncludeInLayout

    @includeinlayout.setter
    def includeinlayout(self, value):
        """Lower case alias for IncludeInLayout.setter"""
        self.IncludeInLayout = value

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
    def Name(self):
        return self.com_object.Name

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def Orientation(self):
        return self.com_object.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.com_object.Orientation = value

    @property
    def orientation(self):
        """Lower case alias for Orientation"""
        return self.Orientation

    @orientation.setter
    def orientation(self, value):
        """Lower case alias for Orientation.setter"""
        self.Orientation = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Position(self):
        return XlChartElementPosition(self.com_object.Position)

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
    def ReadingOrder(self):
        return XlReadingOrder(self.com_object.ReadingOrder)

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.com_object.ReadingOrder = value

    @property
    def readingorder(self):
        """Lower case alias for ReadingOrder"""
        return self.ReadingOrder

    @readingorder.setter
    def readingorder(self, value):
        """Lower case alias for ReadingOrder.setter"""
        self.ReadingOrder = value

    @property
    def Shadow(self):
        return self.com_object.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.com_object.Shadow = value

    @property
    def shadow(self):
        """Lower case alias for Shadow"""
        return self.Shadow

    @shadow.setter
    def shadow(self, value):
        """Lower case alias for Shadow.setter"""
        self.Shadow = value

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
    def VerticalAlignment(self):
        return self.com_object.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.com_object.VerticalAlignment = value

    @property
    def verticalalignment(self):
        """Lower case alias for VerticalAlignment"""
        return self.VerticalAlignment

    @verticalalignment.setter
    def verticalalignment(self, value):
        """Lower case alias for VerticalAlignment.setter"""
        self.VerticalAlignment = value

    @property
    def Width(self):
        return self.com_object.Width

    @property
    def width(self):
        """Lower case alias for Width"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def AttendeeUrl(self):
        return self.com_object.AttendeeUrl

    @property
    def attendeeurl(self):
        """Lower case alias for AttendeeUrl"""
        return self.AttendeeUrl

    @property
    def capabilities(self):
        return self.com_object.capabilities

    @property
    def IsBroadcasting(self):
        return self.com_object.IsBroadcasting

    @property
    def isbroadcasting(self):
        """Lower case alias for IsBroadcasting"""
        return self.IsBroadcasting

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def presenterserviceurl(self):
        return self.com_object.presenterserviceurl

    @property
    def sessionid(self):
        return self.com_object.sessionid

    @property
    def state(self):
        return self.com_object.state

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Character(self):
        return self.com_object.Character

    @Character.setter
    def Character(self, value):
        self.com_object.Character = value

    @property
    def character(self):
        """Lower case alias for Character"""
        return self.Character

    @character.setter
    def character(self, value):
        """Lower case alias for Character.setter"""
        self.Character = value

    @property
    def Font(self):
        return Font(self.com_object.Font)

    @property
    def font(self):
        """Lower case alias for Font"""
        return self.Font

    @property
    def Number(self):
        return self.com_object.Number

    @property
    def number(self):
        """Lower case alias for Number"""
        return self.Number

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def RelativeSize(self):
        return self.com_object.RelativeSize

    @RelativeSize.setter
    def RelativeSize(self, value):
        self.com_object.RelativeSize = value

    @property
    def relativesize(self):
        """Lower case alias for RelativeSize"""
        return self.RelativeSize

    @relativesize.setter
    def relativesize(self, value):
        """Lower case alias for RelativeSize.setter"""
        self.RelativeSize = value

    @property
    def StartValue(self):
        return self.com_object.StartValue

    @StartValue.setter
    def StartValue(self, value):
        self.com_object.StartValue = value

    @property
    def startvalue(self):
        """Lower case alias for StartValue"""
        return self.StartValue

    @startvalue.setter
    def startvalue(self, value):
        """Lower case alias for StartValue.setter"""
        self.StartValue = value

    @property
    def Style(self):
        return self.com_object.Style

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
    def UseTextColor(self):
        return self.com_object.UseTextColor

    @UseTextColor.setter
    def UseTextColor(self, value):
        self.com_object.UseTextColor = value

    @property
    def usetextcolor(self):
        """Lower case alias for UseTextColor"""
        return self.UseTextColor

    @usetextcolor.setter
    def usetextcolor(self, value):
        """Lower case alias for UseTextColor.setter"""
        self.UseTextColor = value

    @property
    def UseTextFont(self):
        return self.com_object.UseTextFont

    @UseTextFont.setter
    def UseTextFont(self, value):
        self.com_object.UseTextFont = value

    @property
    def usetextfont(self):
        """Lower case alias for UseTextFont"""
        return self.UseTextFont

    @usetextfont.setter
    def usetextfont(self, value):
        """Lower case alias for UseTextFont.setter"""
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

    @property
    def accent(self):
        """Lower case alias for Accent"""
        return self.Accent

    @accent.setter
    def accent(self, value):
        """Lower case alias for Accent.setter"""
        self.Accent = value

    @property
    def Angle(self):
        return self.com_object.Angle

    @Angle.setter
    def Angle(self, value):
        self.com_object.Angle = value

    @property
    def angle(self):
        """Lower case alias for Angle"""
        return self.Angle

    @angle.setter
    def angle(self, value):
        """Lower case alias for Angle.setter"""
        self.Angle = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def AutoAttach(self):
        return self.com_object.AutoAttach

    @AutoAttach.setter
    def AutoAttach(self, value):
        self.com_object.AutoAttach = value

    @property
    def autoattach(self):
        """Lower case alias for AutoAttach"""
        return self.AutoAttach

    @autoattach.setter
    def autoattach(self, value):
        """Lower case alias for AutoAttach.setter"""
        self.AutoAttach = value

    @property
    def AutoLength(self):
        return self.com_object.AutoLength

    @property
    def autolength(self):
        """Lower case alias for AutoLength"""
        return self.AutoLength

    @property
    def Border(self):
        return self.com_object.Border

    @Border.setter
    def Border(self, value):
        self.com_object.Border = value

    @property
    def border(self):
        """Lower case alias for Border"""
        return self.Border

    @border.setter
    def border(self, value):
        """Lower case alias for Border.setter"""
        self.Border = value

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Drop(self):
        return self.com_object.Drop

    @property
    def drop(self):
        """Lower case alias for Drop"""
        return self.Drop

    @property
    def DropType(self):
        return self.com_object.DropType

    @property
    def droptype(self):
        """Lower case alias for DropType"""
        return self.DropType

    @property
    def Gap(self):
        return self.com_object.Gap

    @Gap.setter
    def Gap(self, value):
        self.com_object.Gap = value

    @property
    def gap(self):
        """Lower case alias for Gap"""
        return self.Gap

    @gap.setter
    def gap(self, value):
        """Lower case alias for Gap.setter"""
        self.Gap = value

    @property
    def Length(self):
        return self.com_object.Length

    @property
    def length(self):
        """Lower case alias for Length"""
        return self.Length

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

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


class categorycollection:

    def __init__(self, categorycollection=None):
        self.com_object= categorycollection

    @property
    def application(self):
        return self.com_object.application

    @property
    def count(self):
        return self.com_object.count

    @property
    def creator(self):
        return self.com_object.creator

    @property
    def parent(self):
        return self.com_object.parent


class Cell:

    def __init__(self, cell=None):
        self.com_object= cell

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Borders(self):
        return Borders(self.com_object.Borders)

    @property
    def borders(self):
        """Lower case alias for Borders"""
        return self.Borders

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Selected(self):
        return self.com_object.Selected

    @property
    def selected(self):
        """Lower case alias for Selected"""
        return self.Selected

    @property
    def Shape(self):
        return Shape(self.com_object.Shape)

    @property
    def shape(self):
        """Lower case alias for Shape"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Borders(self):
        return Borders(self.com_object.Borders)

    @property
    def borders(self):
        """Lower case alias for Borders"""
        return self.Borders

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

    @property
    def alternativetext(self):
        """Lower case alias for AlternativeText"""
        return self.AlternativeText

    @alternativetext.setter
    def alternativetext(self, value):
        """Lower case alias for AlternativeText.setter"""
        self.AlternativeText = value

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def AutoScaling(self):
        return self.com_object.AutoScaling

    @AutoScaling.setter
    def AutoScaling(self, value):
        self.com_object.AutoScaling = value

    @property
    def autoscaling(self):
        """Lower case alias for AutoScaling"""
        return self.AutoScaling

    @autoscaling.setter
    def autoscaling(self, value):
        """Lower case alias for AutoScaling.setter"""
        self.AutoScaling = value

    @property
    def BackWall(self):
        return Walls(self.com_object.BackWall)

    @property
    def backwall(self):
        """Lower case alias for BackWall"""
        return self.BackWall

    @property
    def BarShape(self):
        return XlBarShape(self.com_object.BarShape)

    @BarShape.setter
    def BarShape(self, value):
        self.com_object.BarShape = value

    @property
    def barshape(self):
        """Lower case alias for BarShape"""
        return self.BarShape

    @barshape.setter
    def barshape(self, value):
        """Lower case alias for BarShape.setter"""
        self.BarShape = value

    @property
    def categorylabellevel(self):
        return self.com_object.categorylabellevel

    @categorylabellevel.setter
    def categorylabellevel(self, value):
        self.com_object.categorylabellevel = value

    @property
    def ChartArea(self):
        return ChartArea(self.com_object.ChartArea)

    @property
    def chartarea(self):
        """Lower case alias for ChartArea"""
        return self.ChartArea

    @property
    def chartcolor(self):
        return self.com_object.chartcolor

    @chartcolor.setter
    def chartcolor(self, value):
        self.com_object.chartcolor = value

    @property
    def ChartData(self):
        return ChartData(self.com_object.ChartData)

    @property
    def chartdata(self):
        """Lower case alias for ChartData"""
        return self.ChartData

    @property
    def ChartStyle(self):
        return self.com_object.ChartStyle

    @ChartStyle.setter
    def ChartStyle(self, value):
        self.com_object.ChartStyle = value

    @property
    def chartstyle(self):
        """Lower case alias for ChartStyle"""
        return self.ChartStyle

    @chartstyle.setter
    def chartstyle(self, value):
        """Lower case alias for ChartStyle.setter"""
        self.ChartStyle = value

    @property
    def ChartTitle(self):
        return ChartTitle(self.com_object.ChartTitle)

    @property
    def charttitle(self):
        """Lower case alias for ChartTitle"""
        return self.ChartTitle

    @property
    def ChartType(self):
        return self.com_object.ChartType

    @ChartType.setter
    def ChartType(self, value):
        self.com_object.ChartType = value

    @property
    def charttype(self):
        """Lower case alias for ChartType"""
        return self.ChartType

    @charttype.setter
    def charttype(self, value):
        """Lower case alias for ChartType.setter"""
        self.ChartType = value

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def DataTable(self):
        return DataTable(self.com_object.DataTable)

    @property
    def datatable(self):
        """Lower case alias for DataTable"""
        return self.DataTable

    @property
    def DepthPercent(self):
        return self.com_object.DepthPercent

    @DepthPercent.setter
    def DepthPercent(self, value):
        self.com_object.DepthPercent = value

    @property
    def depthpercent(self):
        """Lower case alias for DepthPercent"""
        return self.DepthPercent

    @depthpercent.setter
    def depthpercent(self, value):
        """Lower case alias for DepthPercent.setter"""
        self.DepthPercent = value

    @property
    def DisplayBlanksAs(self):
        return XlDisplayBlanksAs(self.com_object.DisplayBlanksAs)

    @DisplayBlanksAs.setter
    def DisplayBlanksAs(self, value):
        self.com_object.DisplayBlanksAs = value

    @property
    def displayblanksas(self):
        """Lower case alias for DisplayBlanksAs"""
        return self.DisplayBlanksAs

    @displayblanksas.setter
    def displayblanksas(self, value):
        """Lower case alias for DisplayBlanksAs.setter"""
        self.DisplayBlanksAs = value

    @property
    def Elevation(self):
        return self.com_object.Elevation

    @Elevation.setter
    def Elevation(self, value):
        self.com_object.Elevation = value

    @property
    def elevation(self):
        """Lower case alias for Elevation"""
        return self.Elevation

    @elevation.setter
    def elevation(self, value):
        """Lower case alias for Elevation.setter"""
        self.Elevation = value

    @property
    def Floor(self):
        return Floor(self.com_object.Floor)

    @property
    def floor(self):
        """Lower case alias for Floor"""
        return self.Floor

    @property
    def Format(self):
        return self.com_object.Format

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

    @property
    def GapDepth(self):
        return self.com_object.GapDepth

    @GapDepth.setter
    def GapDepth(self, value):
        self.com_object.GapDepth = value

    @property
    def gapdepth(self):
        """Lower case alias for GapDepth"""
        return self.GapDepth

    @gapdepth.setter
    def gapdepth(self, value):
        """Lower case alias for GapDepth.setter"""
        self.GapDepth = value

    @property
    def HasAxis(self):
        return self.com_object.HasAxis

    @HasAxis.setter
    def HasAxis(self, value):
        self.com_object.HasAxis = value

    @property
    def hasaxis(self):
        """Lower case alias for HasAxis"""
        return self.HasAxis

    @hasaxis.setter
    def hasaxis(self, value):
        """Lower case alias for HasAxis.setter"""
        self.HasAxis = value

    @property
    def HasDataTable(self):
        return self.com_object.HasDataTable

    @HasDataTable.setter
    def HasDataTable(self, value):
        self.com_object.HasDataTable = value

    @property
    def hasdatatable(self):
        """Lower case alias for HasDataTable"""
        return self.HasDataTable

    @hasdatatable.setter
    def hasdatatable(self, value):
        """Lower case alias for HasDataTable.setter"""
        self.HasDataTable = value

    @property
    def HasLegend(self):
        return self.com_object.HasLegend

    @HasLegend.setter
    def HasLegend(self, value):
        self.com_object.HasLegend = value

    @property
    def haslegend(self):
        """Lower case alias for HasLegend"""
        return self.HasLegend

    @haslegend.setter
    def haslegend(self, value):
        """Lower case alias for HasLegend.setter"""
        self.HasLegend = value

    @property
    def HasTitle(self):
        return self.com_object.HasTitle

    @HasTitle.setter
    def HasTitle(self, value):
        self.com_object.HasTitle = value

    @property
    def hastitle(self):
        """Lower case alias for HasTitle"""
        return self.HasTitle

    @hastitle.setter
    def hastitle(self, value):
        """Lower case alias for HasTitle.setter"""
        self.HasTitle = value

    @property
    def HeightPercent(self):
        return self.com_object.HeightPercent

    @HeightPercent.setter
    def HeightPercent(self, value):
        self.com_object.HeightPercent = value

    @property
    def heightpercent(self):
        """Lower case alias for HeightPercent"""
        return self.HeightPercent

    @heightpercent.setter
    def heightpercent(self, value):
        """Lower case alias for HeightPercent.setter"""
        self.HeightPercent = value

    @property
    def Legend(self):
        return Legend(self.com_object.Legend)

    @property
    def legend(self):
        """Lower case alias for Legend"""
        return self.Legend

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
    def Perspective(self):
        return self.com_object.Perspective

    @Perspective.setter
    def Perspective(self, value):
        self.com_object.Perspective = value

    @property
    def perspective(self):
        """Lower case alias for Perspective"""
        return self.Perspective

    @perspective.setter
    def perspective(self, value):
        """Lower case alias for Perspective.setter"""
        self.Perspective = value

    @property
    def PlotArea(self):
        return PlotArea(self.com_object.PlotArea)

    @property
    def plotarea(self):
        """Lower case alias for PlotArea"""
        return self.PlotArea

    @property
    def PlotBy(self):
        return self.com_object.PlotBy

    @PlotBy.setter
    def PlotBy(self, value):
        self.com_object.PlotBy = value

    @property
    def plotby(self):
        """Lower case alias for PlotBy"""
        return self.PlotBy

    @plotby.setter
    def plotby(self, value):
        """Lower case alias for PlotBy.setter"""
        self.PlotBy = value

    @property
    def PlotVisibleOnly(self):
        return self.com_object.PlotVisibleOnly

    @PlotVisibleOnly.setter
    def PlotVisibleOnly(self, value):
        self.com_object.PlotVisibleOnly = value

    @property
    def plotvisibleonly(self):
        """Lower case alias for PlotVisibleOnly"""
        return self.PlotVisibleOnly

    @plotvisibleonly.setter
    def plotvisibleonly(self, value):
        """Lower case alias for PlotVisibleOnly.setter"""
        self.PlotVisibleOnly = value

    @property
    def RightAngleAxes(self):
        return self.com_object.RightAngleAxes

    @RightAngleAxes.setter
    def RightAngleAxes(self, value):
        self.com_object.RightAngleAxes = value

    @property
    def rightangleaxes(self):
        """Lower case alias for RightAngleAxes"""
        return self.RightAngleAxes

    @rightangleaxes.setter
    def rightangleaxes(self, value):
        """Lower case alias for RightAngleAxes.setter"""
        self.RightAngleAxes = value

    @property
    def Rotation(self):
        return self.com_object.Rotation

    @Rotation.setter
    def Rotation(self, value):
        self.com_object.Rotation = value

    @property
    def rotation(self):
        """Lower case alias for Rotation"""
        return self.Rotation

    @rotation.setter
    def rotation(self, value):
        """Lower case alias for Rotation.setter"""
        self.Rotation = value

    @property
    def seriesnamelevel(self):
        return self.com_object.seriesnamelevel

    @seriesnamelevel.setter
    def seriesnamelevel(self, value):
        self.com_object.seriesnamelevel = value

    @property
    def Shapes(self):
        return Shapes(self.com_object.Shapes)

    @property
    def shapes(self):
        """Lower case alias for Shapes"""
        return self.Shapes

    @property
    def ShowAllFieldButtons(self):
        return self.com_object.ShowAllFieldButtons

    @ShowAllFieldButtons.setter
    def ShowAllFieldButtons(self, value):
        self.com_object.ShowAllFieldButtons = value

    @property
    def showallfieldbuttons(self):
        """Lower case alias for ShowAllFieldButtons"""
        return self.ShowAllFieldButtons

    @showallfieldbuttons.setter
    def showallfieldbuttons(self, value):
        """Lower case alias for ShowAllFieldButtons.setter"""
        self.ShowAllFieldButtons = value

    @property
    def ShowAxisFieldButtons(self):
        return self.com_object.ShowAxisFieldButtons

    @ShowAxisFieldButtons.setter
    def ShowAxisFieldButtons(self, value):
        self.com_object.ShowAxisFieldButtons = value

    @property
    def showaxisfieldbuttons(self):
        """Lower case alias for ShowAxisFieldButtons"""
        return self.ShowAxisFieldButtons

    @showaxisfieldbuttons.setter
    def showaxisfieldbuttons(self, value):
        """Lower case alias for ShowAxisFieldButtons.setter"""
        self.ShowAxisFieldButtons = value

    @property
    def ShowDataLabelsOverMaximum(self):
        return self.com_object.ShowDataLabelsOverMaximum

    @ShowDataLabelsOverMaximum.setter
    def ShowDataLabelsOverMaximum(self, value):
        self.com_object.ShowDataLabelsOverMaximum = value

    @property
    def showdatalabelsovermaximum(self):
        """Lower case alias for ShowDataLabelsOverMaximum"""
        return self.ShowDataLabelsOverMaximum

    @showdatalabelsovermaximum.setter
    def showdatalabelsovermaximum(self, value):
        """Lower case alias for ShowDataLabelsOverMaximum.setter"""
        self.ShowDataLabelsOverMaximum = value

    @property
    def ShowLegendFieldButtons(self):
        return self.com_object.ShowLegendFieldButtons

    @ShowLegendFieldButtons.setter
    def ShowLegendFieldButtons(self, value):
        self.com_object.ShowLegendFieldButtons = value

    @property
    def showlegendfieldbuttons(self):
        """Lower case alias for ShowLegendFieldButtons"""
        return self.ShowLegendFieldButtons

    @showlegendfieldbuttons.setter
    def showlegendfieldbuttons(self, value):
        """Lower case alias for ShowLegendFieldButtons.setter"""
        self.ShowLegendFieldButtons = value

    @property
    def ShowReportFilterFieldButtons(self):
        return self.com_object.ShowReportFilterFieldButtons

    @ShowReportFilterFieldButtons.setter
    def ShowReportFilterFieldButtons(self, value):
        self.com_object.ShowReportFilterFieldButtons = value

    @property
    def showreportfilterfieldbuttons(self):
        """Lower case alias for ShowReportFilterFieldButtons"""
        return self.ShowReportFilterFieldButtons

    @showreportfilterfieldbuttons.setter
    def showreportfilterfieldbuttons(self, value):
        """Lower case alias for ShowReportFilterFieldButtons.setter"""
        self.ShowReportFilterFieldButtons = value

    @property
    def ShowValueFieldButtons(self):
        return self.com_object.ShowValueFieldButtons

    @ShowValueFieldButtons.setter
    def ShowValueFieldButtons(self, value):
        self.com_object.ShowValueFieldButtons = value

    @property
    def showvaluefieldbuttons(self):
        """Lower case alias for ShowValueFieldButtons"""
        return self.ShowValueFieldButtons

    @showvaluefieldbuttons.setter
    def showvaluefieldbuttons(self, value):
        """Lower case alias for ShowValueFieldButtons.setter"""
        self.ShowValueFieldButtons = value

    @property
    def SideWall(self):
        return Walls(self.com_object.SideWall)

    @property
    def sidewall(self):
        """Lower case alias for SideWall"""
        return self.SideWall

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
    def Walls(self):
        return Walls(self.com_object.Walls)

    @property
    def walls(self):
        """Lower case alias for Walls"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

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
    def Shadow(self):
        return self.com_object.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.com_object.Shadow = value

    @property
    def shadow(self):
        """Lower case alias for Shadow"""
        return self.Shadow

    @shadow.setter
    def shadow(self, value):
        """Lower case alias for Shadow.setter"""
        self.Shadow = value

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Color(self):
        return self.com_object.Color

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
    def ColorIndex(self):
        return self.com_object.ColorIndex

    @ColorIndex.setter
    def ColorIndex(self, value):
        self.com_object.ColorIndex = value

    @property
    def colorindex(self):
        """Lower case alias for ColorIndex"""
        return self.ColorIndex

    @colorindex.setter
    def colorindex(self, value):
        """Lower case alias for ColorIndex.setter"""
        self.ColorIndex = value

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def LineStyle(self):
        return XlLineStyle(self.com_object.LineStyle)

    @LineStyle.setter
    def LineStyle(self, value):
        self.com_object.LineStyle = value

    @property
    def linestyle(self):
        """Lower case alias for LineStyle"""
        return self.LineStyle

    @linestyle.setter
    def linestyle(self, value):
        """Lower case alias for LineStyle.setter"""
        self.LineStyle = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Weight(self):
        return XlBorderWeight(self.com_object.Weight)

    @Weight.setter
    def Weight(self, value):
        self.com_object.Weight = value

    @property
    def weight(self):
        """Lower case alias for Weight"""
        return self.Weight

    @weight.setter
    def weight(self, value):
        """Lower case alias for Weight.setter"""
        self.Weight = value


class chartcategory:

    def __init__(self, chartcategory=None):
        self.com_object= chartcategory

    @property
    def isfiltered(self):
        return self.com_object.isfiltered

    @isfiltered.setter
    def isfiltered(self, value):
        self.com_object.isfiltered = value

    @property
    def name(self):
        return self.com_object.name

    @name.setter
    def name(self, value):
        self.com_object.name = value

    @property
    def parent(self):
        return self.com_object.parent


class ChartCharacters:

    def __init__(self, chartcharacters=None):
        self.com_object= chartcharacters

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Caption(self):
        return self.com_object.Caption

    @property
    def caption(self):
        """Lower case alias for Caption"""
        return self.Caption

    @property
    def Count(self):
        return self.com_object.Count

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Font(self):
        return ChartFont(self.com_object.Font)

    @property
    def font(self):
        """Lower case alias for Font"""
        return self.Font

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PhoneticCharacters(self):
        return self.com_object.PhoneticCharacters

    @PhoneticCharacters.setter
    def PhoneticCharacters(self, value):
        self.com_object.PhoneticCharacters = value

    @property
    def phoneticcharacters(self):
        """Lower case alias for PhoneticCharacters"""
        return self.PhoneticCharacters

    @phoneticcharacters.setter
    def phoneticcharacters(self, value):
        """Lower case alias for PhoneticCharacters.setter"""
        self.PhoneticCharacters = value

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

    @property
    def islinked(self):
        """Lower case alias for IsLinked"""
        return self.IsLinked

    @property
    def Workbook(self):
        return self.com_object.Workbook

    @property
    def workbook(self):
        """Lower case alias for Workbook"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Background(self):
        return XlBackground(self.com_object.Background)

    @Background.setter
    def Background(self, value):
        self.com_object.Background = value

    @property
    def background(self):
        """Lower case alias for Background"""
        return self.Background

    @background.setter
    def background(self, value):
        """Lower case alias for Background.setter"""
        self.Background = value

    @property
    def Bold(self):
        return self.com_object.Bold

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
    def Color(self):
        return self.com_object.Color

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
    def ColorIndex(self):
        return self.com_object.ColorIndex

    @ColorIndex.setter
    def ColorIndex(self, value):
        self.com_object.ColorIndex = value

    @property
    def colorindex(self):
        """Lower case alias for ColorIndex"""
        return self.ColorIndex

    @colorindex.setter
    def colorindex(self, value):
        """Lower case alias for ColorIndex.setter"""
        self.ColorIndex = value

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def FontStyle(self):
        return self.com_object.FontStyle

    @FontStyle.setter
    def FontStyle(self, value):
        self.com_object.FontStyle = value

    @property
    def fontstyle(self):
        """Lower case alias for FontStyle"""
        return self.FontStyle

    @fontstyle.setter
    def fontstyle(self, value):
        """Lower case alias for FontStyle.setter"""
        self.FontStyle = value

    @property
    def Italic(self):
        return self.com_object.Italic

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
    def StrikeThrough(self):
        return self.com_object.StrikeThrough

    @StrikeThrough.setter
    def StrikeThrough(self, value):
        self.com_object.StrikeThrough = value

    @property
    def strikethrough(self):
        """Lower case alias for StrikeThrough"""
        return self.StrikeThrough

    @strikethrough.setter
    def strikethrough(self, value):
        """Lower case alias for StrikeThrough.setter"""
        self.StrikeThrough = value

    @property
    def Subscript(self):
        return self.com_object.Subscript

    @Subscript.setter
    def Subscript(self, value):
        self.com_object.Subscript = value

    @property
    def subscript(self):
        """Lower case alias for Subscript"""
        return self.Subscript

    @subscript.setter
    def subscript(self, value):
        """Lower case alias for Subscript.setter"""
        self.Subscript = value

    @property
    def superscript(self):
        return self.com_object.superscript

    @superscript.setter
    def superscript(self, value):
        self.com_object.superscript = value

    @property
    def Underline(self):
        return XlUnderlineStyle(self.com_object.Underline)

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


class ChartFormat:

    def __init__(self, chartformat=None):
        self.com_object= chartformat

    @property
    def adjustments(self):
        return self.com_object.adjustments

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def autoshapetype(self):
        return self.com_object.autoshapetype

    @autoshapetype.setter
    def autoshapetype(self, value):
        self.com_object.autoshapetype = value

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Fill(self):
        return FillFormat(self.com_object.Fill)

    @property
    def fill(self):
        """Lower case alias for Fill"""
        return self.Fill

    @property
    def Glow(self):
        return self.com_object.Glow

    @property
    def glow(self):
        """Lower case alias for Glow"""
        return self.Glow

    @property
    def Line(self):
        return LineFormat(self.com_object.Line)

    @property
    def line(self):
        """Lower case alias for Line"""
        return self.Line

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PictureFormat(self):
        return PictureFormat(self.com_object.PictureFormat)

    @property
    def pictureformat(self):
        """Lower case alias for PictureFormat"""
        return self.PictureFormat

    @property
    def Shadow(self):
        return ShadowFormat(self.com_object.Shadow)

    @property
    def shadow(self):
        """Lower case alias for Shadow"""
        return self.Shadow

    @property
    def SoftEdge(self):
        return self.com_object.SoftEdge

    @property
    def softedge(self):
        """Lower case alias for SoftEdge"""
        return self.SoftEdge

    @property
    def TextFrame2(self):
        return TextFrame2(self.com_object.TextFrame2)

    @property
    def textframe2(self):
        """Lower case alias for TextFrame2"""
        return self.TextFrame2

    @property
    def ThreeD(self):
        return ThreeDFormat(self.com_object.ThreeD)

    @property
    def threed(self):
        """Lower case alias for ThreeD"""
        return self.ThreeD


class ChartGroup:

    def __init__(self, chartgroup=None):
        self.com_object= chartgroup

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def AxisGroup(self):
        return XlAxisGroup(self.com_object.AxisGroup)

    @AxisGroup.setter
    def AxisGroup(self, value):
        self.com_object.AxisGroup = value

    @property
    def axisgroup(self):
        """Lower case alias for AxisGroup"""
        return self.AxisGroup

    @axisgroup.setter
    def axisgroup(self, value):
        """Lower case alias for AxisGroup.setter"""
        self.AxisGroup = value

    @property
    def binscountvalue(self):
        return self.com_object.binscountvalue

    @binscountvalue.setter
    def binscountvalue(self, value):
        self.com_object.binscountvalue = value

    @property
    def binsoverflowenabled(self):
        return self.com_object.binsoverflowenabled

    @binsoverflowenabled.setter
    def binsoverflowenabled(self, value):
        self.com_object.binsoverflowenabled = value

    @property
    def binsoverflowvalue(self):
        return self.com_object.binsoverflowvalue

    @binsoverflowvalue.setter
    def binsoverflowvalue(self, value):
        self.com_object.binsoverflowvalue = value

    @property
    def binstype(self):
        return self.com_object.binstype

    @binstype.setter
    def binstype(self, value):
        self.com_object.binstype = value

    @property
    def binsunderflowenabled(self):
        return self.com_object.binsunderflowenabled

    @binsunderflowenabled.setter
    def binsunderflowenabled(self, value):
        self.com_object.binsunderflowenabled = value

    @property
    def binsunderflowvalue(self):
        return self.com_object.binsunderflowvalue

    @binsunderflowvalue.setter
    def binsunderflowvalue(self, value):
        self.com_object.binsunderflowvalue = value

    @property
    def binwidthvalue(self):
        return self.com_object.binwidthvalue

    @binwidthvalue.setter
    def binwidthvalue(self, value):
        self.com_object.binwidthvalue = value

    @property
    def BubbleScale(self):
        return self.com_object.BubbleScale

    @BubbleScale.setter
    def BubbleScale(self, value):
        self.com_object.BubbleScale = value

    @property
    def bubblescale(self):
        """Lower case alias for BubbleScale"""
        return self.BubbleScale

    @bubblescale.setter
    def bubblescale(self, value):
        """Lower case alias for BubbleScale.setter"""
        self.BubbleScale = value

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def DoughnutHoleSize(self):
        return self.com_object.DoughnutHoleSize

    @DoughnutHoleSize.setter
    def DoughnutHoleSize(self, value):
        self.com_object.DoughnutHoleSize = value

    @property
    def doughnutholesize(self):
        """Lower case alias for DoughnutHoleSize"""
        return self.DoughnutHoleSize

    @doughnutholesize.setter
    def doughnutholesize(self, value):
        """Lower case alias for DoughnutHoleSize.setter"""
        self.DoughnutHoleSize = value

    @property
    def DownBars(self):
        return DownBars(self.com_object.DownBars)

    @property
    def downbars(self):
        """Lower case alias for DownBars"""
        return self.DownBars

    @property
    def DropLines(self):
        return DropLines(self.com_object.DropLines)

    @property
    def droplines(self):
        """Lower case alias for DropLines"""
        return self.DropLines

    @property
    def FirstSliceAngle(self):
        return self.com_object.FirstSliceAngle

    @FirstSliceAngle.setter
    def FirstSliceAngle(self, value):
        self.com_object.FirstSliceAngle = value

    @property
    def firstsliceangle(self):
        """Lower case alias for FirstSliceAngle"""
        return self.FirstSliceAngle

    @firstsliceangle.setter
    def firstsliceangle(self, value):
        """Lower case alias for FirstSliceAngle.setter"""
        self.FirstSliceAngle = value

    @property
    def GapWidth(self):
        return self.com_object.GapWidth

    @GapWidth.setter
    def GapWidth(self, value):
        self.com_object.GapWidth = value

    @property
    def gapwidth(self):
        """Lower case alias for GapWidth"""
        return self.GapWidth

    @gapwidth.setter
    def gapwidth(self, value):
        """Lower case alias for GapWidth.setter"""
        self.GapWidth = value

    @property
    def Has3DShading(self):
        return self.com_object.Has3DShading

    @Has3DShading.setter
    def Has3DShading(self, value):
        self.com_object.Has3DShading = value

    @property
    def has3dshading(self):
        """Lower case alias for Has3DShading"""
        return self.Has3DShading

    @has3dshading.setter
    def has3dshading(self, value):
        """Lower case alias for Has3DShading.setter"""
        self.Has3DShading = value

    @property
    def HasDropLines(self):
        return self.com_object.HasDropLines

    @HasDropLines.setter
    def HasDropLines(self, value):
        self.com_object.HasDropLines = value

    @property
    def hasdroplines(self):
        """Lower case alias for HasDropLines"""
        return self.HasDropLines

    @hasdroplines.setter
    def hasdroplines(self, value):
        """Lower case alias for HasDropLines.setter"""
        self.HasDropLines = value

    @property
    def HasHiLoLines(self):
        return self.com_object.HasHiLoLines

    @HasHiLoLines.setter
    def HasHiLoLines(self, value):
        self.com_object.HasHiLoLines = value

    @property
    def hashilolines(self):
        """Lower case alias for HasHiLoLines"""
        return self.HasHiLoLines

    @hashilolines.setter
    def hashilolines(self, value):
        """Lower case alias for HasHiLoLines.setter"""
        self.HasHiLoLines = value

    @property
    def HasRadarAxisLabels(self):
        return self.com_object.HasRadarAxisLabels

    @HasRadarAxisLabels.setter
    def HasRadarAxisLabels(self, value):
        self.com_object.HasRadarAxisLabels = value

    @property
    def hasradaraxislabels(self):
        """Lower case alias for HasRadarAxisLabels"""
        return self.HasRadarAxisLabels

    @hasradaraxislabels.setter
    def hasradaraxislabels(self, value):
        """Lower case alias for HasRadarAxisLabels.setter"""
        self.HasRadarAxisLabels = value

    @property
    def HasSeriesLines(self):
        return self.com_object.HasSeriesLines

    @HasSeriesLines.setter
    def HasSeriesLines(self, value):
        self.com_object.HasSeriesLines = value

    @property
    def hasserieslines(self):
        """Lower case alias for HasSeriesLines"""
        return self.HasSeriesLines

    @hasserieslines.setter
    def hasserieslines(self, value):
        """Lower case alias for HasSeriesLines.setter"""
        self.HasSeriesLines = value

    @property
    def HasUpDownBars(self):
        return self.com_object.HasUpDownBars

    @HasUpDownBars.setter
    def HasUpDownBars(self, value):
        self.com_object.HasUpDownBars = value

    @property
    def hasupdownbars(self):
        """Lower case alias for HasUpDownBars"""
        return self.HasUpDownBars

    @hasupdownbars.setter
    def hasupdownbars(self, value):
        """Lower case alias for HasUpDownBars.setter"""
        self.HasUpDownBars = value

    @property
    def HiLoLines(self):
        return HiLoLines(self.com_object.HiLoLines)

    @property
    def hilolines(self):
        """Lower case alias for HiLoLines"""
        return self.HiLoLines

    @property
    def Index(self):
        return self.com_object.Index

    @property
    def index(self):
        """Lower case alias for Index"""
        return self.Index

    @property
    def Overlap(self):
        return self.com_object.Overlap

    @Overlap.setter
    def Overlap(self, value):
        self.com_object.Overlap = value

    @property
    def overlap(self):
        """Lower case alias for Overlap"""
        return self.Overlap

    @overlap.setter
    def overlap(self, value):
        """Lower case alias for Overlap.setter"""
        self.Overlap = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def RadarAxisLabels(self):
        return TickLabels(self.com_object.RadarAxisLabels)

    @property
    def radaraxislabels(self):
        """Lower case alias for RadarAxisLabels"""
        return self.RadarAxisLabels

    @property
    def SecondPlotSize(self):
        return self.com_object.SecondPlotSize

    @SecondPlotSize.setter
    def SecondPlotSize(self, value):
        self.com_object.SecondPlotSize = value

    @property
    def secondplotsize(self):
        """Lower case alias for SecondPlotSize"""
        return self.SecondPlotSize

    @secondplotsize.setter
    def secondplotsize(self, value):
        """Lower case alias for SecondPlotSize.setter"""
        self.SecondPlotSize = value

    @property
    def SeriesLines(self):
        return SeriesLines(self.com_object.SeriesLines)

    @property
    def serieslines(self):
        """Lower case alias for SeriesLines"""
        return self.SeriesLines

    @property
    def ShowNegativeBubbles(self):
        return self.com_object.ShowNegativeBubbles

    @ShowNegativeBubbles.setter
    def ShowNegativeBubbles(self, value):
        self.com_object.ShowNegativeBubbles = value

    @property
    def shownegativebubbles(self):
        """Lower case alias for ShowNegativeBubbles"""
        return self.ShowNegativeBubbles

    @shownegativebubbles.setter
    def shownegativebubbles(self, value):
        """Lower case alias for ShowNegativeBubbles.setter"""
        self.ShowNegativeBubbles = value

    @property
    def SizeRepresents(self):
        return self.com_object.SizeRepresents

    @SizeRepresents.setter
    def SizeRepresents(self, value):
        self.com_object.SizeRepresents = value

    @property
    def sizerepresents(self):
        """Lower case alias for SizeRepresents"""
        return self.SizeRepresents

    @sizerepresents.setter
    def sizerepresents(self, value):
        """Lower case alias for SizeRepresents.setter"""
        self.SizeRepresents = value

    @property
    def SplitType(self):
        return XlChartSplitType(self.com_object.SplitType)

    @SplitType.setter
    def SplitType(self, value):
        self.com_object.SplitType = value

    @property
    def splittype(self):
        """Lower case alias for SplitType"""
        return self.SplitType

    @splittype.setter
    def splittype(self, value):
        """Lower case alias for SplitType.setter"""
        self.SplitType = value

    @property
    def SplitValue(self):
        return self.com_object.SplitValue

    @SplitValue.setter
    def SplitValue(self, value):
        self.com_object.SplitValue = value

    @property
    def splitvalue(self):
        """Lower case alias for SplitValue"""
        return self.SplitValue

    @splitvalue.setter
    def splitvalue(self, value):
        """Lower case alias for SplitValue.setter"""
        self.SplitValue = value

    @property
    def UpBars(self):
        return UpBars(self.com_object.UpBars)

    @property
    def upbars(self):
        """Lower case alias for UpBars"""
        return self.UpBars

    @property
    def VaryByCategories(self):
        return self.com_object.VaryByCategories

    @VaryByCategories.setter
    def VaryByCategories(self, value):
        self.com_object.VaryByCategories = value

    @property
    def varybycategories(self):
        """Lower case alias for VaryByCategories"""
        return self.VaryByCategories

    @varybycategories.setter
    def varybycategories(self, value):
        """Lower case alias for VaryByCategories.setter"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Count(self):
        return self.com_object.Count

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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

    def Characters(self, Start=None, Length=None):
        arguments = com_arguments([unwrap(a) for a in [Start, Length]])
        if hasattr(self.com_object, "GetCharacters"):
            return ChartCharacters(self.com_object.GetCharacters(*arguments))
        else:
            return ChartCharacters(self.com_object.Characters(*arguments))

    def characters(self, Start=None, Length=None):
        """Lower case alias for Characters"""
        arguments = [Start, Length]
        return self.Characters(*arguments)

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

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
    def FormulaLocal(self):
        return self.com_object.FormulaLocal

    @FormulaLocal.setter
    def FormulaLocal(self, value):
        self.com_object.FormulaLocal = value

    @property
    def formulalocal(self):
        """Lower case alias for FormulaLocal"""
        return self.FormulaLocal

    @formulalocal.setter
    def formulalocal(self, value):
        """Lower case alias for FormulaLocal.setter"""
        self.FormulaLocal = value

    @property
    def FormulaR1C1(self):
        return self.com_object.FormulaR1C1

    @FormulaR1C1.setter
    def FormulaR1C1(self, value):
        self.com_object.FormulaR1C1 = value

    @property
    def formular1c1(self):
        """Lower case alias for FormulaR1C1"""
        return self.FormulaR1C1

    @formular1c1.setter
    def formular1c1(self, value):
        """Lower case alias for FormulaR1C1.setter"""
        self.FormulaR1C1 = value

    @property
    def FormulaR1C1Local(self):
        return self.com_object.FormulaR1C1Local

    @FormulaR1C1Local.setter
    def FormulaR1C1Local(self, value):
        self.com_object.FormulaR1C1Local = value

    @property
    def formular1c1local(self):
        """Lower case alias for FormulaR1C1Local"""
        return self.FormulaR1C1Local

    @formular1c1local.setter
    def formular1c1local(self, value):
        """Lower case alias for FormulaR1C1Local.setter"""
        self.FormulaR1C1Local = value

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
    def HorizontalAlignment(self):
        return self.com_object.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.com_object.HorizontalAlignment = value

    @property
    def horizontalalignment(self):
        """Lower case alias for HorizontalAlignment"""
        return self.HorizontalAlignment

    @horizontalalignment.setter
    def horizontalalignment(self, value):
        """Lower case alias for HorizontalAlignment.setter"""
        self.HorizontalAlignment = value

    @property
    def IncludeInLayout(self):
        return self.com_object.IncludeInLayout

    @IncludeInLayout.setter
    def IncludeInLayout(self, value):
        self.com_object.IncludeInLayout = value

    @property
    def includeinlayout(self):
        """Lower case alias for IncludeInLayout"""
        return self.IncludeInLayout

    @includeinlayout.setter
    def includeinlayout(self, value):
        """Lower case alias for IncludeInLayout.setter"""
        self.IncludeInLayout = value

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
    def Name(self):
        return self.com_object.Name

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def Orientation(self):
        return self.com_object.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.com_object.Orientation = value

    @property
    def orientation(self):
        """Lower case alias for Orientation"""
        return self.Orientation

    @orientation.setter
    def orientation(self, value):
        """Lower case alias for Orientation.setter"""
        self.Orientation = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Position(self):
        return XlChartElementPosition(self.com_object.Position)

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
    def ReadingOrder(self):
        return XlReadingOrder(self.com_object.ReadingOrder)

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.com_object.ReadingOrder = value

    @property
    def readingorder(self):
        """Lower case alias for ReadingOrder"""
        return self.ReadingOrder

    @readingorder.setter
    def readingorder(self, value):
        """Lower case alias for ReadingOrder.setter"""
        self.ReadingOrder = value

    @property
    def Shadow(self):
        return self.com_object.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.com_object.Shadow = value

    @property
    def shadow(self):
        """Lower case alias for Shadow"""
        return self.Shadow

    @shadow.setter
    def shadow(self, value):
        """Lower case alias for Shadow.setter"""
        self.Shadow = value

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
    def VerticalAlignment(self):
        return self.com_object.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.com_object.VerticalAlignment = value

    @property
    def verticalalignment(self):
        """Lower case alias for VerticalAlignment"""
        return self.VerticalAlignment

    @verticalalignment.setter
    def verticalalignment(self, value):
        """Lower case alias for VerticalAlignment.setter"""
        self.VerticalAlignment = value

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def FavorServerEditsDuringMerge(self):
        return self.com_object.FavorServerEditsDuringMerge

    @FavorServerEditsDuringMerge.setter
    def FavorServerEditsDuringMerge(self, value):
        self.com_object.FavorServerEditsDuringMerge = value

    @property
    def favorservereditsduringmerge(self):
        """Lower case alias for FavorServerEditsDuringMerge"""
        return self.FavorServerEditsDuringMerge

    @favorservereditsduringmerge.setter
    def favorservereditsduringmerge(self, value):
        """Lower case alias for FavorServerEditsDuringMerge.setter"""
        self.FavorServerEditsDuringMerge = value

    @property
    def MergeMode(self):
        return self.com_object.MergeMode

    @property
    def mergemode(self):
        """Lower case alias for MergeMode"""
        return self.MergeMode

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PendingUpdates(self):
        return self.com_object.PendingUpdates

    @property
    def pendingupdates(self):
        """Lower case alias for PendingUpdates"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def By(self):
        return ColorFormat(self.com_object.By)

    @property
    def by(self):
        """Lower case alias for By"""
        return self.By

    @property
    def From(self):
        return self.com_object.From

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

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


class ColorFormat:

    def __init__(self, colorformat=None):
        self.com_object= colorformat

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Brightness(self):
        return self.com_object.Brightness

    @Brightness.setter
    def Brightness(self, value):
        self.com_object.Brightness = value

    @property
    def brightness(self):
        """Lower case alias for Brightness"""
        return self.Brightness

    @brightness.setter
    def brightness(self, value):
        """Lower case alias for Brightness.setter"""
        self.Brightness = value

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def ObjectThemeColor(self):
        return ColorFormat(self.com_object.ObjectThemeColor)

    @ObjectThemeColor.setter
    def ObjectThemeColor(self, value):
        self.com_object.ObjectThemeColor = value

    @property
    def objectthemecolor(self):
        """Lower case alias for ObjectThemeColor"""
        return self.ObjectThemeColor

    @objectthemecolor.setter
    def objectthemecolor(self, value):
        """Lower case alias for ObjectThemeColor.setter"""
        self.ObjectThemeColor = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def RGB(self):
        return self.com_object.RGB

    @RGB.setter
    def RGB(self, value):
        self.com_object.RGB = value

    @property
    def rgb(self):
        """Lower case alias for RGB"""
        return self.RGB

    @rgb.setter
    def rgb(self, value):
        """Lower case alias for RGB.setter"""
        self.RGB = value

    @property
    def SchemeColor(self):
        return self.com_object.SchemeColor

    @SchemeColor.setter
    def SchemeColor(self, value):
        self.com_object.SchemeColor = value

    @property
    def schemecolor(self):
        """Lower case alias for SchemeColor"""
        return self.SchemeColor

    @schemecolor.setter
    def schemecolor(self, value):
        """Lower case alias for SchemeColor.setter"""
        self.SchemeColor = value

    @property
    def TintAndShade(self):
        return self.com_object.TintAndShade

    @TintAndShade.setter
    def TintAndShade(self, value):
        self.com_object.TintAndShade = value

    @property
    def tintandshade(self):
        """Lower case alias for TintAndShade"""
        return self.TintAndShade

    @tintandshade.setter
    def tintandshade(self, value):
        """Lower case alias for TintAndShade.setter"""
        self.TintAndShade = value

    @property
    def Type(self):
        return self.com_object.Type

    @property
    def type(self):
        """Lower case alias for Type"""
        return self.Type


class ColorScheme:

    def __init__(self, colorscheme=None):
        self.com_object= colorscheme

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    def Cells(self, RowIndex=None, ColumnIndex=None):
        arguments = com_arguments([unwrap(a) for a in [RowIndex, ColumnIndex]])
        if hasattr(self.com_object, "GetCells"):
            return CellRange(self.com_object.GetCells(*arguments))
        else:
            return CellRange(self.com_object.Cells(*arguments))

    def cells(self, RowIndex=None, ColumnIndex=None):
        """Lower case alias for Cells"""
        arguments = [RowIndex, ColumnIndex]
        return self.Cells(*arguments)

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Bookmark(self):
        return self.com_object.Bookmark

    @Bookmark.setter
    def Bookmark(self, value):
        self.com_object.Bookmark = value

    @property
    def bookmark(self):
        """Lower case alias for Bookmark"""
        return self.Bookmark

    @bookmark.setter
    def bookmark(self, value):
        """Lower case alias for Bookmark.setter"""
        self.Bookmark = value

    @property
    def Command(self):
        return self.com_object.Command

    @Command.setter
    def Command(self, value):
        self.com_object.Command = value

    @property
    def command(self):
        """Lower case alias for Command"""
        return self.Command

    @command.setter
    def command(self, value):
        """Lower case alias for Command.setter"""
        self.Command = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

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


class Comment:

    def __init__(self, comment=None):
        self.com_object= comment

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Author(self):
        return Comment(self.com_object.Author)

    @property
    def author(self):
        """Lower case alias for Author"""
        return self.Author

    @property
    def AuthorIndex(self):
        return self.com_object.AuthorIndex

    @property
    def authorindex(self):
        """Lower case alias for AuthorIndex"""
        return self.AuthorIndex

    @property
    def AuthorInitials(self):
        return Comment(self.com_object.AuthorInitials)

    @property
    def authorinitials(self):
        """Lower case alias for AuthorInitials"""
        return self.AuthorInitials

    @property
    def collapsed(self):
        return self.com_object.collapsed

    @property
    def DateTime(self):
        return self.com_object.DateTime

    @property
    def datetime(self):
        """Lower case alias for DateTime"""
        return self.DateTime

    @property
    def Left(self):
        return self.com_object.Left

    @property
    def left(self):
        """Lower case alias for Left"""
        return self.Left

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def providerid(self):
        return self.com_object.providerid

    @property
    def replies(self):
        return Comment(self.com_object.replies)

    @property
    def Text(self):
        return self.com_object.Text

    @property
    def text(self):
        """Lower case alias for Text"""
        return self.Text

    @property
    def timezonebias(self):
        return self.com_object.timezonebias

    @property
    def Top(self):
        return self.com_object.Top

    @property
    def top(self):
        """Lower case alias for Top"""
        return self.Top

    @property
    def userid(self):
        return self.com_object.userid

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def BeginConnected(self):
        return self.com_object.BeginConnected

    @BeginConnected.setter
    def BeginConnected(self, value):
        self.com_object.BeginConnected = value

    @property
    def beginconnected(self):
        """Lower case alias for BeginConnected"""
        return self.BeginConnected

    @beginconnected.setter
    def beginconnected(self, value):
        """Lower case alias for BeginConnected.setter"""
        self.BeginConnected = value

    @property
    def BeginConnectedShape(self):
        return Shape(self.com_object.BeginConnectedShape)

    @property
    def beginconnectedshape(self):
        """Lower case alias for BeginConnectedShape"""
        return self.BeginConnectedShape

    @property
    def BeginConnectionSite(self):
        return self.com_object.BeginConnectionSite

    @property
    def beginconnectionsite(self):
        """Lower case alias for BeginConnectionSite"""
        return self.BeginConnectionSite

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def EndConnected(self):
        return self.com_object.EndConnected

    @property
    def endconnected(self):
        """Lower case alias for EndConnected"""
        return self.EndConnected

    @property
    def EndConnectedShape(self):
        return Shape(self.com_object.EndConnectedShape)

    @property
    def endconnectedshape(self):
        """Lower case alias for EndConnectedShape"""
        return self.EndConnectedShape

    @property
    def EndConnectionSite(self):
        return self.com_object.EndConnectionSite

    @property
    def endconnectionsite(self):
        """Lower case alias for EndConnectionSite"""
        return self.EndConnectionSite

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Count(self):
        return self.com_object.Count

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return CustomerData(self.com_object.Parent)

    @property
    def parent(self):
        """Lower case alias for Parent"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Background(self):
        return ShapeRange(self.com_object.Background)

    @property
    def background(self):
        """Lower case alias for Background"""
        return self.Background

    @property
    def CustomerData(self):
        return CustomerData(self.com_object.CustomerData)

    @property
    def customerdata(self):
        """Lower case alias for CustomerData"""
        return self.CustomerData

    @property
    def Design(self):
        return Design(self.com_object.Design)

    @property
    def design(self):
        """Lower case alias for Design"""
        return self.Design

    @property
    def DisplayMasterShapes(self):
        return self.com_object.DisplayMasterShapes

    @DisplayMasterShapes.setter
    def DisplayMasterShapes(self, value):
        self.com_object.DisplayMasterShapes = value

    @property
    def displaymastershapes(self):
        """Lower case alias for DisplayMasterShapes"""
        return self.DisplayMasterShapes

    @displaymastershapes.setter
    def displaymastershapes(self, value):
        """Lower case alias for DisplayMasterShapes.setter"""
        self.DisplayMasterShapes = value

    @property
    def FollowMasterBackground(self):
        return self.com_object.FollowMasterBackground

    @FollowMasterBackground.setter
    def FollowMasterBackground(self, value):
        self.com_object.FollowMasterBackground = value

    @property
    def followmasterbackground(self):
        """Lower case alias for FollowMasterBackground"""
        return self.FollowMasterBackground

    @followmasterbackground.setter
    def followmasterbackground(self, value):
        """Lower case alias for FollowMasterBackground.setter"""
        self.FollowMasterBackground = value

    @property
    def guides(self):
        return self.com_object.guides

    @property
    def HeadersFooters(self):
        return HeadersFooters(self.com_object.HeadersFooters)

    @property
    def headersfooters(self):
        """Lower case alias for HeadersFooters"""
        return self.HeadersFooters

    @property
    def Height(self):
        return self.com_object.Height

    @property
    def height(self):
        """Lower case alias for Height"""
        return self.Height

    @property
    def Hyperlinks(self):
        return Hyperlinks(self.com_object.Hyperlinks)

    @property
    def hyperlinks(self):
        """Lower case alias for Hyperlinks"""
        return self.Hyperlinks

    @property
    def Index(self):
        return CustomLayouts(self.com_object.Index)

    @property
    def index(self):
        """Lower case alias for Index"""
        return self.Index

    @property
    def MatchingName(self):
        return self.com_object.MatchingName

    @MatchingName.setter
    def MatchingName(self, value):
        self.com_object.MatchingName = value

    @property
    def matchingname(self):
        """Lower case alias for MatchingName"""
        return self.MatchingName

    @matchingname.setter
    def matchingname(self, value):
        """Lower case alias for MatchingName.setter"""
        self.MatchingName = value

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
        return CustomLayout(self.com_object.Parent)

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Preserved(self):
        return self.com_object.Preserved

    @Preserved.setter
    def Preserved(self, value):
        self.com_object.Preserved = value

    @property
    def preserved(self):
        """Lower case alias for Preserved"""
        return self.Preserved

    @preserved.setter
    def preserved(self, value):
        """Lower case alias for Preserved.setter"""
        self.Preserved = value

    @property
    def Shapes(self):
        return Shapes(self.com_object.Shapes)

    @property
    def shapes(self):
        """Lower case alias for Shapes"""
        return self.Shapes

    @property
    def SlideShowTransition(self):
        return SlideShowTransition(self.com_object.SlideShowTransition)

    @property
    def slideshowtransition(self):
        """Lower case alias for SlideShowTransition"""
        return self.SlideShowTransition

    @property
    def ThemeColorScheme(self):
        return self.com_object.ThemeColorScheme

    @property
    def themecolorscheme(self):
        """Lower case alias for ThemeColorScheme"""
        return self.ThemeColorScheme

    @property
    def TimeLine(self):
        return TimeLine(self.com_object.TimeLine)

    @property
    def timeline(self):
        """Lower case alias for TimeLine"""
        return self.TimeLine

    @property
    def Width(self):
        return self.com_object.Width

    @property
    def width(self):
        """Lower case alias for Width"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def AutoText(self):
        return self.com_object.AutoText

    @AutoText.setter
    def AutoText(self, value):
        self.com_object.AutoText = value

    @property
    def autotext(self):
        """Lower case alias for AutoText"""
        return self.AutoText

    @autotext.setter
    def autotext(self, value):
        """Lower case alias for AutoText.setter"""
        self.AutoText = value

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

    def Characters(self, Start=None, Length=None):
        arguments = com_arguments([unwrap(a) for a in [Start, Length]])
        if hasattr(self.com_object, "GetCharacters"):
            return ChartCharacters(self.com_object.GetCharacters(*arguments))
        else:
            return ChartCharacters(self.com_object.Characters(*arguments))

    def characters(self, Start=None, Length=None):
        """Lower case alias for Characters"""
        arguments = [Start, Length]
        return self.Characters(*arguments)

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

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
    def FormulaLocal(self):
        return self.com_object.FormulaLocal

    @FormulaLocal.setter
    def FormulaLocal(self, value):
        self.com_object.FormulaLocal = value

    @property
    def formulalocal(self):
        """Lower case alias for FormulaLocal"""
        return self.FormulaLocal

    @formulalocal.setter
    def formulalocal(self, value):
        """Lower case alias for FormulaLocal.setter"""
        self.FormulaLocal = value

    @property
    def FormulaR1C1(self):
        return self.com_object.FormulaR1C1

    @FormulaR1C1.setter
    def FormulaR1C1(self, value):
        self.com_object.FormulaR1C1 = value

    @property
    def formular1c1(self):
        """Lower case alias for FormulaR1C1"""
        return self.FormulaR1C1

    @formular1c1.setter
    def formular1c1(self, value):
        """Lower case alias for FormulaR1C1.setter"""
        self.FormulaR1C1 = value

    @property
    def FormulaR1C1Local(self):
        return self.com_object.FormulaR1C1Local

    @FormulaR1C1Local.setter
    def FormulaR1C1Local(self, value):
        self.com_object.FormulaR1C1Local = value

    @property
    def formular1c1local(self):
        """Lower case alias for FormulaR1C1Local"""
        return self.FormulaR1C1Local

    @formular1c1local.setter
    def formular1c1local(self, value):
        """Lower case alias for FormulaR1C1Local.setter"""
        self.FormulaR1C1Local = value

    @property
    def Height(self):
        return self.com_object.Height

    @property
    def height(self):
        """Lower case alias for Height"""
        return self.Height

    @property
    def HorizontalAlignment(self):
        return self.com_object.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.com_object.HorizontalAlignment = value

    @property
    def horizontalalignment(self):
        """Lower case alias for HorizontalAlignment"""
        return self.HorizontalAlignment

    @horizontalalignment.setter
    def horizontalalignment(self, value):
        """Lower case alias for HorizontalAlignment.setter"""
        self.HorizontalAlignment = value

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
    def Name(self):
        return self.com_object.Name

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def NumberFormat(self):
        return self.com_object.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value):
        self.com_object.NumberFormat = value

    @property
    def numberformat(self):
        """Lower case alias for NumberFormat"""
        return self.NumberFormat

    @numberformat.setter
    def numberformat(self, value):
        """Lower case alias for NumberFormat.setter"""
        self.NumberFormat = value

    @property
    def NumberFormatLinked(self):
        return self.com_object.NumberFormatLinked

    @NumberFormatLinked.setter
    def NumberFormatLinked(self, value):
        self.com_object.NumberFormatLinked = value

    @property
    def numberformatlinked(self):
        """Lower case alias for NumberFormatLinked"""
        return self.NumberFormatLinked

    @numberformatlinked.setter
    def numberformatlinked(self, value):
        """Lower case alias for NumberFormatLinked.setter"""
        self.NumberFormatLinked = value

    @property
    def NumberFormatLocal(self):
        return self.com_object.NumberFormatLocal

    @NumberFormatLocal.setter
    def NumberFormatLocal(self, value):
        self.com_object.NumberFormatLocal = value

    @property
    def numberformatlocal(self):
        """Lower case alias for NumberFormatLocal"""
        return self.NumberFormatLocal

    @numberformatlocal.setter
    def numberformatlocal(self, value):
        """Lower case alias for NumberFormatLocal.setter"""
        self.NumberFormatLocal = value

    @property
    def Orientation(self):
        return self.com_object.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.com_object.Orientation = value

    @property
    def orientation(self):
        """Lower case alias for Orientation"""
        return self.Orientation

    @orientation.setter
    def orientation(self, value):
        """Lower case alias for Orientation.setter"""
        self.Orientation = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Position(self):
        return XlDataLabelPosition(self.com_object.Position)

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
    def ReadingOrder(self):
        return XlReadingOrder(self.com_object.ReadingOrder)

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.com_object.ReadingOrder = value

    @property
    def readingorder(self):
        """Lower case alias for ReadingOrder"""
        return self.ReadingOrder

    @readingorder.setter
    def readingorder(self, value):
        """Lower case alias for ReadingOrder.setter"""
        self.ReadingOrder = value

    @property
    def Separator(self):
        return self.com_object.Separator

    @Separator.setter
    def Separator(self, value):
        self.com_object.Separator = value

    @property
    def separator(self):
        """Lower case alias for Separator"""
        return self.Separator

    @separator.setter
    def separator(self, value):
        """Lower case alias for Separator.setter"""
        self.Separator = value

    @property
    def Shadow(self):
        return self.com_object.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.com_object.Shadow = value

    @property
    def shadow(self):
        """Lower case alias for Shadow"""
        return self.Shadow

    @shadow.setter
    def shadow(self, value):
        """Lower case alias for Shadow.setter"""
        self.Shadow = value

    @property
    def ShowBubbleSize(self):
        return self.com_object.ShowBubbleSize

    @ShowBubbleSize.setter
    def ShowBubbleSize(self, value):
        self.com_object.ShowBubbleSize = value

    @property
    def showbubblesize(self):
        """Lower case alias for ShowBubbleSize"""
        return self.ShowBubbleSize

    @showbubblesize.setter
    def showbubblesize(self, value):
        """Lower case alias for ShowBubbleSize.setter"""
        self.ShowBubbleSize = value

    @property
    def ShowCategoryName(self):
        return self.com_object.ShowCategoryName

    @ShowCategoryName.setter
    def ShowCategoryName(self, value):
        self.com_object.ShowCategoryName = value

    @property
    def showcategoryname(self):
        """Lower case alias for ShowCategoryName"""
        return self.ShowCategoryName

    @showcategoryname.setter
    def showcategoryname(self, value):
        """Lower case alias for ShowCategoryName.setter"""
        self.ShowCategoryName = value

    @property
    def ShowLegendKey(self):
        return self.com_object.ShowLegendKey

    @ShowLegendKey.setter
    def ShowLegendKey(self, value):
        self.com_object.ShowLegendKey = value

    @property
    def showlegendkey(self):
        """Lower case alias for ShowLegendKey"""
        return self.ShowLegendKey

    @showlegendkey.setter
    def showlegendkey(self, value):
        """Lower case alias for ShowLegendKey.setter"""
        self.ShowLegendKey = value

    @property
    def ShowPercentage(self):
        return self.com_object.ShowPercentage

    @ShowPercentage.setter
    def ShowPercentage(self, value):
        self.com_object.ShowPercentage = value

    @property
    def showpercentage(self):
        """Lower case alias for ShowPercentage"""
        return self.ShowPercentage

    @showpercentage.setter
    def showpercentage(self, value):
        """Lower case alias for ShowPercentage.setter"""
        self.ShowPercentage = value

    @property
    def showrange(self):
        return self.com_object.showrange

    @showrange.setter
    def showrange(self, value):
        self.com_object.showrange = value

    @property
    def ShowSeriesName(self):
        return self.com_object.ShowSeriesName

    @ShowSeriesName.setter
    def ShowSeriesName(self, value):
        self.com_object.ShowSeriesName = value

    @property
    def showseriesname(self):
        """Lower case alias for ShowSeriesName"""
        return self.ShowSeriesName

    @showseriesname.setter
    def showseriesname(self, value):
        """Lower case alias for ShowSeriesName.setter"""
        self.ShowSeriesName = value

    @property
    def ShowValue(self):
        return self.com_object.ShowValue

    @ShowValue.setter
    def ShowValue(self, value):
        self.com_object.ShowValue = value

    @property
    def showvalue(self):
        """Lower case alias for ShowValue"""
        return self.ShowValue

    @showvalue.setter
    def showvalue(self, value):
        """Lower case alias for ShowValue.setter"""
        self.ShowValue = value

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
    def VerticalAlignment(self):
        return self.com_object.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.com_object.VerticalAlignment = value

    @property
    def verticalalignment(self):
        """Lower case alias for VerticalAlignment"""
        return self.VerticalAlignment

    @verticalalignment.setter
    def verticalalignment(self, value):
        """Lower case alias for VerticalAlignment.setter"""
        self.VerticalAlignment = value

    @property
    def Width(self):
        return self.com_object.Width

    @property
    def width(self):
        """Lower case alias for Width"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def AutoText(self):
        return self.com_object.AutoText

    @AutoText.setter
    def AutoText(self, value):
        self.com_object.AutoText = value

    @property
    def autotext(self):
        """Lower case alias for AutoText"""
        return self.AutoText

    @autotext.setter
    def autotext(self, value):
        """Lower case alias for AutoText.setter"""
        self.AutoText = value

    @property
    def Count(self):
        return self.com_object.Count

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

    @property
    def HorizontalAlignment(self):
        return self.com_object.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.com_object.HorizontalAlignment = value

    @property
    def horizontalalignment(self):
        """Lower case alias for HorizontalAlignment"""
        return self.HorizontalAlignment

    @horizontalalignment.setter
    def horizontalalignment(self, value):
        """Lower case alias for HorizontalAlignment.setter"""
        self.HorizontalAlignment = value

    @property
    def Name(self):
        return self.com_object.Name

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def NumberFormat(self):
        return self.com_object.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value):
        self.com_object.NumberFormat = value

    @property
    def numberformat(self):
        """Lower case alias for NumberFormat"""
        return self.NumberFormat

    @numberformat.setter
    def numberformat(self, value):
        """Lower case alias for NumberFormat.setter"""
        self.NumberFormat = value

    @property
    def NumberFormatLinked(self):
        return self.com_object.NumberFormatLinked

    @NumberFormatLinked.setter
    def NumberFormatLinked(self, value):
        self.com_object.NumberFormatLinked = value

    @property
    def numberformatlinked(self):
        """Lower case alias for NumberFormatLinked"""
        return self.NumberFormatLinked

    @numberformatlinked.setter
    def numberformatlinked(self, value):
        """Lower case alias for NumberFormatLinked.setter"""
        self.NumberFormatLinked = value

    @property
    def NumberFormatLocal(self):
        return self.com_object.NumberFormatLocal

    @NumberFormatLocal.setter
    def NumberFormatLocal(self, value):
        self.com_object.NumberFormatLocal = value

    @property
    def numberformatlocal(self):
        """Lower case alias for NumberFormatLocal"""
        return self.NumberFormatLocal

    @numberformatlocal.setter
    def numberformatlocal(self, value):
        """Lower case alias for NumberFormatLocal.setter"""
        self.NumberFormatLocal = value

    @property
    def Orientation(self):
        return self.com_object.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.com_object.Orientation = value

    @property
    def orientation(self):
        """Lower case alias for Orientation"""
        return self.Orientation

    @orientation.setter
    def orientation(self, value):
        """Lower case alias for Orientation.setter"""
        self.Orientation = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def position(self):
        return self.com_object.position

    @position.setter
    def position(self, value):
        self.com_object.position = value

    @property
    def ReadingOrder(self):
        return XlReadingOrder(self.com_object.ReadingOrder)

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.com_object.ReadingOrder = value

    @property
    def readingorder(self):
        """Lower case alias for ReadingOrder"""
        return self.ReadingOrder

    @readingorder.setter
    def readingorder(self, value):
        """Lower case alias for ReadingOrder.setter"""
        self.ReadingOrder = value

    @property
    def Separator(self):
        return self.com_object.Separator

    @Separator.setter
    def Separator(self, value):
        self.com_object.Separator = value

    @property
    def separator(self):
        """Lower case alias for Separator"""
        return self.Separator

    @separator.setter
    def separator(self, value):
        """Lower case alias for Separator.setter"""
        self.Separator = value

    @property
    def Shadow(self):
        return self.com_object.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.com_object.Shadow = value

    @property
    def shadow(self):
        """Lower case alias for Shadow"""
        return self.Shadow

    @shadow.setter
    def shadow(self, value):
        """Lower case alias for Shadow.setter"""
        self.Shadow = value

    @property
    def ShowBubbleSize(self):
        return self.com_object.ShowBubbleSize

    @ShowBubbleSize.setter
    def ShowBubbleSize(self, value):
        self.com_object.ShowBubbleSize = value

    @property
    def showbubblesize(self):
        """Lower case alias for ShowBubbleSize"""
        return self.ShowBubbleSize

    @showbubblesize.setter
    def showbubblesize(self, value):
        """Lower case alias for ShowBubbleSize.setter"""
        self.ShowBubbleSize = value

    @property
    def ShowCategoryName(self):
        return self.com_object.ShowCategoryName

    @ShowCategoryName.setter
    def ShowCategoryName(self, value):
        self.com_object.ShowCategoryName = value

    @property
    def showcategoryname(self):
        """Lower case alias for ShowCategoryName"""
        return self.ShowCategoryName

    @showcategoryname.setter
    def showcategoryname(self, value):
        """Lower case alias for ShowCategoryName.setter"""
        self.ShowCategoryName = value

    @property
    def ShowLegendKey(self):
        return self.com_object.ShowLegendKey

    @ShowLegendKey.setter
    def ShowLegendKey(self, value):
        self.com_object.ShowLegendKey = value

    @property
    def showlegendkey(self):
        """Lower case alias for ShowLegendKey"""
        return self.ShowLegendKey

    @showlegendkey.setter
    def showlegendkey(self, value):
        """Lower case alias for ShowLegendKey.setter"""
        self.ShowLegendKey = value

    @property
    def ShowPercentage(self):
        return self.com_object.ShowPercentage

    @ShowPercentage.setter
    def ShowPercentage(self, value):
        self.com_object.ShowPercentage = value

    @property
    def showpercentage(self):
        """Lower case alias for ShowPercentage"""
        return self.ShowPercentage

    @showpercentage.setter
    def showpercentage(self, value):
        """Lower case alias for ShowPercentage.setter"""
        self.ShowPercentage = value

    @property
    def showrange(self):
        return self.com_object.showrange

    @showrange.setter
    def showrange(self, value):
        self.com_object.showrange = value

    @property
    def ShowSeriesName(self):
        return self.com_object.ShowSeriesName

    @ShowSeriesName.setter
    def ShowSeriesName(self, value):
        self.com_object.ShowSeriesName = value

    @property
    def showseriesname(self):
        """Lower case alias for ShowSeriesName"""
        return self.ShowSeriesName

    @showseriesname.setter
    def showseriesname(self, value):
        """Lower case alias for ShowSeriesName.setter"""
        self.ShowSeriesName = value

    @property
    def ShowValue(self):
        return self.com_object.ShowValue

    @ShowValue.setter
    def ShowValue(self, value):
        self.com_object.ShowValue = value

    @property
    def showvalue(self):
        """Lower case alias for ShowValue"""
        return self.ShowValue

    @showvalue.setter
    def showvalue(self, value):
        """Lower case alias for ShowValue.setter"""
        self.ShowValue = value

    @property
    def VerticalAlignment(self):
        return self.com_object.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.com_object.VerticalAlignment = value

    @property
    def verticalalignment(self):
        """Lower case alias for VerticalAlignment"""
        return self.VerticalAlignment

    @verticalalignment.setter
    def verticalalignment(self, value):
        """Lower case alias for VerticalAlignment.setter"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Border(self):
        return ChartBorder(self.com_object.Border)

    @property
    def border(self):
        """Lower case alias for Border"""
        return self.Border

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Font(self):
        return ChartFont(self.com_object.Font)

    @property
    def font(self):
        """Lower case alias for Font"""
        return self.Font

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

    @property
    def HasBorderHorizontal(self):
        return self.com_object.HasBorderHorizontal

    @HasBorderHorizontal.setter
    def HasBorderHorizontal(self, value):
        self.com_object.HasBorderHorizontal = value

    @property
    def hasborderhorizontal(self):
        """Lower case alias for HasBorderHorizontal"""
        return self.HasBorderHorizontal

    @hasborderhorizontal.setter
    def hasborderhorizontal(self, value):
        """Lower case alias for HasBorderHorizontal.setter"""
        self.HasBorderHorizontal = value

    @property
    def HasBorderOutline(self):
        return self.com_object.HasBorderOutline

    @HasBorderOutline.setter
    def HasBorderOutline(self, value):
        self.com_object.HasBorderOutline = value

    @property
    def hasborderoutline(self):
        """Lower case alias for HasBorderOutline"""
        return self.HasBorderOutline

    @hasborderoutline.setter
    def hasborderoutline(self, value):
        """Lower case alias for HasBorderOutline.setter"""
        self.HasBorderOutline = value

    @property
    def HasBorderVertical(self):
        return self.com_object.HasBorderVertical

    @HasBorderVertical.setter
    def HasBorderVertical(self, value):
        self.com_object.HasBorderVertical = value

    @property
    def hasbordervertical(self):
        """Lower case alias for HasBorderVertical"""
        return self.HasBorderVertical

    @hasbordervertical.setter
    def hasbordervertical(self, value):
        """Lower case alias for HasBorderVertical.setter"""
        self.HasBorderVertical = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def ShowLegendKey(self):
        return self.com_object.ShowLegendKey

    @ShowLegendKey.setter
    def ShowLegendKey(self, value):
        self.com_object.ShowLegendKey = value

    @property
    def showlegendkey(self):
        """Lower case alias for ShowLegendKey"""
        return self.ShowLegendKey

    @showlegendkey.setter
    def showlegendkey(self, value):
        """Lower case alias for ShowLegendKey.setter"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Index(self):
        return self.com_object.Index

    @property
    def index(self):
        """Lower case alias for Index"""
        return self.Index

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
    def Preserved(self):
        return self.com_object.Preserved

    @Preserved.setter
    def Preserved(self, value):
        self.com_object.Preserved = value

    @property
    def preserved(self):
        """Lower case alias for Preserved"""
        return self.Preserved

    @preserved.setter
    def preserved(self, value):
        """Lower case alias for Preserved.setter"""
        self.Preserved = value

    @property
    def SlideMaster(self):
        return Master(self.com_object.SlideMaster)

    @property
    def slidemaster(self):
        """Lower case alias for SlideMaster"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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

    def Characters(self, Start=None, Length=None):
        arguments = com_arguments([unwrap(a) for a in [Start, Length]])
        if hasattr(self.com_object, "GetCharacters"):
            return ChartCharacters(self.com_object.GetCharacters(*arguments))
        else:
            return ChartCharacters(self.com_object.Characters(*arguments))

    def characters(self, Start=None, Length=None):
        """Lower case alias for Characters"""
        arguments = [Start, Length]
        return self.Characters(*arguments)

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

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
    def FormulaLocal(self):
        return self.com_object.FormulaLocal

    @FormulaLocal.setter
    def FormulaLocal(self, value):
        self.com_object.FormulaLocal = value

    @property
    def formulalocal(self):
        """Lower case alias for FormulaLocal"""
        return self.FormulaLocal

    @formulalocal.setter
    def formulalocal(self, value):
        """Lower case alias for FormulaLocal.setter"""
        self.FormulaLocal = value

    @property
    def FormulaR1C1(self):
        return self.com_object.FormulaR1C1

    @FormulaR1C1.setter
    def FormulaR1C1(self, value):
        self.com_object.FormulaR1C1 = value

    @property
    def formular1c1(self):
        """Lower case alias for FormulaR1C1"""
        return self.FormulaR1C1

    @formular1c1.setter
    def formular1c1(self, value):
        """Lower case alias for FormulaR1C1.setter"""
        self.FormulaR1C1 = value

    @property
    def FormulaR1C1Local(self):
        return self.com_object.FormulaR1C1Local

    @FormulaR1C1Local.setter
    def FormulaR1C1Local(self, value):
        self.com_object.FormulaR1C1Local = value

    @property
    def formular1c1local(self):
        """Lower case alias for FormulaR1C1Local"""
        return self.FormulaR1C1Local

    @formular1c1local.setter
    def formular1c1local(self, value):
        """Lower case alias for FormulaR1C1Local.setter"""
        self.FormulaR1C1Local = value

    @property
    def Height(self):
        return self.com_object.Height

    @property
    def height(self):
        """Lower case alias for Height"""
        return self.Height

    @property
    def HorizontalAlignment(self):
        return self.com_object.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.com_object.HorizontalAlignment = value

    @property
    def horizontalalignment(self):
        """Lower case alias for HorizontalAlignment"""
        return self.HorizontalAlignment

    @horizontalalignment.setter
    def horizontalalignment(self, value):
        """Lower case alias for HorizontalAlignment.setter"""
        self.HorizontalAlignment = value

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
    def Name(self):
        return self.com_object.Name

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def Orientation(self):
        return self.com_object.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.com_object.Orientation = value

    @property
    def orientation(self):
        """Lower case alias for Orientation"""
        return self.Orientation

    @orientation.setter
    def orientation(self, value):
        """Lower case alias for Orientation.setter"""
        self.Orientation = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Position(self):
        return XlChartElementPosition(self.com_object.Position)

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
    def ReadingOrder(self):
        return XlReadingOrder(self.com_object.ReadingOrder)

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.com_object.ReadingOrder = value

    @property
    def readingorder(self):
        """Lower case alias for ReadingOrder"""
        return self.ReadingOrder

    @readingorder.setter
    def readingorder(self, value):
        """Lower case alias for ReadingOrder.setter"""
        self.ReadingOrder = value

    @property
    def Shadow(self):
        return self.com_object.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.com_object.Shadow = value

    @property
    def shadow(self):
        """Lower case alias for Shadow"""
        return self.Shadow

    @shadow.setter
    def shadow(self, value):
        """Lower case alias for Shadow.setter"""
        self.Shadow = value

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
    def VerticalAlignment(self):
        return self.com_object.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.com_object.VerticalAlignment = value

    @property
    def verticalalignment(self):
        """Lower case alias for VerticalAlignment"""
        return self.VerticalAlignment

    @verticalalignment.setter
    def verticalalignment(self, value):
        """Lower case alias for VerticalAlignment.setter"""
        self.VerticalAlignment = value

    @property
    def Width(self):
        return self.com_object.Width

    @property
    def width(self):
        """Lower case alias for Width"""
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

    @property
    def active(self):
        """Lower case alias for Active"""
        return self.Active

    @property
    def ActivePane(self):
        return Pane(self.com_object.ActivePane)

    @property
    def activepane(self):
        """Lower case alias for ActivePane"""
        return self.ActivePane

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def BlackAndWhite(self):
        return self.com_object.BlackAndWhite

    @BlackAndWhite.setter
    def BlackAndWhite(self, value):
        self.com_object.BlackAndWhite = value

    @property
    def blackandwhite(self):
        """Lower case alias for BlackAndWhite"""
        return self.BlackAndWhite

    @blackandwhite.setter
    def blackandwhite(self, value):
        """Lower case alias for BlackAndWhite.setter"""
        self.BlackAndWhite = value

    @property
    def Caption(self):
        return self.com_object.Caption

    @property
    def caption(self):
        """Lower case alias for Caption"""
        return self.Caption

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
    def Presentation(self):
        return Presentation(self.com_object.Presentation)

    @property
    def presentation(self):
        """Lower case alias for Presentation"""
        return self.Presentation

    @property
    def Selection(self):
        return Selection(self.com_object.Selection)

    @property
    def selection(self):
        """Lower case alias for Selection"""
        return self.Selection

    @property
    def SplitHorizontal(self):
        return self.com_object.SplitHorizontal

    @SplitHorizontal.setter
    def SplitHorizontal(self, value):
        self.com_object.SplitHorizontal = value

    @property
    def splithorizontal(self):
        """Lower case alias for SplitHorizontal"""
        return self.SplitHorizontal

    @splithorizontal.setter
    def splithorizontal(self, value):
        """Lower case alias for SplitHorizontal.setter"""
        self.SplitHorizontal = value

    @property
    def SplitVertical(self):
        return self.com_object.SplitVertical

    @SplitVertical.setter
    def SplitVertical(self, value):
        self.com_object.SplitVertical = value

    @property
    def splitvertical(self):
        """Lower case alias for SplitVertical"""
        return self.SplitVertical

    @splitvertical.setter
    def splitvertical(self, value):
        """Lower case alias for SplitVertical.setter"""
        self.SplitVertical = value

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
    def View(self):
        return View(self.com_object.View)

    @property
    def view(self):
        """Lower case alias for View"""
        return self.View

    @property
    def ViewType(self):
        return self.com_object.ViewType

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
        return self.com_object.WindowState

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Border(self):
        return ChartBorder(self.com_object.Border)

    @property
    def border(self):
        """Lower case alias for Border"""
        return self.Border

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Behaviors(self):
        return AnimationBehaviors(self.com_object.Behaviors)

    @property
    def behaviors(self):
        """Lower case alias for Behaviors"""
        return self.Behaviors

    @property
    def DisplayName(self):
        return self.com_object.DisplayName

    @property
    def displayname(self):
        """Lower case alias for DisplayName"""
        return self.DisplayName

    @property
    def EffectInformation(self):
        return EffectInformation(self.com_object.EffectInformation)

    @property
    def effectinformation(self):
        """Lower case alias for EffectInformation"""
        return self.EffectInformation

    @property
    def EffectParameters(self):
        return EffectParameters(self.com_object.EffectParameters)

    @property
    def effectparameters(self):
        """Lower case alias for EffectParameters"""
        return self.EffectParameters

    @property
    def EffectType(self):
        return self.com_object.EffectType

    @EffectType.setter
    def EffectType(self, value):
        self.com_object.EffectType = value

    @property
    def effecttype(self):
        """Lower case alias for EffectType"""
        return self.EffectType

    @effecttype.setter
    def effecttype(self, value):
        """Lower case alias for EffectType.setter"""
        self.EffectType = value

    @property
    def Exit(self):
        return self.com_object.Exit

    @Exit.setter
    def Exit(self, value):
        self.com_object.Exit = value

    @property
    def exit(self):
        """Lower case alias for Exit"""
        return self.Exit

    @exit.setter
    def exit(self, value):
        """Lower case alias for Exit.setter"""
        self.Exit = value

    @property
    def Index(self):
        return self.com_object.Index

    @property
    def index(self):
        """Lower case alias for Index"""
        return self.Index

    @property
    def Paragraph(self):
        return self.com_object.Paragraph

    @Paragraph.setter
    def Paragraph(self, value):
        self.com_object.Paragraph = value

    @property
    def paragraph(self):
        """Lower case alias for Paragraph"""
        return self.Paragraph

    @paragraph.setter
    def paragraph(self, value):
        """Lower case alias for Paragraph.setter"""
        self.Paragraph = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Shape(self):
        return Shape(self.com_object.Shape)

    @property
    def shape(self):
        """Lower case alias for Shape"""
        return self.Shape

    @property
    def TextRangeLength(self):
        return self.com_object.TextRangeLength

    @TextRangeLength.setter
    def TextRangeLength(self, value):
        self.com_object.TextRangeLength = value

    @property
    def textrangelength(self):
        """Lower case alias for TextRangeLength"""
        return self.TextRangeLength

    @textrangelength.setter
    def textrangelength(self, value):
        """Lower case alias for TextRangeLength.setter"""
        self.TextRangeLength = value

    @property
    def TextRangeStart(self):
        return self.com_object.TextRangeStart

    @TextRangeStart.setter
    def TextRangeStart(self, value):
        self.com_object.TextRangeStart = value

    @property
    def textrangestart(self):
        """Lower case alias for TextRangeStart"""
        return self.TextRangeStart

    @textrangestart.setter
    def textrangestart(self, value):
        """Lower case alias for TextRangeStart.setter"""
        self.TextRangeStart = value

    @property
    def Timing(self):
        return Timing(self.com_object.Timing)

    @property
    def timing(self):
        """Lower case alias for Timing"""
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

    @property
    def aftereffect(self):
        """Lower case alias for AfterEffect"""
        return self.AfterEffect

    @property
    def AnimateBackground(self):
        return self.com_object.AnimateBackground

    @property
    def animatebackground(self):
        """Lower case alias for AnimateBackground"""
        return self.AnimateBackground

    @property
    def AnimateTextInReverse(self):
        return self.com_object.AnimateTextInReverse

    @AnimateTextInReverse.setter
    def AnimateTextInReverse(self, value):
        self.com_object.AnimateTextInReverse = value

    @property
    def animatetextinreverse(self):
        """Lower case alias for AnimateTextInReverse"""
        return self.AnimateTextInReverse

    @animatetextinreverse.setter
    def animatetextinreverse(self, value):
        """Lower case alias for AnimateTextInReverse.setter"""
        self.AnimateTextInReverse = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def BuildByLevelEffect(self):
        return self.com_object.BuildByLevelEffect

    @property
    def buildbyleveleffect(self):
        """Lower case alias for BuildByLevelEffect"""
        return self.BuildByLevelEffect

    @property
    def Dim(self):
        return ColorFormat(self.com_object.Dim)

    @property
    def dim(self):
        """Lower case alias for Dim"""
        return self.Dim

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PlaySettings(self):
        return PlaySettings(self.com_object.PlaySettings)

    @property
    def playsettings(self):
        """Lower case alias for PlaySettings"""
        return self.PlaySettings

    @property
    def SoundEffect(self):
        return SoundEffect(self.com_object.SoundEffect)

    @property
    def soundeffect(self):
        """Lower case alias for SoundEffect"""
        return self.SoundEffect

    @property
    def TextUnitEffect(self):
        return self.com_object.TextUnitEffect

    @property
    def textuniteffect(self):
        """Lower case alias for TextUnitEffect"""
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

    @property
    def amount(self):
        """Lower case alias for Amount"""
        return self.Amount

    @amount.setter
    def amount(self, value):
        """Lower case alias for Amount.setter"""
        self.Amount = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Color2(self):
        return ColorFormat(self.com_object.Color2)

    @property
    def color2(self):
        """Lower case alias for Color2"""
        return self.Color2

    @property
    def Direction(self):
        return self.com_object.Direction

    @Direction.setter
    def Direction(self, value):
        self.com_object.Direction = value

    @property
    def direction(self):
        """Lower case alias for Direction"""
        return self.Direction

    @direction.setter
    def direction(self, value):
        """Lower case alias for Direction.setter"""
        self.Direction = value

    @property
    def FontName(self):
        return self.com_object.FontName

    @FontName.setter
    def FontName(self, value):
        self.com_object.FontName = value

    @property
    def fontname(self):
        """Lower case alias for FontName"""
        return self.FontName

    @fontname.setter
    def fontname(self, value):
        """Lower case alias for FontName.setter"""
        self.FontName = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Relative(self):
        return self.com_object.Relative

    @Relative.setter
    def Relative(self, value):
        self.com_object.Relative = value

    @property
    def relative(self):
        """Lower case alias for Relative"""
        return self.Relative

    @relative.setter
    def relative(self, value):
        """Lower case alias for Relative.setter"""
        self.Relative = value

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


class ErrorBars:

    def __init__(self, errorbars=None):
        self.com_object= errorbars

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Border(self):
        return ChartBorder(self.com_object.Border)

    @property
    def border(self):
        """Lower case alias for Border"""
        return self.Border

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def EndStyle(self):
        return self.com_object.EndStyle

    @EndStyle.setter
    def EndStyle(self, value):
        self.com_object.EndStyle = value

    @property
    def endstyle(self):
        """Lower case alias for EndStyle"""
        return self.EndStyle

    @endstyle.setter
    def endstyle(self, value):
        """Lower case alias for EndStyle.setter"""
        self.EndStyle = value

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def CanOpen(self):
        return self.com_object.CanOpen

    @property
    def canopen(self):
        """Lower case alias for CanOpen"""
        return self.CanOpen

    @property
    def CanSave(self):
        return self.com_object.CanSave

    @property
    def cansave(self):
        """Lower case alias for CanSave"""
        return self.CanSave

    @property
    def ClassName(self):
        return self.com_object.ClassName

    @property
    def classname(self):
        """Lower case alias for ClassName"""
        return self.ClassName

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Extensions(self):
        return FileConverter(self.com_object.Extensions)

    @property
    def extensions(self):
        """Lower case alias for Extensions"""
        return self.Extensions

    @property
    def FormatName(self):
        return self.com_object.FormatName

    @property
    def formatname(self):
        """Lower case alias for FormatName"""
        return self.FormatName

    @property
    def Name(self):
        return self.com_object.Name

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def OpenFormat(self):
        return self.com_object.OpenFormat

    @property
    def openformat(self):
        """Lower case alias for OpenFormat"""
        return self.OpenFormat

    @property
    def Parent(self):
        return FileConverter(self.com_object.Parent)

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Path(self):
        return self.com_object.Path

    @property
    def path(self):
        """Lower case alias for Path"""
        return self.Path

    @property
    def SaveFormat(self):
        return self.com_object.SaveFormat

    @property
    def saveformat(self):
        """Lower case alias for SaveFormat"""
        return self.SaveFormat


class FileConverters:

    def __init__(self, fileconverters=None):
        self.com_object= fileconverters

    def __call__(self, item):
        return FileConverter(self.com_object(item))

    @property
    def Count(self):
        return self.com_object.Count

    @property
    def count(self):
        """Lower case alias for Count"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def BackColor(self):
        return ColorFormat(self.com_object.BackColor)

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
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def ForeColor(self):
        return ColorFormat(self.com_object.ForeColor)

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
    def GradientAngle(self):
        return self.com_object.GradientAngle

    @GradientAngle.setter
    def GradientAngle(self, value):
        self.com_object.GradientAngle = value

    @property
    def gradientangle(self):
        """Lower case alias for GradientAngle"""
        return self.GradientAngle

    @gradientangle.setter
    def gradientangle(self, value):
        """Lower case alias for GradientAngle.setter"""
        self.GradientAngle = value

    @property
    def GradientColorType(self):
        return self.com_object.GradientColorType

    @property
    def gradientcolortype(self):
        """Lower case alias for GradientColorType"""
        return self.GradientColorType

    @property
    def GradientDegree(self):
        return self.com_object.GradientDegree

    @property
    def gradientdegree(self):
        """Lower case alias for GradientDegree"""
        return self.GradientDegree

    @property
    def GradientStops(self):
        return self.com_object.GradientStops

    @property
    def gradientstops(self):
        """Lower case alias for GradientStops"""
        return self.GradientStops

    @property
    def GradientStyle(self):
        return self.com_object.GradientStyle

    @property
    def gradientstyle(self):
        """Lower case alias for GradientStyle"""
        return self.GradientStyle

    @property
    def GradientVariant(self):
        return self.com_object.GradientVariant

    @property
    def gradientvariant(self):
        """Lower case alias for GradientVariant"""
        return self.GradientVariant

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Pattern(self):
        return self.com_object.Pattern

    @property
    def pattern(self):
        """Lower case alias for Pattern"""
        return self.Pattern

    @property
    def PictureEffects(self):
        return self.com_object.PictureEffects

    @property
    def pictureeffects(self):
        """Lower case alias for PictureEffects"""
        return self.PictureEffects

    @property
    def PresetGradientType(self):
        return self.com_object.PresetGradientType

    @property
    def presetgradienttype(self):
        """Lower case alias for PresetGradientType"""
        return self.PresetGradientType

    @property
    def PresetTexture(self):
        return self.com_object.PresetTexture

    @property
    def presettexture(self):
        """Lower case alias for PresetTexture"""
        return self.PresetTexture

    @property
    def RotateWithObject(self):
        return self.com_object.RotateWithObject

    @RotateWithObject.setter
    def RotateWithObject(self, value):
        self.com_object.RotateWithObject = value

    @property
    def rotatewithobject(self):
        """Lower case alias for RotateWithObject"""
        return self.RotateWithObject

    @rotatewithobject.setter
    def rotatewithobject(self, value):
        """Lower case alias for RotateWithObject.setter"""
        self.RotateWithObject = value

    @property
    def TextureAlignment(self):
        return self.com_object.TextureAlignment

    @TextureAlignment.setter
    def TextureAlignment(self, value):
        self.com_object.TextureAlignment = value

    @property
    def texturealignment(self):
        """Lower case alias for TextureAlignment"""
        return self.TextureAlignment

    @texturealignment.setter
    def texturealignment(self, value):
        """Lower case alias for TextureAlignment.setter"""
        self.TextureAlignment = value

    @property
    def TextureHorizontalScale(self):
        return self.com_object.TextureHorizontalScale

    @TextureHorizontalScale.setter
    def TextureHorizontalScale(self, value):
        self.com_object.TextureHorizontalScale = value

    @property
    def texturehorizontalscale(self):
        """Lower case alias for TextureHorizontalScale"""
        return self.TextureHorizontalScale

    @texturehorizontalscale.setter
    def texturehorizontalscale(self, value):
        """Lower case alias for TextureHorizontalScale.setter"""
        self.TextureHorizontalScale = value

    @property
    def TextureName(self):
        return self.com_object.TextureName

    @property
    def texturename(self):
        """Lower case alias for TextureName"""
        return self.TextureName

    @property
    def TextureOffsetX(self):
        return self.com_object.TextureOffsetX

    @TextureOffsetX.setter
    def TextureOffsetX(self, value):
        self.com_object.TextureOffsetX = value

    @property
    def textureoffsetx(self):
        """Lower case alias for TextureOffsetX"""
        return self.TextureOffsetX

    @textureoffsetx.setter
    def textureoffsetx(self, value):
        """Lower case alias for TextureOffsetX.setter"""
        self.TextureOffsetX = value

    @property
    def TextureOffsetY(self):
        return self.com_object.TextureOffsetY

    @TextureOffsetY.setter
    def TextureOffsetY(self, value):
        self.com_object.TextureOffsetY = value

    @property
    def textureoffsety(self):
        """Lower case alias for TextureOffsetY"""
        return self.TextureOffsetY

    @textureoffsety.setter
    def textureoffsety(self, value):
        """Lower case alias for TextureOffsetY.setter"""
        self.TextureOffsetY = value

    @property
    def TextureTile(self):
        return self.com_object.TextureTile

    @TextureTile.setter
    def TextureTile(self, value):
        self.com_object.TextureTile = value

    @property
    def texturetile(self):
        """Lower case alias for TextureTile"""
        return self.TextureTile

    @texturetile.setter
    def texturetile(self, value):
        """Lower case alias for TextureTile.setter"""
        self.TextureTile = value

    @property
    def TextureType(self):
        return self.com_object.TextureType

    @property
    def texturetype(self):
        """Lower case alias for TextureType"""
        return self.TextureType

    @property
    def TextureVerticalScale(self):
        return self.com_object.TextureVerticalScale

    @TextureVerticalScale.setter
    def TextureVerticalScale(self, value):
        self.com_object.TextureVerticalScale = value

    @property
    def textureverticalscale(self):
        """Lower case alias for TextureVerticalScale"""
        return self.TextureVerticalScale

    @textureverticalscale.setter
    def textureverticalscale(self, value):
        """Lower case alias for TextureVerticalScale.setter"""
        self.TextureVerticalScale = value

    @property
    def Transparency(self):
        return self.com_object.Transparency

    @Transparency.setter
    def Transparency(self, value):
        self.com_object.Transparency = value

    @property
    def transparency(self):
        """Lower case alias for Transparency"""
        return self.Transparency

    @transparency.setter
    def transparency(self, value):
        """Lower case alias for Transparency.setter"""
        self.Transparency = value

    @property
    def Type(self):
        return self.com_object.Type

    @property
    def type(self):
        """Lower case alias for Type"""
        return self.Type

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Reveal(self):
        return self.com_object.Reveal

    @Reveal.setter
    def Reveal(self, value):
        self.com_object.Reveal = value

    @property
    def reveal(self):
        """Lower case alias for Reveal"""
        return self.Reveal

    @reveal.setter
    def reveal(self, value):
        """Lower case alias for Reveal.setter"""
        self.Reveal = value

    @property
    def Subtype(self):
        return self.com_object.Subtype

    @Subtype.setter
    def Subtype(self, value):
        self.com_object.Subtype = value

    @property
    def subtype(self):
        """Lower case alias for Subtype"""
        return self.Subtype

    @subtype.setter
    def subtype(self, value):
        """Lower case alias for Subtype.setter"""
        self.Subtype = value

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


class Floor:

    def __init__(self, floor=None):
        self.com_object= floor

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

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
    def PictureType(self):
        return self.com_object.PictureType

    @PictureType.setter
    def PictureType(self, value):
        self.com_object.PictureType = value

    @property
    def picturetype(self):
        """Lower case alias for PictureType"""
        return self.PictureType

    @picturetype.setter
    def picturetype(self, value):
        """Lower case alias for PictureType.setter"""
        self.PictureType = value

    @property
    def Thickness(self):
        return self.com_object.Thickness

    @Thickness.setter
    def Thickness(self, value):
        self.com_object.Thickness = value

    @property
    def thickness(self):
        """Lower case alias for Thickness"""
        return self.Thickness

    @thickness.setter
    def thickness(self, value):
        """Lower case alias for Thickness.setter"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def AutoRotateNumbers(self):
        return self.com_object.AutoRotateNumbers

    @AutoRotateNumbers.setter
    def AutoRotateNumbers(self, value):
        self.com_object.AutoRotateNumbers = value

    @property
    def autorotatenumbers(self):
        """Lower case alias for AutoRotateNumbers"""
        return self.AutoRotateNumbers

    @autorotatenumbers.setter
    def autorotatenumbers(self, value):
        """Lower case alias for AutoRotateNumbers.setter"""
        self.AutoRotateNumbers = value

    @property
    def BaselineOffset(self):
        return self.com_object.BaselineOffset

    @BaselineOffset.setter
    def BaselineOffset(self, value):
        self.com_object.BaselineOffset = value

    @property
    def baselineoffset(self):
        """Lower case alias for BaselineOffset"""
        return self.BaselineOffset

    @baselineoffset.setter
    def baselineoffset(self, value):
        """Lower case alias for BaselineOffset.setter"""
        self.BaselineOffset = value

    @property
    def Bold(self):
        return self.com_object.Bold

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
    def Color(self):
        return Font(self.com_object.Color)

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
    def Embeddable(self):
        return self.com_object.Embeddable

    @property
    def embeddable(self):
        """Lower case alias for Embeddable"""
        return self.Embeddable

    @property
    def Embedded(self):
        return self.com_object.Embedded

    @property
    def embedded(self):
        """Lower case alias for Embedded"""
        return self.Embedded

    @property
    def Emboss(self):
        return self.com_object.Emboss

    @Emboss.setter
    def Emboss(self, value):
        self.com_object.Emboss = value

    @property
    def emboss(self):
        """Lower case alias for Emboss"""
        return self.Emboss

    @emboss.setter
    def emboss(self, value):
        """Lower case alias for Emboss.setter"""
        self.Emboss = value

    @property
    def italic(self):
        return self.com_object.italic

    @italic.setter
    def italic(self, value):
        self.com_object.italic = value

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
    def NameAscii(self):
        return self.com_object.NameAscii

    @NameAscii.setter
    def NameAscii(self, value):
        self.com_object.NameAscii = value

    @property
    def nameascii(self):
        """Lower case alias for NameAscii"""
        return self.NameAscii

    @nameascii.setter
    def nameascii(self, value):
        """Lower case alias for NameAscii.setter"""
        self.NameAscii = value

    @property
    def NameComplexScript(self):
        return self.com_object.NameComplexScript

    @NameComplexScript.setter
    def NameComplexScript(self, value):
        self.com_object.NameComplexScript = value

    @property
    def namecomplexscript(self):
        """Lower case alias for NameComplexScript"""
        return self.NameComplexScript

    @namecomplexscript.setter
    def namecomplexscript(self, value):
        """Lower case alias for NameComplexScript.setter"""
        self.NameComplexScript = value

    @property
    def NameFarEast(self):
        return self.com_object.NameFarEast

    @NameFarEast.setter
    def NameFarEast(self, value):
        self.com_object.NameFarEast = value

    @property
    def namefareast(self):
        """Lower case alias for NameFarEast"""
        return self.NameFarEast

    @namefareast.setter
    def namefareast(self, value):
        """Lower case alias for NameFarEast.setter"""
        self.NameFarEast = value

    @property
    def NameOther(self):
        return self.com_object.NameOther

    @NameOther.setter
    def NameOther(self, value):
        self.com_object.NameOther = value

    @property
    def nameother(self):
        """Lower case alias for NameOther"""
        return self.NameOther

    @nameother.setter
    def nameother(self, value):
        """Lower case alias for NameOther.setter"""
        self.NameOther = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Shadow(self):
        return self.com_object.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.com_object.Shadow = value

    @property
    def shadow(self):
        """Lower case alias for Shadow"""
        return self.Shadow

    @shadow.setter
    def shadow(self, value):
        """Lower case alias for Shadow.setter"""
        self.Shadow = value

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
    def Subscript(self):
        return self.com_object.Subscript

    @Subscript.setter
    def Subscript(self, value):
        self.com_object.Subscript = value

    @property
    def subscript(self):
        """Lower case alias for Subscript"""
        return self.Subscript

    @subscript.setter
    def subscript(self, value):
        """Lower case alias for Subscript.setter"""
        self.Subscript = value

    @property
    def Superscript(self):
        return self.com_object.Superscript

    @Superscript.setter
    def Superscript(self, value):
        self.com_object.Superscript = value

    @property
    def superscript(self):
        """Lower case alias for Superscript"""
        return self.Superscript

    @superscript.setter
    def superscript(self, value):
        """Lower case alias for Superscript.setter"""
        self.Superscript = value

    @property
    def Underline(self):
        return self.com_object.Underline

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


class Fonts:

    def __init__(self, fonts=None):
        self.com_object= fonts

    def __call__(self, item):
        return Font(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
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


class fullseriescollection:

    def __init__(self, fullseriescollection=None):
        self.com_object= fullseriescollection

    @property
    def application(self):
        return self.com_object.application

    @property
    def count(self):
        return self.com_object.count

    @property
    def creator(self):
        return self.com_object.creator

    @property
    def parent(self):
        return self.com_object.parent


class GridLines:

    def __init__(self, gridlines=None):
        self.com_object= gridlines

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Border(self):
        return ChartBorder(self.com_object.Border)

    @property
    def border(self):
        """Lower case alias for Border"""
        return self.Border

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def format(self):
        return ChartFormat(self.com_object.format)

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Count(self):
        return self.com_object.Count

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
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


class guide:

    def __init__(self, guide=None):
        self.com_object= guide

    @property
    def application(self):
        return self.com_object.application

    @property
    def color(self):
        return self.com_object.color

    @property
    def orientation(self):
        return self.com_object.orientation

    @property
    def parent(self):
        return self.com_object.parent

    @property
    def position(self):
        return self.com_object.position

    @position.setter
    def position(self, value):
        self.com_object.position = value


class guides:

    def __init__(self, guides=None):
        self.com_object= guides

    def __call__(self, item):
        return guide(self.com_object(item))

    @property
    def application(self):
        return self.com_object.application

    @property
    def count(self):
        return self.com_object.count

    @property
    def parent(self):
        return self.com_object.parent


class HeaderFooter:

    def __init__(self, headerfooter=None):
        self.com_object= headerfooter

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

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
    def UseFormat(self):
        return self.com_object.UseFormat

    @UseFormat.setter
    def UseFormat(self, value):
        self.com_object.UseFormat = value

    @property
    def useformat(self):
        """Lower case alias for UseFormat"""
        return self.UseFormat

    @useformat.setter
    def useformat(self, value):
        """Lower case alias for UseFormat.setter"""
        self.UseFormat = value

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


class HeadersFooters:

    def __init__(self, headersfooters=None):
        self.com_object= headersfooters

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def DateAndTime(self):
        return HeaderFooter(self.com_object.DateAndTime)

    @property
    def dateandtime(self):
        """Lower case alias for DateAndTime"""
        return self.DateAndTime

    @property
    def DisplayOnTitleSlide(self):
        return self.com_object.DisplayOnTitleSlide

    @DisplayOnTitleSlide.setter
    def DisplayOnTitleSlide(self, value):
        self.com_object.DisplayOnTitleSlide = value

    @property
    def displayontitleslide(self):
        """Lower case alias for DisplayOnTitleSlide"""
        return self.DisplayOnTitleSlide

    @displayontitleslide.setter
    def displayontitleslide(self, value):
        """Lower case alias for DisplayOnTitleSlide.setter"""
        self.DisplayOnTitleSlide = value

    @property
    def Footer(self):
        return HeaderFooter(self.com_object.Footer)

    @property
    def footer(self):
        """Lower case alias for Footer"""
        return self.Footer

    @property
    def Header(self):
        return HeaderFooter(self.com_object.Header)

    @property
    def header(self):
        """Lower case alias for Header"""
        return self.Header

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def SlideNumber(self):
        return HeaderFooter(self.com_object.SlideNumber)

    @property
    def slidenumber(self):
        """Lower case alias for SlideNumber"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Border(self):
        return ChartBorder(self.com_object.Border)

    @property
    def border(self):
        """Lower case alias for Border"""
        return self.Border

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def EmailSubject(self):
        return self.com_object.EmailSubject

    @EmailSubject.setter
    def EmailSubject(self, value):
        self.com_object.EmailSubject = value

    @property
    def emailsubject(self):
        """Lower case alias for EmailSubject"""
        return self.EmailSubject

    @emailsubject.setter
    def emailsubject(self, value):
        """Lower case alias for EmailSubject.setter"""
        self.EmailSubject = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def ScreenTip(self):
        return self.com_object.ScreenTip

    @ScreenTip.setter
    def ScreenTip(self, value):
        self.com_object.ScreenTip = value

    @property
    def screentip(self):
        """Lower case alias for ScreenTip"""
        return self.ScreenTip

    @screentip.setter
    def screentip(self, value):
        """Lower case alias for ScreenTip.setter"""
        self.ScreenTip = value

    @property
    def ShowAndReturn(self):
        return self.com_object.ShowAndReturn

    @ShowAndReturn.setter
    def ShowAndReturn(self, value):
        self.com_object.ShowAndReturn = value

    @property
    def showandreturn(self):
        """Lower case alias for ShowAndReturn"""
        return self.ShowAndReturn

    @showandreturn.setter
    def showandreturn(self, value):
        """Lower case alias for ShowAndReturn.setter"""
        self.ShowAndReturn = value

    @property
    def SubAddress(self):
        return self.com_object.SubAddress

    @SubAddress.setter
    def SubAddress(self, value):
        self.com_object.SubAddress = value

    @property
    def subaddress(self):
        """Lower case alias for SubAddress"""
        return self.SubAddress

    @subaddress.setter
    def subaddress(self, value):
        """Lower case alias for SubAddress.setter"""
        self.SubAddress = value

    @property
    def TextToDisplay(self):
        return self.com_object.TextToDisplay

    @TextToDisplay.setter
    def TextToDisplay(self, value):
        self.com_object.TextToDisplay = value

    @property
    def texttodisplay(self):
        """Lower case alias for TextToDisplay"""
        return self.TextToDisplay

    @texttodisplay.setter
    def texttodisplay(self, value):
        """Lower case alias for TextToDisplay.setter"""
        self.TextToDisplay = value

    @property
    def Type(self):
        return self.com_object.Type

    @property
    def type(self):
        """Lower case alias for Type"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Color(self):
        return self.com_object.Color

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
    def ColorIndex(self):
        return self.com_object.ColorIndex

    @ColorIndex.setter
    def ColorIndex(self, value):
        self.com_object.ColorIndex = value

    @property
    def colorindex(self):
        """Lower case alias for ColorIndex"""
        return self.ColorIndex

    @colorindex.setter
    def colorindex(self, value):
        """Lower case alias for ColorIndex.setter"""
        self.ColorIndex = value

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def InvertIfNegative(self):
        return self.com_object.InvertIfNegative

    @InvertIfNegative.setter
    def InvertIfNegative(self, value):
        self.com_object.InvertIfNegative = value

    @property
    def invertifnegative(self):
        """Lower case alias for InvertIfNegative"""
        return self.InvertIfNegative

    @invertifnegative.setter
    def invertifnegative(self, value):
        """Lower case alias for InvertIfNegative.setter"""
        self.InvertIfNegative = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Pattern(self):
        return XlPattern(self.com_object.Pattern)

    @Pattern.setter
    def Pattern(self, value):
        self.com_object.Pattern = value

    @property
    def pattern(self):
        """Lower case alias for Pattern"""
        return self.Pattern

    @pattern.setter
    def pattern(self, value):
        """Lower case alias for Pattern.setter"""
        self.Pattern = value

    @property
    def PatternColor(self):
        return self.com_object.PatternColor

    @PatternColor.setter
    def PatternColor(self, value):
        self.com_object.PatternColor = value

    @property
    def patterncolor(self):
        """Lower case alias for PatternColor"""
        return self.PatternColor

    @patterncolor.setter
    def patterncolor(self, value):
        """Lower case alias for PatternColor.setter"""
        self.PatternColor = value

    @property
    def PatternColorIndex(self):
        return XlColorIndex(self.com_object.PatternColorIndex)

    @PatternColorIndex.setter
    def PatternColorIndex(self, value):
        self.com_object.PatternColorIndex = value

    @property
    def patterncolorindex(self):
        """Lower case alias for PatternColorIndex"""
        return self.PatternColorIndex

    @patterncolorindex.setter
    def patterncolorindex(self, value):
        """Lower case alias for PatternColorIndex.setter"""
        self.PatternColorIndex = value


class LeaderLines:

    def __init__(self, leaderlines=None):
        self.com_object= leaderlines

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Border(self):
        return ChartBorder(self.com_object.Border)

    @property
    def border(self):
        """Lower case alias for Border"""
        return self.Border

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

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
    def IncludeInLayout(self):
        return self.com_object.IncludeInLayout

    @IncludeInLayout.setter
    def IncludeInLayout(self, value):
        self.com_object.IncludeInLayout = value

    @property
    def includeinlayout(self):
        """Lower case alias for IncludeInLayout"""
        return self.IncludeInLayout

    @includeinlayout.setter
    def includeinlayout(self, value):
        """Lower case alias for IncludeInLayout.setter"""
        self.IncludeInLayout = value

    @property
    def Left(self):
        return self.com_object.Left

    @property
    def left(self):
        """Lower case alias for Left"""
        return self.Left

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
    def Position(self):
        return XlLegendPosition(self.com_object.Position)

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
    def Shadow(self):
        return self.com_object.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.com_object.Shadow = value

    @property
    def shadow(self):
        """Lower case alias for Shadow"""
        return self.Shadow

    @shadow.setter
    def shadow(self, value):
        """Lower case alias for Shadow.setter"""
        self.Shadow = value

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Count(self):
        return self.com_object.Count

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Font(self):
        return ChartFont(self.com_object.Font)

    @property
    def font(self):
        """Lower case alias for Font"""
        return self.Font

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

    @property
    def Height(self):
        return self.com_object.Height

    @property
    def height(self):
        """Lower case alias for Height"""
        return self.Height

    @property
    def Index(self):
        return self.com_object.Index

    @property
    def index(self):
        """Lower case alias for Index"""
        return self.Index

    @property
    def Left(self):
        return self.com_object.Left

    @property
    def left(self):
        """Lower case alias for Left"""
        return self.Left

    @property
    def LegendKey(self):
        return LegendKey(self.com_object.LegendKey)

    @property
    def legendkey(self):
        """Lower case alias for LegendKey"""
        return self.LegendKey

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Top(self):
        return self.com_object.Top

    @property
    def top(self):
        """Lower case alias for Top"""
        return self.Top

    @property
    def Width(self):
        return self.com_object.Width

    @property
    def width(self):
        """Lower case alias for Width"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

    @property
    def Height(self):
        return self.com_object.Height

    @property
    def height(self):
        """Lower case alias for Height"""
        return self.Height

    @property
    def InvertIfNegative(self):
        return self.com_object.InvertIfNegative

    @InvertIfNegative.setter
    def InvertIfNegative(self, value):
        self.com_object.InvertIfNegative = value

    @property
    def invertifnegative(self):
        """Lower case alias for InvertIfNegative"""
        return self.InvertIfNegative

    @invertifnegative.setter
    def invertifnegative(self, value):
        """Lower case alias for InvertIfNegative.setter"""
        self.InvertIfNegative = value

    @property
    def Left(self):
        return self.com_object.Left

    @property
    def left(self):
        """Lower case alias for Left"""
        return self.Left

    @property
    def MarkerBackgroundColor(self):
        return self.com_object.MarkerBackgroundColor

    @MarkerBackgroundColor.setter
    def MarkerBackgroundColor(self, value):
        self.com_object.MarkerBackgroundColor = value

    @property
    def markerbackgroundcolor(self):
        """Lower case alias for MarkerBackgroundColor"""
        return self.MarkerBackgroundColor

    @markerbackgroundcolor.setter
    def markerbackgroundcolor(self, value):
        """Lower case alias for MarkerBackgroundColor.setter"""
        self.MarkerBackgroundColor = value

    @property
    def MarkerBackgroundColorIndex(self):
        return XlColorIndex(self.com_object.MarkerBackgroundColorIndex)

    @MarkerBackgroundColorIndex.setter
    def MarkerBackgroundColorIndex(self, value):
        self.com_object.MarkerBackgroundColorIndex = value

    @property
    def markerbackgroundcolorindex(self):
        """Lower case alias for MarkerBackgroundColorIndex"""
        return self.MarkerBackgroundColorIndex

    @markerbackgroundcolorindex.setter
    def markerbackgroundcolorindex(self, value):
        """Lower case alias for MarkerBackgroundColorIndex.setter"""
        self.MarkerBackgroundColorIndex = value

    @property
    def MarkerForegroundColor(self):
        return self.com_object.MarkerForegroundColor

    @MarkerForegroundColor.setter
    def MarkerForegroundColor(self, value):
        self.com_object.MarkerForegroundColor = value

    @property
    def markerforegroundcolor(self):
        """Lower case alias for MarkerForegroundColor"""
        return self.MarkerForegroundColor

    @markerforegroundcolor.setter
    def markerforegroundcolor(self, value):
        """Lower case alias for MarkerForegroundColor.setter"""
        self.MarkerForegroundColor = value

    @property
    def MarkerForegroundColorIndex(self):
        return XlColorIndex(self.com_object.MarkerForegroundColorIndex)

    @MarkerForegroundColorIndex.setter
    def MarkerForegroundColorIndex(self, value):
        self.com_object.MarkerForegroundColorIndex = value

    @property
    def markerforegroundcolorindex(self):
        """Lower case alias for MarkerForegroundColorIndex"""
        return self.MarkerForegroundColorIndex

    @markerforegroundcolorindex.setter
    def markerforegroundcolorindex(self, value):
        """Lower case alias for MarkerForegroundColorIndex.setter"""
        self.MarkerForegroundColorIndex = value

    @property
    def MarkerSize(self):
        return self.com_object.MarkerSize

    @MarkerSize.setter
    def MarkerSize(self, value):
        self.com_object.MarkerSize = value

    @property
    def markersize(self):
        """Lower case alias for MarkerSize"""
        return self.MarkerSize

    @markersize.setter
    def markersize(self, value):
        """Lower case alias for MarkerSize.setter"""
        self.MarkerSize = value

    @property
    def MarkerStyle(self):
        return XlMarkerStyle(self.com_object.MarkerStyle)

    @MarkerStyle.setter
    def MarkerStyle(self, value):
        self.com_object.MarkerStyle = value

    @property
    def markerstyle(self):
        """Lower case alias for MarkerStyle"""
        return self.MarkerStyle

    @markerstyle.setter
    def markerstyle(self, value):
        """Lower case alias for MarkerStyle.setter"""
        self.MarkerStyle = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PictureType(self):
        return XlChartPictureType(self.com_object.PictureType)

    @PictureType.setter
    def PictureType(self, value):
        self.com_object.PictureType = value

    @property
    def picturetype(self):
        """Lower case alias for PictureType"""
        return self.PictureType

    @picturetype.setter
    def picturetype(self, value):
        """Lower case alias for PictureType.setter"""
        self.PictureType = value

    @property
    def PictureUnit2(self):
        return self.com_object.PictureUnit2

    @PictureUnit2.setter
    def PictureUnit2(self, value):
        self.com_object.PictureUnit2 = value

    @property
    def pictureunit2(self):
        """Lower case alias for PictureUnit2"""
        return self.PictureUnit2

    @pictureunit2.setter
    def pictureunit2(self, value):
        """Lower case alias for PictureUnit2.setter"""
        self.PictureUnit2 = value

    @property
    def Shadow(self):
        return self.com_object.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.com_object.Shadow = value

    @property
    def shadow(self):
        """Lower case alias for Shadow"""
        return self.Shadow

    @shadow.setter
    def shadow(self, value):
        """Lower case alias for Shadow.setter"""
        self.Shadow = value

    @property
    def Smooth(self):
        return self.com_object.Smooth

    @Smooth.setter
    def Smooth(self, value):
        self.com_object.Smooth = value

    @property
    def smooth(self):
        """Lower case alias for Smooth"""
        return self.Smooth

    @smooth.setter
    def smooth(self, value):
        """Lower case alias for Smooth.setter"""
        self.Smooth = value

    @property
    def Top(self):
        return self.com_object.Top

    @property
    def top(self):
        """Lower case alias for Top"""
        return self.Top

    @property
    def Width(self):
        return self.com_object.Width

    @property
    def width(self):
        """Lower case alias for Width"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def BackColor(self):
        return ColorFormat(self.com_object.BackColor)

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
    def BeginArrowheadLength(self):
        return self.com_object.BeginArrowheadLength

    @BeginArrowheadLength.setter
    def BeginArrowheadLength(self, value):
        self.com_object.BeginArrowheadLength = value

    @property
    def beginarrowheadlength(self):
        """Lower case alias for BeginArrowheadLength"""
        return self.BeginArrowheadLength

    @beginarrowheadlength.setter
    def beginarrowheadlength(self, value):
        """Lower case alias for BeginArrowheadLength.setter"""
        self.BeginArrowheadLength = value

    @property
    def BeginArrowheadStyle(self):
        return self.com_object.BeginArrowheadStyle

    @BeginArrowheadStyle.setter
    def BeginArrowheadStyle(self, value):
        self.com_object.BeginArrowheadStyle = value

    @property
    def beginarrowheadstyle(self):
        """Lower case alias for BeginArrowheadStyle"""
        return self.BeginArrowheadStyle

    @beginarrowheadstyle.setter
    def beginarrowheadstyle(self, value):
        """Lower case alias for BeginArrowheadStyle.setter"""
        self.BeginArrowheadStyle = value

    @property
    def BeginArrowheadWidth(self):
        return self.com_object.BeginArrowheadWidth

    @BeginArrowheadWidth.setter
    def BeginArrowheadWidth(self, value):
        self.com_object.BeginArrowheadWidth = value

    @property
    def beginarrowheadwidth(self):
        """Lower case alias for BeginArrowheadWidth"""
        return self.BeginArrowheadWidth

    @beginarrowheadwidth.setter
    def beginarrowheadwidth(self, value):
        """Lower case alias for BeginArrowheadWidth.setter"""
        self.BeginArrowheadWidth = value

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def DashStyle(self):
        return self.com_object.DashStyle

    @DashStyle.setter
    def DashStyle(self, value):
        self.com_object.DashStyle = value

    @property
    def dashstyle(self):
        """Lower case alias for DashStyle"""
        return self.DashStyle

    @dashstyle.setter
    def dashstyle(self, value):
        """Lower case alias for DashStyle.setter"""
        self.DashStyle = value

    @property
    def EndArrowheadLength(self):
        return self.com_object.EndArrowheadLength

    @EndArrowheadLength.setter
    def EndArrowheadLength(self, value):
        self.com_object.EndArrowheadLength = value

    @property
    def endarrowheadlength(self):
        """Lower case alias for EndArrowheadLength"""
        return self.EndArrowheadLength

    @endarrowheadlength.setter
    def endarrowheadlength(self, value):
        """Lower case alias for EndArrowheadLength.setter"""
        self.EndArrowheadLength = value

    @property
    def EndArrowheadStyle(self):
        return self.com_object.EndArrowheadStyle

    @EndArrowheadStyle.setter
    def EndArrowheadStyle(self, value):
        self.com_object.EndArrowheadStyle = value

    @property
    def endarrowheadstyle(self):
        """Lower case alias for EndArrowheadStyle"""
        return self.EndArrowheadStyle

    @endarrowheadstyle.setter
    def endarrowheadstyle(self, value):
        """Lower case alias for EndArrowheadStyle.setter"""
        self.EndArrowheadStyle = value

    @property
    def EndArrowheadWidth(self):
        return self.com_object.EndArrowheadWidth

    @EndArrowheadWidth.setter
    def EndArrowheadWidth(self, value):
        self.com_object.EndArrowheadWidth = value

    @property
    def endarrowheadwidth(self):
        """Lower case alias for EndArrowheadWidth"""
        return self.EndArrowheadWidth

    @endarrowheadwidth.setter
    def endarrowheadwidth(self, value):
        """Lower case alias for EndArrowheadWidth.setter"""
        self.EndArrowheadWidth = value

    @property
    def ForeColor(self):
        return ColorFormat(self.com_object.ForeColor)

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
    def InsetPen(self):
        return self.com_object.InsetPen

    @InsetPen.setter
    def InsetPen(self, value):
        self.com_object.InsetPen = value

    @property
    def insetpen(self):
        """Lower case alias for InsetPen"""
        return self.InsetPen

    @insetpen.setter
    def insetpen(self, value):
        """Lower case alias for InsetPen.setter"""
        self.InsetPen = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Pattern(self):
        return self.com_object.Pattern

    @Pattern.setter
    def Pattern(self, value):
        self.com_object.Pattern = value

    @property
    def pattern(self):
        """Lower case alias for Pattern"""
        return self.Pattern

    @pattern.setter
    def pattern(self, value):
        """Lower case alias for Pattern.setter"""
        self.Pattern = value

    @property
    def Style(self):
        return self.com_object.Style

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
    def Transparency(self):
        return self.com_object.Transparency

    @Transparency.setter
    def Transparency(self, value):
        self.com_object.Transparency = value

    @property
    def transparency(self):
        """Lower case alias for Transparency"""
        return self.Transparency

    @transparency.setter
    def transparency(self, value):
        """Lower case alias for Transparency.setter"""
        self.Transparency = value

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

    @property
    def Weight(self):
        return self.com_object.Weight

    @Weight.setter
    def Weight(self, value):
        self.com_object.Weight = value

    @property
    def weight(self):
        """Lower case alias for Weight"""
        return self.Weight

    @weight.setter
    def weight(self, value):
        """Lower case alias for Weight.setter"""
        self.Weight = value


class LinkFormat:

    def __init__(self, linkformat=None):
        self.com_object= linkformat

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def AutoUpdate(self):
        return self.com_object.AutoUpdate

    @AutoUpdate.setter
    def AutoUpdate(self, value):
        self.com_object.AutoUpdate = value

    @property
    def autoupdate(self):
        """Lower case alias for AutoUpdate"""
        return self.AutoUpdate

    @autoupdate.setter
    def autoupdate(self, value):
        """Lower case alias for AutoUpdate.setter"""
        self.AutoUpdate = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def SourceFullName(self):
        return self.com_object.SourceFullName

    @SourceFullName.setter
    def SourceFullName(self, value):
        self.com_object.SourceFullName = value

    @property
    def sourcefullname(self):
        """Lower case alias for SourceFullName"""
        return self.SourceFullName

    @sourcefullname.setter
    def sourcefullname(self, value):
        """Lower case alias for SourceFullName.setter"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Background(self):
        return ShapeRange(self.com_object.Background)

    @property
    def background(self):
        """Lower case alias for Background"""
        return self.Background

    @property
    def BackgroundStyle(self):
        return self.com_object.BackgroundStyle

    @BackgroundStyle.setter
    def BackgroundStyle(self, value):
        self.com_object.BackgroundStyle = value

    @property
    def backgroundstyle(self):
        """Lower case alias for BackgroundStyle"""
        return self.BackgroundStyle

    @backgroundstyle.setter
    def backgroundstyle(self, value):
        """Lower case alias for BackgroundStyle.setter"""
        self.BackgroundStyle = value

    @property
    def ColorScheme(self):
        return ColorScheme(self.com_object.ColorScheme)

    @ColorScheme.setter
    def ColorScheme(self, value):
        self.com_object.ColorScheme = value

    @property
    def colorscheme(self):
        """Lower case alias for ColorScheme"""
        return self.ColorScheme

    @colorscheme.setter
    def colorscheme(self, value):
        """Lower case alias for ColorScheme.setter"""
        self.ColorScheme = value

    @property
    def CustomerData(self):
        return CustomerData(self.com_object.CustomerData)

    @property
    def customerdata(self):
        """Lower case alias for CustomerData"""
        return self.CustomerData

    @property
    def CustomLayouts(self):
        return CustomLayouts(self.com_object.CustomLayouts)

    @property
    def customlayouts(self):
        """Lower case alias for CustomLayouts"""
        return self.CustomLayouts

    @property
    def Design(self):
        return Design(self.com_object.Design)

    @property
    def design(self):
        """Lower case alias for Design"""
        return self.Design

    @property
    def guides(self):
        return Guides(self.com_object.guides)

    @property
    def HeadersFooters(self):
        return HeadersFooters(self.com_object.HeadersFooters)

    @property
    def headersfooters(self):
        """Lower case alias for HeadersFooters"""
        return self.HeadersFooters

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
    def Hyperlinks(self):
        return Hyperlinks(self.com_object.Hyperlinks)

    @property
    def hyperlinks(self):
        """Lower case alias for Hyperlinks"""
        return self.Hyperlinks

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
    def Shapes(self):
        return Shapes(self.com_object.Shapes)

    @property
    def shapes(self):
        """Lower case alias for Shapes"""
        return self.Shapes

    @property
    def SlideShowTransition(self):
        return SlideShowTransition(self.com_object.SlideShowTransition)

    @property
    def slideshowtransition(self):
        """Lower case alias for SlideShowTransition"""
        return self.SlideShowTransition

    @property
    def TextStyles(self):
        return TextStyles(self.com_object.TextStyles)

    @property
    def textstyles(self):
        """Lower case alias for TextStyles"""
        return self.TextStyles

    @property
    def Theme(self):
        return Theme(self.com_object.Theme)

    @property
    def theme(self):
        """Lower case alias for Theme"""
        return self.Theme

    @property
    def TimeLine(self):
        return TimeLine(self.com_object.TimeLine)

    @property
    def timeline(self):
        """Lower case alias for TimeLine"""
        return self.TimeLine

    @property
    def Width(self):
        return self.com_object.Width

    @property
    def width(self):
        """Lower case alias for Width"""
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

    @property
    def index(self):
        """Lower case alias for Index"""
        return self.Index

    @property
    def Name(self):
        return self.com_object.Name

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def Position(self):
        return self.com_object.Position

    @property
    def position(self):
        """Lower case alias for Position"""
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

    @property
    def count(self):
        """Lower case alias for Count"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def AudioCompressionType(self):
        return self.com_object.AudioCompressionType

    @property
    def audiocompressiontype(self):
        """Lower case alias for AudioCompressionType"""
        return self.AudioCompressionType

    @property
    def AudioSamplingRate(self):
        return self.com_object.AudioSamplingRate

    @property
    def audiosamplingrate(self):
        """Lower case alias for AudioSamplingRate"""
        return self.AudioSamplingRate

    @property
    def EndPoint(self):
        return self.com_object.EndPoint

    @EndPoint.setter
    def EndPoint(self, value):
        self.com_object.EndPoint = value

    @property
    def endpoint(self):
        """Lower case alias for EndPoint"""
        return self.EndPoint

    @endpoint.setter
    def endpoint(self, value):
        """Lower case alias for EndPoint.setter"""
        self.EndPoint = value

    @property
    def FadeInDuration(self):
        return self.com_object.FadeInDuration

    @FadeInDuration.setter
    def FadeInDuration(self, value):
        self.com_object.FadeInDuration = value

    @property
    def fadeinduration(self):
        """Lower case alias for FadeInDuration"""
        return self.FadeInDuration

    @fadeinduration.setter
    def fadeinduration(self, value):
        """Lower case alias for FadeInDuration.setter"""
        self.FadeInDuration = value

    @property
    def FadeOutDuration(self):
        return self.com_object.FadeOutDuration

    @FadeOutDuration.setter
    def FadeOutDuration(self, value):
        self.com_object.FadeOutDuration = value

    @property
    def fadeoutduration(self):
        """Lower case alias for FadeOutDuration"""
        return self.FadeOutDuration

    @fadeoutduration.setter
    def fadeoutduration(self, value):
        """Lower case alias for FadeOutDuration.setter"""
        self.FadeOutDuration = value

    @property
    def IsEmbedded(self):
        return self.com_object.IsEmbedded

    @property
    def isembedded(self):
        """Lower case alias for IsEmbedded"""
        return self.IsEmbedded

    @property
    def IsLinked(self):
        return self.com_object.IsLinked

    @property
    def islinked(self):
        """Lower case alias for IsLinked"""
        return self.IsLinked

    @property
    def Length(self):
        return self.com_object.Length

    @property
    def length(self):
        """Lower case alias for Length"""
        return self.Length

    @property
    def MediaBookmarks(self):
        return MediaBookmarks(self.com_object.MediaBookmarks)

    @property
    def mediabookmarks(self):
        """Lower case alias for MediaBookmarks"""
        return self.MediaBookmarks

    @property
    def Muted(self):
        return self.com_object.Muted

    @Muted.setter
    def Muted(self, value):
        self.com_object.Muted = value

    @property
    def muted(self):
        """Lower case alias for Muted"""
        return self.Muted

    @muted.setter
    def muted(self, value):
        """Lower case alias for Muted.setter"""
        self.Muted = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def ResamplingStatus(self):
        return self.com_object.ResamplingStatus

    @property
    def resamplingstatus(self):
        """Lower case alias for ResamplingStatus"""
        return self.ResamplingStatus

    @property
    def SampleHeight(self):
        return self.com_object.SampleHeight

    @property
    def sampleheight(self):
        """Lower case alias for SampleHeight"""
        return self.SampleHeight

    @property
    def SampleWidth(self):
        return self.com_object.SampleWidth

    @property
    def samplewidth(self):
        """Lower case alias for SampleWidth"""
        return self.SampleWidth

    @property
    def StartPoint(self):
        return self.com_object.StartPoint

    @StartPoint.setter
    def StartPoint(self, value):
        self.com_object.StartPoint = value

    @property
    def startpoint(self):
        """Lower case alias for StartPoint"""
        return self.StartPoint

    @startpoint.setter
    def startpoint(self, value):
        """Lower case alias for StartPoint.setter"""
        self.StartPoint = value

    @property
    def VideoCompressionType(self):
        return self.com_object.VideoCompressionType

    @property
    def videocompressiontype(self):
        """Lower case alias for VideoCompressionType"""
        return self.VideoCompressionType

    @property
    def VideoFrameRate(self):
        return self.com_object.VideoFrameRate

    @property
    def videoframerate(self):
        """Lower case alias for VideoFrameRate"""
        return self.VideoFrameRate

    @property
    def Volume(self):
        return self.com_object.Volume

    @Volume.setter
    def Volume(self, value):
        self.com_object.Volume = value

    @property
    def volume(self):
        """Lower case alias for Volume"""
        return self.Volume

    @volume.setter
    def volume(self, value):
        """Lower case alias for Volume.setter"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def AutoFit(self):
        return self.com_object.AutoFit

    @AutoFit.setter
    def AutoFit(self, value):
        self.com_object.AutoFit = value

    @property
    def autofit(self):
        """Lower case alias for AutoFit"""
        return self.AutoFit

    @autofit.setter
    def autofit(self, value):
        """Lower case alias for AutoFit.setter"""
        self.AutoFit = value

    @property
    def CameraPositionX(self):
        return self.com_object.CameraPositionX

    @CameraPositionX.setter
    def CameraPositionX(self, value):
        self.com_object.CameraPositionX = value

    @property
    def camerapositionx(self):
        """Lower case alias for CameraPositionX"""
        return self.CameraPositionX

    @camerapositionx.setter
    def camerapositionx(self, value):
        """Lower case alias for CameraPositionX.setter"""
        self.CameraPositionX = value

    @property
    def CameraPositionY(self):
        return self.com_object.CameraPositionY

    @CameraPositionY.setter
    def CameraPositionY(self, value):
        self.com_object.CameraPositionY = value

    @property
    def camerapositiony(self):
        """Lower case alias for CameraPositionY"""
        return self.CameraPositionY

    @camerapositiony.setter
    def camerapositiony(self, value):
        """Lower case alias for CameraPositionY.setter"""
        self.CameraPositionY = value

    @property
    def CameraPositionZ(self):
        return self.com_object.CameraPositionZ

    @CameraPositionZ.setter
    def CameraPositionZ(self, value):
        self.com_object.CameraPositionZ = value

    @property
    def camerapositionz(self):
        """Lower case alias for CameraPositionZ"""
        return self.CameraPositionZ

    @camerapositionz.setter
    def camerapositionz(self, value):
        """Lower case alias for CameraPositionZ.setter"""
        self.CameraPositionZ = value

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def FieldOfView(self):
        return self.com_object.FieldOfView

    @FieldOfView.setter
    def FieldOfView(self, value):
        self.com_object.FieldOfView = value

    @property
    def fieldofview(self):
        """Lower case alias for FieldOfView"""
        return self.FieldOfView

    @fieldofview.setter
    def fieldofview(self, value):
        """Lower case alias for FieldOfView.setter"""
        self.FieldOfView = value

    @property
    def LookAtPointX(self):
        return self.com_object.LookAtPointX

    @LookAtPointX.setter
    def LookAtPointX(self, value):
        self.com_object.LookAtPointX = value

    @property
    def lookatpointx(self):
        """Lower case alias for LookAtPointX"""
        return self.LookAtPointX

    @lookatpointx.setter
    def lookatpointx(self, value):
        """Lower case alias for LookAtPointX.setter"""
        self.LookAtPointX = value

    @property
    def LookAtPointY(self):
        return self.com_object.LookAtPointY

    @LookAtPointY.setter
    def LookAtPointY(self, value):
        self.com_object.LookAtPointY = value

    @property
    def lookatpointy(self):
        """Lower case alias for LookAtPointY"""
        return self.LookAtPointY

    @lookatpointy.setter
    def lookatpointy(self, value):
        """Lower case alias for LookAtPointY.setter"""
        self.LookAtPointY = value

    @property
    def LookAtPointZ(self):
        return self.com_object.LookAtPointZ

    @LookAtPointZ.setter
    def LookAtPointZ(self, value):
        self.com_object.LookAtPointZ = value

    @property
    def lookatpointz(self):
        """Lower case alias for LookAtPointZ"""
        return self.LookAtPointZ

    @lookatpointz.setter
    def lookatpointz(self, value):
        """Lower case alias for LookAtPointZ.setter"""
        self.LookAtPointZ = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def RotationX(self):
        return self.com_object.RotationX

    @RotationX.setter
    def RotationX(self, value):
        self.com_object.RotationX = value

    @property
    def rotationx(self):
        """Lower case alias for RotationX"""
        return self.RotationX

    @rotationx.setter
    def rotationx(self, value):
        """Lower case alias for RotationX.setter"""
        self.RotationX = value

    @property
    def RotationY(self):
        return self.com_object.RotationY

    @RotationY.setter
    def RotationY(self, value):
        self.com_object.RotationY = value

    @property
    def rotationy(self):
        """Lower case alias for RotationY"""
        return self.RotationY

    @rotationy.setter
    def rotationy(self, value):
        """Lower case alias for RotationY.setter"""
        self.RotationY = value

    @property
    def RotationZ(self):
        return self.com_object.RotationZ

    @RotationZ.setter
    def RotationZ(self, value):
        self.com_object.RotationZ = value

    @property
    def rotationz(self):
        """Lower case alias for RotationZ"""
        return self.RotationZ

    @rotationz.setter
    def rotationz(self, value):
        """Lower case alias for RotationZ.setter"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def ByX(self):
        return self.com_object.ByX

    @ByX.setter
    def ByX(self, value):
        self.com_object.ByX = value

    @property
    def byx(self):
        """Lower case alias for ByX"""
        return self.ByX

    @byx.setter
    def byx(self, value):
        """Lower case alias for ByX.setter"""
        self.ByX = value

    @property
    def ByY(self):
        return self.com_object.ByY

    @ByY.setter
    def ByY(self, value):
        self.com_object.ByY = value

    @property
    def byy(self):
        """Lower case alias for ByY"""
        return self.ByY

    @byy.setter
    def byy(self, value):
        """Lower case alias for ByY.setter"""
        self.ByY = value

    @property
    def FromX(self):
        return self.com_object.FromX

    @FromX.setter
    def FromX(self, value):
        self.com_object.FromX = value

    @property
    def fromx(self):
        """Lower case alias for FromX"""
        return self.FromX

    @fromx.setter
    def fromx(self, value):
        """Lower case alias for FromX.setter"""
        self.FromX = value

    @property
    def FromY(self):
        return MotionEffect(self.com_object.FromY)

    @FromY.setter
    def FromY(self, value):
        self.com_object.FromY = value

    @property
    def fromy(self):
        """Lower case alias for FromY"""
        return self.FromY

    @fromy.setter
    def fromy(self, value):
        """Lower case alias for FromY.setter"""
        self.FromY = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Path(self):
        return self.com_object.Path

    @Path.setter
    def Path(self, value):
        self.com_object.Path = value

    @property
    def path(self):
        """Lower case alias for Path"""
        return self.Path

    @path.setter
    def path(self, value):
        """Lower case alias for Path.setter"""
        self.Path = value

    @property
    def ToX(self):
        return self.com_object.ToX

    @ToX.setter
    def ToX(self, value):
        self.com_object.ToX = value

    @property
    def tox(self):
        """Lower case alias for ToX"""
        return self.ToX

    @tox.setter
    def tox(self, value):
        """Lower case alias for ToX.setter"""
        self.ToX = value

    @property
    def ToY(self):
        return MotionEffect(self.com_object.ToY)

    @ToY.setter
    def ToY(self, value):
        self.com_object.ToY = value

    @property
    def toy(self):
        """Lower case alias for ToY"""
        return self.ToY

    @toy.setter
    def toy(self, value):
        """Lower case alias for ToY.setter"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Count(self):
        return self.com_object.Count

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

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
    def SlideIDs(self):
        return self.com_object.SlideIDs

    @property
    def slideids(self):
        """Lower case alias for SlideIDs"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def FollowColors(self):
        return self.com_object.FollowColors

    @FollowColors.setter
    def FollowColors(self, value):
        self.com_object.FollowColors = value

    @property
    def followcolors(self):
        """Lower case alias for FollowColors"""
        return self.FollowColors

    @followcolors.setter
    def followcolors(self, value):
        """Lower case alias for FollowColors.setter"""
        self.FollowColors = value

    @property
    def Object(self):
        return self.com_object.Object

    @property
    def object(self):
        """Lower case alias for Object"""
        return self.Object

    @property
    def ObjectVerbs(self):
        return ObjectVerbs(self.com_object.ObjectVerbs)

    @property
    def objectverbs(self):
        """Lower case alias for ObjectVerbs"""
        return self.ObjectVerbs

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def ProgID(self):
        return self.com_object.ProgID

    @property
    def progid(self):
        """Lower case alias for ProgID"""
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

    @property
    def displaypasteoptions(self):
        """Lower case alias for DisplayPasteOptions"""
        return self.DisplayPasteOptions

    @displaypasteoptions.setter
    def displaypasteoptions(self, value):
        """Lower case alias for DisplayPasteOptions.setter"""
        self.DisplayPasteOptions = value

    @property
    def ShowCoauthoringMergeChanges(self):
        return self.com_object.ShowCoauthoringMergeChanges

    @property
    def showcoauthoringmergechanges(self):
        """Lower case alias for ShowCoauthoringMergeChanges"""
        return self.ShowCoauthoringMergeChanges


class PageSetup:

    def __init__(self, pagesetup=None):
        self.com_object= pagesetup

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def FirstSlideNumber(self):
        return self.com_object.FirstSlideNumber

    @FirstSlideNumber.setter
    def FirstSlideNumber(self, value):
        self.com_object.FirstSlideNumber = value

    @property
    def firstslidenumber(self):
        """Lower case alias for FirstSlideNumber"""
        return self.FirstSlideNumber

    @firstslidenumber.setter
    def firstslidenumber(self, value):
        """Lower case alias for FirstSlideNumber.setter"""
        self.FirstSlideNumber = value

    @property
    def NotesOrientation(self):
        return self.com_object.NotesOrientation

    @NotesOrientation.setter
    def NotesOrientation(self, value):
        self.com_object.NotesOrientation = value

    @property
    def notesorientation(self):
        """Lower case alias for NotesOrientation"""
        return self.NotesOrientation

    @notesorientation.setter
    def notesorientation(self, value):
        """Lower case alias for NotesOrientation.setter"""
        self.NotesOrientation = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def SlideHeight(self):
        return self.com_object.SlideHeight

    @SlideHeight.setter
    def SlideHeight(self, value):
        self.com_object.SlideHeight = value

    @property
    def slideheight(self):
        """Lower case alias for SlideHeight"""
        return self.SlideHeight

    @slideheight.setter
    def slideheight(self, value):
        """Lower case alias for SlideHeight.setter"""
        self.SlideHeight = value

    @property
    def SlideOrientation(self):
        return self.com_object.SlideOrientation

    @SlideOrientation.setter
    def SlideOrientation(self, value):
        self.com_object.SlideOrientation = value

    @property
    def slideorientation(self):
        """Lower case alias for SlideOrientation"""
        return self.SlideOrientation

    @slideorientation.setter
    def slideorientation(self, value):
        """Lower case alias for SlideOrientation.setter"""
        self.SlideOrientation = value

    @property
    def SlideSize(self):
        return self.com_object.SlideSize

    @SlideSize.setter
    def SlideSize(self, value):
        self.com_object.SlideSize = value

    @property
    def slidesize(self):
        """Lower case alias for SlideSize"""
        return self.SlideSize

    @slidesize.setter
    def slidesize(self, value):
        """Lower case alias for SlideSize.setter"""
        self.SlideSize = value

    @property
    def SlideWidth(self):
        return self.com_object.SlideWidth

    @SlideWidth.setter
    def SlideWidth(self, value):
        self.com_object.SlideWidth = value

    @property
    def slidewidth(self):
        """Lower case alias for SlideWidth"""
        return self.SlideWidth

    @slidewidth.setter
    def slidewidth(self, value):
        """Lower case alias for SlideWidth.setter"""
        self.SlideWidth = value


class Pane:

    def __init__(self, pane=None):
        self.com_object= pane

    @property
    def Active(self):
        return self.com_object.Active

    @property
    def active(self):
        """Lower case alias for Active"""
        return self.Active

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def ViewType(self):
        return self.com_object.ViewType

    @property
    def viewtype(self):
        """Lower case alias for ViewType"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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

    @property
    def alignment(self):
        """Lower case alias for Alignment"""
        return self.Alignment

    @alignment.setter
    def alignment(self, value):
        """Lower case alias for Alignment.setter"""
        self.Alignment = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def BaseLineAlignment(self):
        return self.com_object.BaseLineAlignment

    @BaseLineAlignment.setter
    def BaseLineAlignment(self, value):
        self.com_object.BaseLineAlignment = value

    @property
    def baselinealignment(self):
        """Lower case alias for BaseLineAlignment"""
        return self.BaseLineAlignment

    @baselinealignment.setter
    def baselinealignment(self, value):
        """Lower case alias for BaseLineAlignment.setter"""
        self.BaseLineAlignment = value

    @property
    def Bullet(self):
        return BulletFormat(self.com_object.Bullet)

    @property
    def bullet(self):
        """Lower case alias for Bullet"""
        return self.Bullet

    @property
    def FarEastLineBreakControl(self):
        return self.com_object.FarEastLineBreakControl

    @FarEastLineBreakControl.setter
    def FarEastLineBreakControl(self, value):
        self.com_object.FarEastLineBreakControl = value

    @property
    def fareastlinebreakcontrol(self):
        """Lower case alias for FarEastLineBreakControl"""
        return self.FarEastLineBreakControl

    @fareastlinebreakcontrol.setter
    def fareastlinebreakcontrol(self, value):
        """Lower case alias for FarEastLineBreakControl.setter"""
        self.FarEastLineBreakControl = value

    @property
    def HangingPunctuation(self):
        return self.com_object.HangingPunctuation

    @HangingPunctuation.setter
    def HangingPunctuation(self, value):
        self.com_object.HangingPunctuation = value

    @property
    def hangingpunctuation(self):
        """Lower case alias for HangingPunctuation"""
        return self.HangingPunctuation

    @hangingpunctuation.setter
    def hangingpunctuation(self, value):
        """Lower case alias for HangingPunctuation.setter"""
        self.HangingPunctuation = value

    @property
    def LineRuleAfter(self):
        return self.com_object.LineRuleAfter

    @LineRuleAfter.setter
    def LineRuleAfter(self, value):
        self.com_object.LineRuleAfter = value

    @property
    def lineruleafter(self):
        """Lower case alias for LineRuleAfter"""
        return self.LineRuleAfter

    @lineruleafter.setter
    def lineruleafter(self, value):
        """Lower case alias for LineRuleAfter.setter"""
        self.LineRuleAfter = value

    @property
    def LineRuleBefore(self):
        return self.com_object.LineRuleBefore

    @LineRuleBefore.setter
    def LineRuleBefore(self, value):
        self.com_object.LineRuleBefore = value

    @property
    def linerulebefore(self):
        """Lower case alias for LineRuleBefore"""
        return self.LineRuleBefore

    @linerulebefore.setter
    def linerulebefore(self, value):
        """Lower case alias for LineRuleBefore.setter"""
        self.LineRuleBefore = value

    @property
    def LineRuleWithin(self):
        return self.com_object.LineRuleWithin

    @LineRuleWithin.setter
    def LineRuleWithin(self, value):
        self.com_object.LineRuleWithin = value

    @property
    def linerulewithin(self):
        """Lower case alias for LineRuleWithin"""
        return self.LineRuleWithin

    @linerulewithin.setter
    def linerulewithin(self, value):
        """Lower case alias for LineRuleWithin.setter"""
        self.LineRuleWithin = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def SpaceAfter(self):
        return self.com_object.SpaceAfter

    @SpaceAfter.setter
    def SpaceAfter(self, value):
        self.com_object.SpaceAfter = value

    @property
    def spaceafter(self):
        """Lower case alias for SpaceAfter"""
        return self.SpaceAfter

    @spaceafter.setter
    def spaceafter(self, value):
        """Lower case alias for SpaceAfter.setter"""
        self.SpaceAfter = value

    @property
    def SpaceBefore(self):
        return self.com_object.SpaceBefore

    @SpaceBefore.setter
    def SpaceBefore(self, value):
        self.com_object.SpaceBefore = value

    @property
    def spacebefore(self):
        """Lower case alias for SpaceBefore"""
        return self.SpaceBefore

    @spacebefore.setter
    def spacebefore(self, value):
        """Lower case alias for SpaceBefore.setter"""
        self.SpaceBefore = value

    @property
    def SpaceWithin(self):
        return self.com_object.SpaceWithin

    @SpaceWithin.setter
    def SpaceWithin(self, value):
        self.com_object.SpaceWithin = value

    @property
    def spacewithin(self):
        """Lower case alias for SpaceWithin"""
        return self.SpaceWithin

    @spacewithin.setter
    def spacewithin(self, value):
        """Lower case alias for SpaceWithin.setter"""
        self.SpaceWithin = value

    @property
    def TextDirection(self):
        return self.com_object.TextDirection

    @TextDirection.setter
    def TextDirection(self, value):
        self.com_object.TextDirection = value

    @property
    def textdirection(self):
        """Lower case alias for TextDirection"""
        return self.TextDirection

    @textdirection.setter
    def textdirection(self, value):
        """Lower case alias for TextDirection.setter"""
        self.TextDirection = value

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


class PictureFormat:

    def __init__(self, pictureformat=None):
        self.com_object= pictureformat

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Brightness(self):
        return self.com_object.Brightness

    @Brightness.setter
    def Brightness(self, value):
        self.com_object.Brightness = value

    @property
    def brightness(self):
        """Lower case alias for Brightness"""
        return self.Brightness

    @brightness.setter
    def brightness(self, value):
        """Lower case alias for Brightness.setter"""
        self.Brightness = value

    @property
    def ColorType(self):
        return self.com_object.ColorType

    @ColorType.setter
    def ColorType(self, value):
        self.com_object.ColorType = value

    @property
    def colortype(self):
        """Lower case alias for ColorType"""
        return self.ColorType

    @colortype.setter
    def colortype(self, value):
        """Lower case alias for ColorType.setter"""
        self.ColorType = value

    @property
    def Contrast(self):
        return self.com_object.Contrast

    @Contrast.setter
    def Contrast(self, value):
        self.com_object.Contrast = value

    @property
    def contrast(self):
        """Lower case alias for Contrast"""
        return self.Contrast

    @contrast.setter
    def contrast(self, value):
        """Lower case alias for Contrast.setter"""
        self.Contrast = value

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Crop(self):
        return self.com_object.Crop

    @Crop.setter
    def Crop(self, value):
        self.com_object.Crop = value

    @property
    def crop(self):
        """Lower case alias for Crop"""
        return self.Crop

    @crop.setter
    def crop(self, value):
        """Lower case alias for Crop.setter"""
        self.Crop = value

    @property
    def CropBottom(self):
        return self.com_object.CropBottom

    @CropBottom.setter
    def CropBottom(self, value):
        self.com_object.CropBottom = value

    @property
    def cropbottom(self):
        """Lower case alias for CropBottom"""
        return self.CropBottom

    @cropbottom.setter
    def cropbottom(self, value):
        """Lower case alias for CropBottom.setter"""
        self.CropBottom = value

    @property
    def CropLeft(self):
        return self.com_object.CropLeft

    @CropLeft.setter
    def CropLeft(self, value):
        self.com_object.CropLeft = value

    @property
    def cropleft(self):
        """Lower case alias for CropLeft"""
        return self.CropLeft

    @cropleft.setter
    def cropleft(self, value):
        """Lower case alias for CropLeft.setter"""
        self.CropLeft = value

    @property
    def CropRight(self):
        return self.com_object.CropRight

    @CropRight.setter
    def CropRight(self, value):
        self.com_object.CropRight = value

    @property
    def cropright(self):
        """Lower case alias for CropRight"""
        return self.CropRight

    @cropright.setter
    def cropright(self, value):
        """Lower case alias for CropRight.setter"""
        self.CropRight = value

    @property
    def CropTop(self):
        return self.com_object.CropTop

    @CropTop.setter
    def CropTop(self, value):
        self.com_object.CropTop = value

    @property
    def croptop(self):
        """Lower case alias for CropTop"""
        return self.CropTop

    @croptop.setter
    def croptop(self, value):
        """Lower case alias for CropTop.setter"""
        self.CropTop = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def TransparencyColor(self):
        return self.com_object.TransparencyColor

    @TransparencyColor.setter
    def TransparencyColor(self, value):
        self.com_object.TransparencyColor = value

    @property
    def transparencycolor(self):
        """Lower case alias for TransparencyColor"""
        return self.TransparencyColor

    @transparencycolor.setter
    def transparencycolor(self, value):
        """Lower case alias for TransparencyColor.setter"""
        self.TransparencyColor = value

    @property
    def TransparentBackground(self):
        return self.com_object.TransparentBackground

    @TransparentBackground.setter
    def TransparentBackground(self, value):
        self.com_object.TransparentBackground = value

    @property
    def transparentbackground(self):
        """Lower case alias for TransparentBackground"""
        return self.TransparentBackground

    @transparentbackground.setter
    def transparentbackground(self, value):
        """Lower case alias for TransparentBackground.setter"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def ContainedType(self):
        return self.com_object.ContainedType

    @property
    def containedtype(self):
        """Lower case alias for ContainedType"""
        return self.ContainedType

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
    def Type(self):
        return self.com_object.Type

    @property
    def type(self):
        """Lower case alias for Type"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def CurrentPosition(self):
        return self.com_object.CurrentPosition

    @CurrentPosition.setter
    def CurrentPosition(self, value):
        self.com_object.CurrentPosition = value

    @property
    def currentposition(self):
        """Lower case alias for CurrentPosition"""
        return self.CurrentPosition

    @currentposition.setter
    def currentposition(self, value):
        """Lower case alias for CurrentPosition.setter"""
        self.CurrentPosition = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def State(self):
        return self.com_object.State

    @property
    def state(self):
        """Lower case alias for State"""
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

    @property
    def actionverb(self):
        """Lower case alias for ActionVerb"""
        return self.ActionVerb

    @actionverb.setter
    def actionverb(self, value):
        """Lower case alias for ActionVerb.setter"""
        self.ActionVerb = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def HideWhileNotPlaying(self):
        return self.com_object.HideWhileNotPlaying

    @HideWhileNotPlaying.setter
    def HideWhileNotPlaying(self, value):
        self.com_object.HideWhileNotPlaying = value

    @property
    def hidewhilenotplaying(self):
        """Lower case alias for HideWhileNotPlaying"""
        return self.HideWhileNotPlaying

    @hidewhilenotplaying.setter
    def hidewhilenotplaying(self, value):
        """Lower case alias for HideWhileNotPlaying.setter"""
        self.HideWhileNotPlaying = value

    @property
    def LoopUntilStopped(self):
        return self.com_object.LoopUntilStopped

    @LoopUntilStopped.setter
    def LoopUntilStopped(self, value):
        self.com_object.LoopUntilStopped = value

    @property
    def loopuntilstopped(self):
        """Lower case alias for LoopUntilStopped"""
        return self.LoopUntilStopped

    @loopuntilstopped.setter
    def loopuntilstopped(self, value):
        """Lower case alias for LoopUntilStopped.setter"""
        self.LoopUntilStopped = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PauseAnimation(self):
        return self.com_object.PauseAnimation

    @PauseAnimation.setter
    def PauseAnimation(self, value):
        self.com_object.PauseAnimation = value

    @property
    def pauseanimation(self):
        """Lower case alias for PauseAnimation"""
        return self.PauseAnimation

    @pauseanimation.setter
    def pauseanimation(self, value):
        """Lower case alias for PauseAnimation.setter"""
        self.PauseAnimation = value

    @property
    def PlayOnEntry(self):
        return self.com_object.PlayOnEntry

    @PlayOnEntry.setter
    def PlayOnEntry(self, value):
        self.com_object.PlayOnEntry = value

    @property
    def playonentry(self):
        """Lower case alias for PlayOnEntry"""
        return self.PlayOnEntry

    @playonentry.setter
    def playonentry(self, value):
        """Lower case alias for PlayOnEntry.setter"""
        self.PlayOnEntry = value

    @property
    def RewindMovie(self):
        return self.com_object.RewindMovie

    @RewindMovie.setter
    def RewindMovie(self, value):
        self.com_object.RewindMovie = value

    @property
    def rewindmovie(self):
        """Lower case alias for RewindMovie"""
        return self.RewindMovie

    @rewindmovie.setter
    def rewindmovie(self, value):
        """Lower case alias for RewindMovie.setter"""
        self.RewindMovie = value

    @property
    def StopAfterSlides(self):
        return self.com_object.StopAfterSlides

    @StopAfterSlides.setter
    def StopAfterSlides(self, value):
        self.com_object.StopAfterSlides = value

    @property
    def stopafterslides(self):
        """Lower case alias for StopAfterSlides"""
        return self.StopAfterSlides

    @stopafterslides.setter
    def stopafterslides(self, value):
        """Lower case alias for StopAfterSlides.setter"""
        self.StopAfterSlides = value


class PlotArea:

    def __init__(self, plotarea=None):
        self.com_object= plotarea

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

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
    def InsideHeight(self):
        return self.com_object.InsideHeight

    @InsideHeight.setter
    def InsideHeight(self, value):
        self.com_object.InsideHeight = value

    @property
    def insideheight(self):
        """Lower case alias for InsideHeight"""
        return self.InsideHeight

    @insideheight.setter
    def insideheight(self, value):
        """Lower case alias for InsideHeight.setter"""
        self.InsideHeight = value

    @property
    def InsideLeft(self):
        return self.com_object.InsideLeft

    @InsideLeft.setter
    def InsideLeft(self, value):
        self.com_object.InsideLeft = value

    @property
    def insideleft(self):
        """Lower case alias for InsideLeft"""
        return self.InsideLeft

    @insideleft.setter
    def insideleft(self, value):
        """Lower case alias for InsideLeft.setter"""
        self.InsideLeft = value

    @property
    def InsideTop(self):
        return self.com_object.InsideTop

    @InsideTop.setter
    def InsideTop(self, value):
        self.com_object.InsideTop = value

    @property
    def insidetop(self):
        """Lower case alias for InsideTop"""
        return self.InsideTop

    @insidetop.setter
    def insidetop(self, value):
        """Lower case alias for InsideTop.setter"""
        self.InsideTop = value

    @property
    def InsideWidth(self):
        return self.com_object.InsideWidth

    @InsideWidth.setter
    def InsideWidth(self, value):
        self.com_object.InsideWidth = value

    @property
    def insidewidth(self):
        """Lower case alias for InsideWidth"""
        return self.InsideWidth

    @insidewidth.setter
    def insidewidth(self, value):
        """Lower case alias for InsideWidth.setter"""
        self.InsideWidth = value

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
    def Position(self):
        return XlChartElementPosition(self.com_object.Position)

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def ApplyPictToEnd(self):
        return self.com_object.ApplyPictToEnd

    @ApplyPictToEnd.setter
    def ApplyPictToEnd(self, value):
        self.com_object.ApplyPictToEnd = value

    @property
    def applypicttoend(self):
        """Lower case alias for ApplyPictToEnd"""
        return self.ApplyPictToEnd

    @applypicttoend.setter
    def applypicttoend(self, value):
        """Lower case alias for ApplyPictToEnd.setter"""
        self.ApplyPictToEnd = value

    @property
    def ApplyPictToFront(self):
        return self.com_object.ApplyPictToFront

    @ApplyPictToFront.setter
    def ApplyPictToFront(self, value):
        self.com_object.ApplyPictToFront = value

    @property
    def applypicttofront(self):
        """Lower case alias for ApplyPictToFront"""
        return self.ApplyPictToFront

    @applypicttofront.setter
    def applypicttofront(self, value):
        """Lower case alias for ApplyPictToFront.setter"""
        self.ApplyPictToFront = value

    @property
    def ApplyPictToSides(self):
        return self.com_object.ApplyPictToSides

    @ApplyPictToSides.setter
    def ApplyPictToSides(self, value):
        self.com_object.ApplyPictToSides = value

    @property
    def applypicttosides(self):
        """Lower case alias for ApplyPictToSides"""
        return self.ApplyPictToSides

    @applypicttosides.setter
    def applypicttosides(self, value):
        """Lower case alias for ApplyPictToSides.setter"""
        self.ApplyPictToSides = value

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def DataLabel(self):
        return DataLabel(self.com_object.DataLabel)

    @property
    def datalabel(self):
        """Lower case alias for DataLabel"""
        return self.DataLabel

    @property
    def Explosion(self):
        return self.com_object.Explosion

    @Explosion.setter
    def Explosion(self, value):
        self.com_object.Explosion = value

    @property
    def explosion(self):
        """Lower case alias for Explosion"""
        return self.Explosion

    @explosion.setter
    def explosion(self, value):
        """Lower case alias for Explosion.setter"""
        self.Explosion = value

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

    @property
    def Has3DEffect(self):
        return self.com_object.Has3DEffect

    @Has3DEffect.setter
    def Has3DEffect(self, value):
        self.com_object.Has3DEffect = value

    @property
    def has3deffect(self):
        """Lower case alias for Has3DEffect"""
        return self.Has3DEffect

    @has3deffect.setter
    def has3deffect(self, value):
        """Lower case alias for Has3DEffect.setter"""
        self.Has3DEffect = value

    @property
    def HasDataLabel(self):
        return self.com_object.HasDataLabel

    @HasDataLabel.setter
    def HasDataLabel(self, value):
        self.com_object.HasDataLabel = value

    @property
    def hasdatalabel(self):
        """Lower case alias for HasDataLabel"""
        return self.HasDataLabel

    @hasdatalabel.setter
    def hasdatalabel(self, value):
        """Lower case alias for HasDataLabel.setter"""
        self.HasDataLabel = value

    @property
    def Height(self):
        return self.com_object.Height

    @property
    def height(self):
        """Lower case alias for Height"""
        return self.Height

    @property
    def InvertIfNegative(self):
        return self.com_object.InvertIfNegative

    @InvertIfNegative.setter
    def InvertIfNegative(self, value):
        self.com_object.InvertIfNegative = value

    @property
    def invertifnegative(self):
        """Lower case alias for InvertIfNegative"""
        return self.InvertIfNegative

    @invertifnegative.setter
    def invertifnegative(self, value):
        """Lower case alias for InvertIfNegative.setter"""
        self.InvertIfNegative = value

    @property
    def istotal(self):
        return self.com_object.istotal

    @istotal.setter
    def istotal(self, value):
        self.com_object.istotal = value

    @property
    def Left(self):
        return self.com_object.Left

    @property
    def left(self):
        """Lower case alias for Left"""
        return self.Left

    @property
    def MarkerBackgroundColor(self):
        return self.com_object.MarkerBackgroundColor

    @MarkerBackgroundColor.setter
    def MarkerBackgroundColor(self, value):
        self.com_object.MarkerBackgroundColor = value

    @property
    def markerbackgroundcolor(self):
        """Lower case alias for MarkerBackgroundColor"""
        return self.MarkerBackgroundColor

    @markerbackgroundcolor.setter
    def markerbackgroundcolor(self, value):
        """Lower case alias for MarkerBackgroundColor.setter"""
        self.MarkerBackgroundColor = value

    @property
    def MarkerBackgroundColorIndex(self):
        return XlColorIndex(self.com_object.MarkerBackgroundColorIndex)

    @MarkerBackgroundColorIndex.setter
    def MarkerBackgroundColorIndex(self, value):
        self.com_object.MarkerBackgroundColorIndex = value

    @property
    def markerbackgroundcolorindex(self):
        """Lower case alias for MarkerBackgroundColorIndex"""
        return self.MarkerBackgroundColorIndex

    @markerbackgroundcolorindex.setter
    def markerbackgroundcolorindex(self, value):
        """Lower case alias for MarkerBackgroundColorIndex.setter"""
        self.MarkerBackgroundColorIndex = value

    @property
    def MarkerForegroundColor(self):
        return self.com_object.MarkerForegroundColor

    @MarkerForegroundColor.setter
    def MarkerForegroundColor(self, value):
        self.com_object.MarkerForegroundColor = value

    @property
    def markerforegroundcolor(self):
        """Lower case alias for MarkerForegroundColor"""
        return self.MarkerForegroundColor

    @markerforegroundcolor.setter
    def markerforegroundcolor(self, value):
        """Lower case alias for MarkerForegroundColor.setter"""
        self.MarkerForegroundColor = value

    @property
    def MarkerForegroundColorIndex(self):
        return XlColorIndex(self.com_object.MarkerForegroundColorIndex)

    @MarkerForegroundColorIndex.setter
    def MarkerForegroundColorIndex(self, value):
        self.com_object.MarkerForegroundColorIndex = value

    @property
    def markerforegroundcolorindex(self):
        """Lower case alias for MarkerForegroundColorIndex"""
        return self.MarkerForegroundColorIndex

    @markerforegroundcolorindex.setter
    def markerforegroundcolorindex(self, value):
        """Lower case alias for MarkerForegroundColorIndex.setter"""
        self.MarkerForegroundColorIndex = value

    @property
    def MarkerSize(self):
        return self.com_object.MarkerSize

    @MarkerSize.setter
    def MarkerSize(self, value):
        self.com_object.MarkerSize = value

    @property
    def markersize(self):
        """Lower case alias for MarkerSize"""
        return self.MarkerSize

    @markersize.setter
    def markersize(self, value):
        """Lower case alias for MarkerSize.setter"""
        self.MarkerSize = value

    @property
    def MarkerStyle(self):
        return XlMarkerStyle(self.com_object.MarkerStyle)

    @MarkerStyle.setter
    def MarkerStyle(self, value):
        self.com_object.MarkerStyle = value

    @property
    def markerstyle(self):
        """Lower case alias for MarkerStyle"""
        return self.MarkerStyle

    @markerstyle.setter
    def markerstyle(self, value):
        """Lower case alias for MarkerStyle.setter"""
        self.MarkerStyle = value

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
    def PictureType(self):
        return XlChartPictureType(self.com_object.PictureType)

    @PictureType.setter
    def PictureType(self, value):
        self.com_object.PictureType = value

    @property
    def picturetype(self):
        """Lower case alias for PictureType"""
        return self.PictureType

    @picturetype.setter
    def picturetype(self, value):
        """Lower case alias for PictureType.setter"""
        self.PictureType = value

    @property
    def PictureUnit2(self):
        return self.com_object.PictureUnit2

    @PictureUnit2.setter
    def PictureUnit2(self, value):
        self.com_object.PictureUnit2 = value

    @property
    def pictureunit2(self):
        """Lower case alias for PictureUnit2"""
        return self.PictureUnit2

    @pictureunit2.setter
    def pictureunit2(self, value):
        """Lower case alias for PictureUnit2.setter"""
        self.PictureUnit2 = value

    @property
    def SecondaryPlot(self):
        return self.com_object.SecondaryPlot

    @SecondaryPlot.setter
    def SecondaryPlot(self, value):
        self.com_object.SecondaryPlot = value

    @property
    def secondaryplot(self):
        """Lower case alias for SecondaryPlot"""
        return self.SecondaryPlot

    @secondaryplot.setter
    def secondaryplot(self, value):
        """Lower case alias for SecondaryPlot.setter"""
        self.SecondaryPlot = value

    @property
    def Shadow(self):
        return self.com_object.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.com_object.Shadow = value

    @property
    def shadow(self):
        """Lower case alias for Shadow"""
        return self.Shadow

    @shadow.setter
    def shadow(self, value):
        """Lower case alias for Shadow.setter"""
        self.Shadow = value

    @property
    def Top(self):
        return self.com_object.Top

    @property
    def top(self):
        """Lower case alias for Top"""
        return self.Top

    @property
    def Width(self):
        return self.com_object.Width

    @property
    def width(self):
        """Lower case alias for Width"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Count(self):
        return self.com_object.Count

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def AutoSaveOn(self):
        return self.com_object.AutoSaveOn

    @AutoSaveOn.setter
    def AutoSaveOn(self, value):
        self.com_object.AutoSaveOn = value

    @property
    def autosaveon(self):
        """Lower case alias for AutoSaveOn"""
        return self.AutoSaveOn

    @autosaveon.setter
    def autosaveon(self, value):
        """Lower case alias for AutoSaveOn.setter"""
        self.AutoSaveOn = value

    @property
    def Broadcast(self):
        return Broadcast(self.com_object.Broadcast)

    @property
    def broadcast(self):
        """Lower case alias for Broadcast"""
        return self.Broadcast

    @property
    def BuiltInDocumentProperties(self):
        return self.com_object.BuiltInDocumentProperties

    @property
    def builtindocumentproperties(self):
        """Lower case alias for BuiltInDocumentProperties"""
        return self.BuiltInDocumentProperties

    @property
    def chartdatapointtrack(self):
        return self.com_object.chartdatapointtrack

    @chartdatapointtrack.setter
    def chartdatapointtrack(self, value):
        self.com_object.chartdatapointtrack = value

    @property
    def Coauthoring(self):
        return Coauthoring(self.com_object.Coauthoring)

    @property
    def coauthoring(self):
        """Lower case alias for Coauthoring"""
        return self.Coauthoring

    @property
    def ColorSchemes(self):
        return ColorSchemes(self.com_object.ColorSchemes)

    @property
    def colorschemes(self):
        """Lower case alias for ColorSchemes"""
        return self.ColorSchemes

    @property
    def CommandBars(self):
        return self.com_object.CommandBars

    @property
    def commandbars(self):
        """Lower case alias for CommandBars"""
        return self.CommandBars

    @property
    def Container(self):
        return self.com_object.Container

    @property
    def container(self):
        """Lower case alias for Container"""
        return self.Container

    @property
    def ContentTypeProperties(self):
        return self.com_object.ContentTypeProperties

    @property
    def contenttypeproperties(self):
        """Lower case alias for ContentTypeProperties"""
        return self.ContentTypeProperties

    @property
    def CreateVideoStatus(self):
        return Presentation(self.com_object.CreateVideoStatus)

    @property
    def createvideostatus(self):
        """Lower case alias for CreateVideoStatus"""
        return self.CreateVideoStatus

    @property
    def CustomDocumentProperties(self):
        return self.com_object.CustomDocumentProperties

    @property
    def customdocumentproperties(self):
        """Lower case alias for CustomDocumentProperties"""
        return self.CustomDocumentProperties

    @property
    def CustomerData(self):
        return CustomerData(self.com_object.CustomerData)

    @property
    def customerdata(self):
        """Lower case alias for CustomerData"""
        return self.CustomerData

    @property
    def CustomXMLParts(self):
        return self.com_object.CustomXMLParts

    @property
    def customxmlparts(self):
        """Lower case alias for CustomXMLParts"""
        return self.CustomXMLParts

    @property
    def DefaultLanguageID(self):
        return self.com_object.DefaultLanguageID

    @DefaultLanguageID.setter
    def DefaultLanguageID(self, value):
        self.com_object.DefaultLanguageID = value

    @property
    def defaultlanguageid(self):
        """Lower case alias for DefaultLanguageID"""
        return self.DefaultLanguageID

    @defaultlanguageid.setter
    def defaultlanguageid(self, value):
        """Lower case alias for DefaultLanguageID.setter"""
        self.DefaultLanguageID = value

    @property
    def DefaultShape(self):
        return Shape(self.com_object.DefaultShape)

    @property
    def defaultshape(self):
        """Lower case alias for DefaultShape"""
        return self.DefaultShape

    @property
    def Designs(self):
        return Designs(self.com_object.Designs)

    @property
    def designs(self):
        """Lower case alias for Designs"""
        return self.Designs

    @property
    def DisplayComments(self):
        return self.com_object.DisplayComments

    @DisplayComments.setter
    def DisplayComments(self, value):
        self.com_object.DisplayComments = value

    @property
    def displaycomments(self):
        """Lower case alias for DisplayComments"""
        return self.DisplayComments

    @displaycomments.setter
    def displaycomments(self, value):
        """Lower case alias for DisplayComments.setter"""
        self.DisplayComments = value

    @property
    def DocumentInspectors(self):
        return self.com_object.DocumentInspectors

    @property
    def documentinspectors(self):
        """Lower case alias for DocumentInspectors"""
        return self.DocumentInspectors

    @property
    def DocumentLibraryVersions(self):
        return self.com_object.DocumentLibraryVersions

    @property
    def documentlibraryversions(self):
        """Lower case alias for DocumentLibraryVersions"""
        return self.DocumentLibraryVersions

    @property
    def EncryptionProvider(self):
        return self.com_object.EncryptionProvider

    @EncryptionProvider.setter
    def EncryptionProvider(self, value):
        self.com_object.EncryptionProvider = value

    @property
    def encryptionprovider(self):
        """Lower case alias for EncryptionProvider"""
        return self.EncryptionProvider

    @encryptionprovider.setter
    def encryptionprovider(self, value):
        """Lower case alias for EncryptionProvider.setter"""
        self.EncryptionProvider = value

    @property
    def EnvelopeVisible(self):
        return self.com_object.EnvelopeVisible

    @EnvelopeVisible.setter
    def EnvelopeVisible(self, value):
        self.com_object.EnvelopeVisible = value

    @property
    def envelopevisible(self):
        """Lower case alias for EnvelopeVisible"""
        return self.EnvelopeVisible

    @envelopevisible.setter
    def envelopevisible(self, value):
        """Lower case alias for EnvelopeVisible.setter"""
        self.EnvelopeVisible = value

    @property
    def ExtraColors(self):
        return ExtraColors(self.com_object.ExtraColors)

    @property
    def extracolors(self):
        """Lower case alias for ExtraColors"""
        return self.ExtraColors

    @property
    def FarEastLineBreakLanguage(self):
        return self.com_object.FarEastLineBreakLanguage

    @FarEastLineBreakLanguage.setter
    def FarEastLineBreakLanguage(self, value):
        self.com_object.FarEastLineBreakLanguage = value

    @property
    def fareastlinebreaklanguage(self):
        """Lower case alias for FarEastLineBreakLanguage"""
        return self.FarEastLineBreakLanguage

    @fareastlinebreaklanguage.setter
    def fareastlinebreaklanguage(self, value):
        """Lower case alias for FarEastLineBreakLanguage.setter"""
        self.FarEastLineBreakLanguage = value

    @property
    def FarEastLineBreakLevel(self):
        return self.com_object.FarEastLineBreakLevel

    @FarEastLineBreakLevel.setter
    def FarEastLineBreakLevel(self, value):
        self.com_object.FarEastLineBreakLevel = value

    @property
    def fareastlinebreaklevel(self):
        """Lower case alias for FarEastLineBreakLevel"""
        return self.FarEastLineBreakLevel

    @fareastlinebreaklevel.setter
    def fareastlinebreaklevel(self, value):
        """Lower case alias for FarEastLineBreakLevel.setter"""
        self.FarEastLineBreakLevel = value

    @property
    def Final(self):
        return self.com_object.Final

    @Final.setter
    def Final(self, value):
        self.com_object.Final = value

    @property
    def final(self):
        """Lower case alias for Final"""
        return self.Final

    @final.setter
    def final(self, value):
        """Lower case alias for Final.setter"""
        self.Final = value

    @property
    def Fonts(self):
        return Fonts(self.com_object.Fonts)

    @property
    def fonts(self):
        """Lower case alias for Fonts"""
        return self.Fonts

    @property
    def FullName(self):
        return self.com_object.FullName

    @property
    def fullname(self):
        """Lower case alias for FullName"""
        return self.FullName

    @property
    def GridDistance(self):
        return self.com_object.GridDistance

    @GridDistance.setter
    def GridDistance(self, value):
        self.com_object.GridDistance = value

    @property
    def griddistance(self):
        """Lower case alias for GridDistance"""
        return self.GridDistance

    @griddistance.setter
    def griddistance(self, value):
        """Lower case alias for GridDistance.setter"""
        self.GridDistance = value

    @property
    def guides(self):
        return self.com_object.guides

    @property
    def HandoutMaster(self):
        return Master(self.com_object.HandoutMaster)

    @property
    def handoutmaster(self):
        """Lower case alias for HandoutMaster"""
        return self.HandoutMaster

    @property
    def HasHandoutMaster(self):
        return self.com_object.HasHandoutMaster

    @property
    def hashandoutmaster(self):
        """Lower case alias for HasHandoutMaster"""
        return self.HasHandoutMaster

    @property
    def HasNotesMaster(self):
        return self.com_object.HasNotesMaster

    @property
    def hasnotesmaster(self):
        """Lower case alias for HasNotesMaster"""
        return self.HasNotesMaster

    @property
    def HasTitleMaster(self):
        return self.com_object.HasTitleMaster

    @property
    def hastitlemaster(self):
        """Lower case alias for HasTitleMaster"""
        return self.HasTitleMaster

    @property
    def HasVBProject(self):
        return self.com_object.HasVBProject

    @property
    def hasvbproject(self):
        """Lower case alias for HasVBProject"""
        return self.HasVBProject

    @property
    def InMergeMode(self):
        return self.com_object.InMergeMode

    @property
    def inmergemode(self):
        """Lower case alias for InMergeMode"""
        return self.InMergeMode

    @property
    def IsFullyDownloaded(self):
        return self.com_object.IsFullyDownloaded

    @property
    def isfullydownloaded(self):
        """Lower case alias for IsFullyDownloaded"""
        return self.IsFullyDownloaded

    @property
    def LayoutDirection(self):
        return self.com_object.LayoutDirection

    @LayoutDirection.setter
    def LayoutDirection(self, value):
        self.com_object.LayoutDirection = value

    @property
    def layoutdirection(self):
        """Lower case alias for LayoutDirection"""
        return self.LayoutDirection

    @layoutdirection.setter
    def layoutdirection(self, value):
        """Lower case alias for LayoutDirection.setter"""
        self.LayoutDirection = value

    @property
    def Name(self):
        return self.com_object.Name

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def NoLineBreakAfter(self):
        return self.com_object.NoLineBreakAfter

    @NoLineBreakAfter.setter
    def NoLineBreakAfter(self, value):
        self.com_object.NoLineBreakAfter = value

    @property
    def nolinebreakafter(self):
        """Lower case alias for NoLineBreakAfter"""
        return self.NoLineBreakAfter

    @nolinebreakafter.setter
    def nolinebreakafter(self, value):
        """Lower case alias for NoLineBreakAfter.setter"""
        self.NoLineBreakAfter = value

    @property
    def NoLineBreakBefore(self):
        return self.com_object.NoLineBreakBefore

    @NoLineBreakBefore.setter
    def NoLineBreakBefore(self, value):
        self.com_object.NoLineBreakBefore = value

    @property
    def nolinebreakbefore(self):
        """Lower case alias for NoLineBreakBefore"""
        return self.NoLineBreakBefore

    @nolinebreakbefore.setter
    def nolinebreakbefore(self, value):
        """Lower case alias for NoLineBreakBefore.setter"""
        self.NoLineBreakBefore = value

    @property
    def NotesMaster(self):
        return Master(self.com_object.NotesMaster)

    @property
    def notesmaster(self):
        """Lower case alias for NotesMaster"""
        return self.NotesMaster

    @property
    def PageSetup(self):
        return PageSetup(self.com_object.PageSetup)

    @property
    def pagesetup(self):
        """Lower case alias for PageSetup"""
        return self.PageSetup

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Password(self):
        return self.com_object.Password

    @Password.setter
    def Password(self, value):
        self.com_object.Password = value

    @property
    def password(self):
        """Lower case alias for Password"""
        return self.Password

    @password.setter
    def password(self, value):
        """Lower case alias for Password.setter"""
        self.Password = value

    @property
    def PasswordEncryptionAlgorithm(self):
        return self.com_object.PasswordEncryptionAlgorithm

    @property
    def passwordencryptionalgorithm(self):
        """Lower case alias for PasswordEncryptionAlgorithm"""
        return self.PasswordEncryptionAlgorithm

    @property
    def PasswordEncryptionFileProperties(self):
        return self.com_object.PasswordEncryptionFileProperties

    @property
    def passwordencryptionfileproperties(self):
        """Lower case alias for PasswordEncryptionFileProperties"""
        return self.PasswordEncryptionFileProperties

    @property
    def PasswordEncryptionKeyLength(self):
        return self.com_object.PasswordEncryptionKeyLength

    @property
    def passwordencryptionkeylength(self):
        """Lower case alias for PasswordEncryptionKeyLength"""
        return self.PasswordEncryptionKeyLength

    @property
    def PasswordEncryptionProvider(self):
        return self.com_object.PasswordEncryptionProvider

    @property
    def passwordencryptionprovider(self):
        """Lower case alias for PasswordEncryptionProvider"""
        return self.PasswordEncryptionProvider

    @property
    def Path(self):
        return Presentation(self.com_object.Path)

    @property
    def path(self):
        """Lower case alias for Path"""
        return self.Path

    @property
    def PrintOptions(self):
        return PrintOptions(self.com_object.PrintOptions)

    @property
    def printoptions(self):
        """Lower case alias for PrintOptions"""
        return self.PrintOptions

    @property
    def ReadOnly(self):
        return self.com_object.ReadOnly

    @property
    def readonly(self):
        """Lower case alias for ReadOnly"""
        return self.ReadOnly

    @property
    def ReadOnlyRecommended(self):
        return self.com_object.ReadOnlyRecommended

    @property
    def readonlyrecommended(self):
        """Lower case alias for ReadOnlyRecommended"""
        return self.ReadOnlyRecommended

    @property
    def RemovePersonalInformation(self):
        return self.com_object.RemovePersonalInformation

    @RemovePersonalInformation.setter
    def RemovePersonalInformation(self, value):
        self.com_object.RemovePersonalInformation = value

    @property
    def removepersonalinformation(self):
        """Lower case alias for RemovePersonalInformation"""
        return self.RemovePersonalInformation

    @removepersonalinformation.setter
    def removepersonalinformation(self, value):
        """Lower case alias for RemovePersonalInformation.setter"""
        self.RemovePersonalInformation = value

    @property
    def Research(self):
        return Research(self.com_object.Research)

    @property
    def research(self):
        """Lower case alias for Research"""
        return self.Research

    @property
    def Saved(self):
        return self.com_object.Saved

    @Saved.setter
    def Saved(self, value):
        self.com_object.Saved = value

    @property
    def saved(self):
        """Lower case alias for Saved"""
        return self.Saved

    @saved.setter
    def saved(self, value):
        """Lower case alias for Saved.setter"""
        self.Saved = value

    @property
    def SectionProperties(self):
        return SectionProperties(self.com_object.SectionProperties)

    @property
    def sectionproperties(self):
        """Lower case alias for SectionProperties"""
        return self.SectionProperties

    @property
    def SensitivityLabel(self):
        return self.com_object.SensitivityLabel

    @property
    def sensitivitylabel(self):
        """Lower case alias for SensitivityLabel"""
        return self.SensitivityLabel

    @property
    def ServerPolicy(self):
        return self.com_object.ServerPolicy

    @property
    def serverpolicy(self):
        """Lower case alias for ServerPolicy"""
        return self.ServerPolicy

    @property
    def SharedWorkspace(self):
        return self.com_object.SharedWorkspace

    @property
    def sharedworkspace(self):
        """Lower case alias for SharedWorkspace"""
        return self.SharedWorkspace

    @property
    def Signatures(self):
        return self.com_object.Signatures

    @property
    def signatures(self):
        """Lower case alias for Signatures"""
        return self.Signatures

    @property
    def SlideMaster(self):
        return Master(self.com_object.SlideMaster)

    @property
    def slidemaster(self):
        """Lower case alias for SlideMaster"""
        return self.SlideMaster

    @property
    def Slides(self):
        return Slides(self.com_object.Slides)

    @property
    def slides(self):
        """Lower case alias for Slides"""
        return self.Slides

    @property
    def SlideShowSettings(self):
        return SlideShowSettings(self.com_object.SlideShowSettings)

    @property
    def slideshowsettings(self):
        """Lower case alias for SlideShowSettings"""
        return self.SlideShowSettings

    @property
    def SlideShowWindow(self):
        return SlideShowWindow(self.com_object.SlideShowWindow)

    @property
    def slideshowwindow(self):
        """Lower case alias for SlideShowWindow"""
        return self.SlideShowWindow

    @property
    def SnapToGrid(self):
        return self.com_object.SnapToGrid

    @SnapToGrid.setter
    def SnapToGrid(self, value):
        self.com_object.SnapToGrid = value

    @property
    def snaptogrid(self):
        """Lower case alias for SnapToGrid"""
        return self.SnapToGrid

    @snaptogrid.setter
    def snaptogrid(self, value):
        """Lower case alias for SnapToGrid.setter"""
        self.SnapToGrid = value

    @property
    def Sync(self):
        return self.com_object.Sync

    @property
    def sync(self):
        """Lower case alias for Sync"""
        return self.Sync

    @property
    def Tags(self):
        return Tags(self.com_object.Tags)

    @property
    def tags(self):
        """Lower case alias for Tags"""
        return self.Tags

    @property
    def TemplateName(self):
        return self.com_object.TemplateName

    @property
    def templatename(self):
        """Lower case alias for TemplateName"""
        return self.TemplateName

    @property
    def TitleMaster(self):
        return Master(self.com_object.TitleMaster)

    @property
    def titlemaster(self):
        """Lower case alias for TitleMaster"""
        return self.TitleMaster

    @property
    def VBASigned(self):
        return self.com_object.VBASigned

    @property
    def vbasigned(self):
        """Lower case alias for VBASigned"""
        return self.VBASigned

    @property
    def VBProject(self):
        return self.com_object.VBProject

    @property
    def vbproject(self):
        """Lower case alias for VBProject"""
        return self.VBProject

    @property
    def Windows(self):
        return DocumentWindows(self.com_object.Windows)

    @property
    def windows(self):
        """Lower case alias for Windows"""
        return self.Windows

    @property
    def WritePassword(self):
        return self.com_object.WritePassword

    @WritePassword.setter
    def WritePassword(self, value):
        self.com_object.WritePassword = value

    @property
    def writepassword(self):
        """Lower case alias for WritePassword"""
        return self.WritePassword

    @writepassword.setter
    def writepassword(self, value):
        """Lower case alias for WritePassword.setter"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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

    @property
    def activeprinter(self):
        """Lower case alias for ActivePrinter"""
        return self.ActivePrinter

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Collate(self):
        return self.com_object.Collate

    @Collate.setter
    def Collate(self, value):
        self.com_object.Collate = value

    @property
    def collate(self):
        """Lower case alias for Collate"""
        return self.Collate

    @collate.setter
    def collate(self, value):
        """Lower case alias for Collate.setter"""
        self.Collate = value

    @property
    def FitToPage(self):
        return self.com_object.FitToPage

    @FitToPage.setter
    def FitToPage(self, value):
        self.com_object.FitToPage = value

    @property
    def fittopage(self):
        """Lower case alias for FitToPage"""
        return self.FitToPage

    @fittopage.setter
    def fittopage(self, value):
        """Lower case alias for FitToPage.setter"""
        self.FitToPage = value

    @property
    def FrameSlides(self):
        return self.com_object.FrameSlides

    @FrameSlides.setter
    def FrameSlides(self, value):
        self.com_object.FrameSlides = value

    @property
    def frameslides(self):
        """Lower case alias for FrameSlides"""
        return self.FrameSlides

    @frameslides.setter
    def frameslides(self, value):
        """Lower case alias for FrameSlides.setter"""
        self.FrameSlides = value

    @property
    def HandoutOrder(self):
        return self.com_object.HandoutOrder

    @HandoutOrder.setter
    def HandoutOrder(self, value):
        self.com_object.HandoutOrder = value

    @property
    def handoutorder(self):
        """Lower case alias for HandoutOrder"""
        return self.HandoutOrder

    @handoutorder.setter
    def handoutorder(self, value):
        """Lower case alias for HandoutOrder.setter"""
        self.HandoutOrder = value

    @property
    def HighQuality(self):
        return self.com_object.HighQuality

    @HighQuality.setter
    def HighQuality(self, value):
        self.com_object.HighQuality = value

    @property
    def highquality(self):
        """Lower case alias for HighQuality"""
        return self.HighQuality

    @highquality.setter
    def highquality(self, value):
        """Lower case alias for HighQuality.setter"""
        self.HighQuality = value

    @property
    def NumberOfCopies(self):
        return self.com_object.NumberOfCopies

    @NumberOfCopies.setter
    def NumberOfCopies(self, value):
        self.com_object.NumberOfCopies = value

    @property
    def numberofcopies(self):
        """Lower case alias for NumberOfCopies"""
        return self.NumberOfCopies

    @numberofcopies.setter
    def numberofcopies(self, value):
        """Lower case alias for NumberOfCopies.setter"""
        self.NumberOfCopies = value

    @property
    def OutputType(self):
        return self.com_object.OutputType

    @OutputType.setter
    def OutputType(self, value):
        self.com_object.OutputType = value

    @property
    def outputtype(self):
        """Lower case alias for OutputType"""
        return self.OutputType

    @outputtype.setter
    def outputtype(self, value):
        """Lower case alias for OutputType.setter"""
        self.OutputType = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PrintColorType(self):
        return self.com_object.PrintColorType

    @PrintColorType.setter
    def PrintColorType(self, value):
        self.com_object.PrintColorType = value

    @property
    def printcolortype(self):
        """Lower case alias for PrintColorType"""
        return self.PrintColorType

    @printcolortype.setter
    def printcolortype(self, value):
        """Lower case alias for PrintColorType.setter"""
        self.PrintColorType = value

    @property
    def PrintComments(self):
        return self.com_object.PrintComments

    @PrintComments.setter
    def PrintComments(self, value):
        self.com_object.PrintComments = value

    @property
    def printcomments(self):
        """Lower case alias for PrintComments"""
        return self.PrintComments

    @printcomments.setter
    def printcomments(self, value):
        """Lower case alias for PrintComments.setter"""
        self.PrintComments = value

    @property
    def PrintFontsAsGraphics(self):
        return self.com_object.PrintFontsAsGraphics

    @PrintFontsAsGraphics.setter
    def PrintFontsAsGraphics(self, value):
        self.com_object.PrintFontsAsGraphics = value

    @property
    def printfontsasgraphics(self):
        """Lower case alias for PrintFontsAsGraphics"""
        return self.PrintFontsAsGraphics

    @printfontsasgraphics.setter
    def printfontsasgraphics(self, value):
        """Lower case alias for PrintFontsAsGraphics.setter"""
        self.PrintFontsAsGraphics = value

    @property
    def PrintHiddenSlides(self):
        return self.com_object.PrintHiddenSlides

    @PrintHiddenSlides.setter
    def PrintHiddenSlides(self, value):
        self.com_object.PrintHiddenSlides = value

    @property
    def printhiddenslides(self):
        """Lower case alias for PrintHiddenSlides"""
        return self.PrintHiddenSlides

    @printhiddenslides.setter
    def printhiddenslides(self, value):
        """Lower case alias for PrintHiddenSlides.setter"""
        self.PrintHiddenSlides = value

    @property
    def PrintInBackground(self):
        return self.com_object.PrintInBackground

    @PrintInBackground.setter
    def PrintInBackground(self, value):
        self.com_object.PrintInBackground = value

    @property
    def printinbackground(self):
        """Lower case alias for PrintInBackground"""
        return self.PrintInBackground

    @printinbackground.setter
    def printinbackground(self, value):
        """Lower case alias for PrintInBackground.setter"""
        self.PrintInBackground = value

    @property
    def Ranges(self):
        return PrintRanges(self.com_object.Ranges)

    @property
    def ranges(self):
        """Lower case alias for Ranges"""
        return self.Ranges

    @property
    def RangeType(self):
        return self.com_object.RangeType

    @RangeType.setter
    def RangeType(self, value):
        self.com_object.RangeType = value

    @property
    def rangetype(self):
        """Lower case alias for RangeType"""
        return self.RangeType

    @rangetype.setter
    def rangetype(self, value):
        """Lower case alias for RangeType.setter"""
        self.RangeType = value

    @property
    def sectionIndex(self):
        return PrintOptions(self.com_object.sectionIndex)

    @sectionIndex.setter
    def sectionIndex(self, value):
        self.com_object.sectionIndex = value

    @property
    def sectionindex(self):
        """Lower case alias for sectionIndex"""
        return self.sectionIndex

    @sectionindex.setter
    def sectionindex(self, value):
        """Lower case alias for sectionIndex.setter"""
        self.sectionIndex = value

    @property
    def SlideShowName(self):
        return self.com_object.SlideShowName

    @SlideShowName.setter
    def SlideShowName(self, value):
        self.com_object.SlideShowName = value

    @property
    def slideshowname(self):
        """Lower case alias for SlideShowName"""
        return self.SlideShowName

    @slideshowname.setter
    def slideshowname(self, value):
        """Lower case alias for SlideShowName.setter"""
        self.SlideShowName = value


class PrintRange:

    def __init__(self, printrange=None):
        self.com_object= printrange

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def End(self):
        return self.com_object.End

    @property
    def end(self):
        """Lower case alias for End"""
        return self.End

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Start(self):
        return self.com_object.Start

    @property
    def start(self):
        """Lower case alias for Start"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def From(self):
        return self.com_object.From

    @From.setter
    def From(self, value):
        self.com_object.From = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Points(self):
        return AnimationPoints(self.com_object.Points)

    @property
    def points(self):
        """Lower case alias for Points"""
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

    @property
    def to(self):
        """Lower case alias for To"""
        return self.To

    @to.setter
    def to(self, value):
        """Lower case alias for To.setter"""
        self.To = value


class ProtectedViewWindow:

    def __init__(self, protectedviewwindow=None):
        self.com_object= protectedviewwindow

    @property
    def Active(self):
        return self.com_object.Active

    @property
    def active(self):
        """Lower case alias for Active"""
        return self.Active

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Caption(self):
        return self.com_object.Caption

    @property
    def caption(self):
        """Lower case alias for Caption"""
        return self.Caption

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
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Presentation(self):
        return Presentation(self.com_object.Presentation)

    @property
    def presentation(self):
        """Lower case alias for Presentation"""
        return self.Presentation

    @property
    def SourceName(self):
        return ProtectedViewWindow(self.com_object.SourceName)

    @property
    def sourcename(self):
        """Lower case alias for SourceName"""
        return self.SourceName

    @property
    def SourcePath(self):
        return ProtectedViewWindow(self.com_object.SourcePath)

    @property
    def sourcepath(self):
        """Lower case alias for SourcePath"""
        return self.SourcePath

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
        return self.com_object.WindowState

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def FileName(self):
        return self.com_object.FileName

    @FileName.setter
    def FileName(self, value):
        self.com_object.FileName = value

    @property
    def filename(self):
        """Lower case alias for FileName"""
        return self.FileName

    @filename.setter
    def filename(self, value):
        """Lower case alias for FileName.setter"""
        self.FileName = value

    @property
    def HTMLVersion(self):
        return self.com_object.HTMLVersion

    @HTMLVersion.setter
    def HTMLVersion(self, value):
        self.com_object.HTMLVersion = value

    @property
    def htmlversion(self):
        """Lower case alias for HTMLVersion"""
        return self.HTMLVersion

    @htmlversion.setter
    def htmlversion(self, value):
        """Lower case alias for HTMLVersion.setter"""
        self.HTMLVersion = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def RangeEnd(self):
        return self.com_object.RangeEnd

    @RangeEnd.setter
    def RangeEnd(self, value):
        self.com_object.RangeEnd = value

    @property
    def rangeend(self):
        """Lower case alias for RangeEnd"""
        return self.RangeEnd

    @rangeend.setter
    def rangeend(self, value):
        """Lower case alias for RangeEnd.setter"""
        self.RangeEnd = value

    @property
    def RangeStart(self):
        return self.com_object.RangeStart

    @RangeStart.setter
    def RangeStart(self, value):
        self.com_object.RangeStart = value

    @property
    def rangestart(self):
        """Lower case alias for RangeStart"""
        return self.RangeStart

    @rangestart.setter
    def rangestart(self, value):
        """Lower case alias for RangeStart.setter"""
        self.RangeStart = value

    @property
    def SlideShowName(self):
        return self.com_object.SlideShowName

    @SlideShowName.setter
    def SlideShowName(self, value):
        self.com_object.SlideShowName = value

    @property
    def slideshowname(self):
        """Lower case alias for SlideShowName"""
        return self.SlideShowName

    @slideshowname.setter
    def slideshowname(self, value):
        """Lower case alias for SlideShowName.setter"""
        self.SlideShowName = value

    @property
    def SourceType(self):
        return self.com_object.SourceType

    @SourceType.setter
    def SourceType(self, value):
        self.com_object.SourceType = value

    @property
    def sourcetype(self):
        """Lower case alias for SourceType"""
        return self.SourceType

    @sourcetype.setter
    def sourcetype(self, value):
        """Lower case alias for SourceType.setter"""
        self.SourceType = value

    @property
    def SpeakerNotes(self):
        return self.com_object.SpeakerNotes

    @SpeakerNotes.setter
    def SpeakerNotes(self, value):
        self.com_object.SpeakerNotes = value

    @property
    def speakernotes(self):
        """Lower case alias for SpeakerNotes"""
        return self.SpeakerNotes

    @speakernotes.setter
    def speakernotes(self, value):
        """Lower case alias for SpeakerNotes.setter"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def PercentComplete(self):
        return self.com_object.PercentComplete

    @property
    def percentcomplete(self):
        """Lower case alias for PercentComplete"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def RGB(self):
        return PpColorSchemeIndex(self.com_object.RGB)

    @RGB.setter
    def RGB(self, value):
        self.com_object.RGB = value

    @property
    def rgb(self):
        """Lower case alias for RGB"""
        return self.RGB

    @rgb.setter
    def rgb(self, value):
        """Lower case alias for RGB.setter"""
        self.RGB = value


class RotationEffect:

    def __init__(self, rotationeffect=None):
        self.com_object= rotationeffect

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def By(self):
        return self.com_object.By

    @By.setter
    def By(self, value):
        self.com_object.By = value

    @property
    def by(self):
        """Lower case alias for By"""
        return self.By

    @by.setter
    def by(self, value):
        """Lower case alias for By.setter"""
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

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

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


class Row:

    def __init__(self, row=None):
        self.com_object= row

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    def Cells(self, RowIndex=None, ColumnIndex=None):
        arguments = com_arguments([unwrap(a) for a in [RowIndex, ColumnIndex]])
        if hasattr(self.com_object, "GetCells"):
            return CellRange(self.com_object.GetCells(*arguments))
        else:
            return CellRange(self.com_object.Cells(*arguments))

    def cells(self, RowIndex=None, ColumnIndex=None):
        """Lower case alias for Cells"""
        arguments = [RowIndex, ColumnIndex]
        return self.Cells(*arguments)

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
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Levels(self):
        return RulerLevels(self.com_object.Levels)

    @property
    def levels(self):
        """Lower case alias for Levels"""
        return self.Levels

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def TabStops(self):
        return TabStops(self.com_object.TabStops)

    @property
    def tabstops(self):
        """Lower case alias for TabStops"""
        return self.TabStops


class RulerLevel:

    def __init__(self, rulerlevel=None):
        self.com_object= rulerlevel

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def FirstMargin(self):
        return self.com_object.FirstMargin

    @FirstMargin.setter
    def FirstMargin(self, value):
        self.com_object.FirstMargin = value

    @property
    def firstmargin(self):
        """Lower case alias for FirstMargin"""
        return self.FirstMargin

    @firstmargin.setter
    def firstmargin(self, value):
        """Lower case alias for FirstMargin.setter"""
        self.FirstMargin = value

    @property
    def LeftMargin(self):
        return self.com_object.LeftMargin

    @LeftMargin.setter
    def LeftMargin(self, value):
        self.com_object.LeftMargin = value

    @property
    def leftmargin(self):
        """Lower case alias for LeftMargin"""
        return self.LeftMargin

    @leftmargin.setter
    def leftmargin(self, value):
        """Lower case alias for LeftMargin.setter"""
        self.LeftMargin = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def ByX(self):
        return self.com_object.ByX

    @ByX.setter
    def ByX(self, value):
        self.com_object.ByX = value

    @property
    def byx(self):
        """Lower case alias for ByX"""
        return self.ByX

    @byx.setter
    def byx(self, value):
        """Lower case alias for ByX.setter"""
        self.ByX = value

    @property
    def ByY(self):
        return self.com_object.ByY

    @ByY.setter
    def ByY(self, value):
        self.com_object.ByY = value

    @property
    def byy(self):
        """Lower case alias for ByY"""
        return self.ByY

    @byy.setter
    def byy(self, value):
        """Lower case alias for ByY.setter"""
        self.ByY = value

    @property
    def FromX(self):
        return self.com_object.FromX

    @FromX.setter
    def FromX(self, value):
        self.com_object.FromX = value

    @property
    def fromx(self):
        """Lower case alias for FromX"""
        return self.FromX

    @fromx.setter
    def fromx(self, value):
        """Lower case alias for FromX.setter"""
        self.FromX = value

    @property
    def FromY(self):
        return ScaleEffect(self.com_object.FromY)

    @FromY.setter
    def FromY(self, value):
        self.com_object.FromY = value

    @property
    def fromy(self):
        """Lower case alias for FromY"""
        return self.FromY

    @fromy.setter
    def fromy(self, value):
        """Lower case alias for FromY.setter"""
        self.FromY = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def ToX(self):
        return self.com_object.ToX

    @ToX.setter
    def ToX(self, value):
        self.com_object.ToX = value

    @property
    def tox(self):
        """Lower case alias for ToX"""
        return self.ToX

    @tox.setter
    def tox(self, value):
        """Lower case alias for ToX.setter"""
        self.ToX = value

    @property
    def ToY(self):
        return ScaleEffect(self.com_object.ToY)

    @ToY.setter
    def ToY(self, value):
        self.com_object.ToY = value

    @property
    def toy(self):
        """Lower case alias for ToY"""
        return self.ToY

    @toy.setter
    def toy(self, value):
        """Lower case alias for ToY.setter"""
        self.ToY = value


class SectionProperties:

    def __init__(self, sectionproperties=None):
        self.com_object= sectionproperties

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def ChildShapeRange(self):
        return ShapeRange(self.com_object.ChildShapeRange)

    @property
    def childshaperange(self):
        """Lower case alias for ChildShapeRange"""
        return self.ChildShapeRange

    @property
    def HasChildShapeRange(self):
        return self.com_object.HasChildShapeRange

    @property
    def haschildshaperange(self):
        """Lower case alias for HasChildShapeRange"""
        return self.HasChildShapeRange

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def ShapeRange(self):
        return ShapeRange(self.com_object.ShapeRange)

    @property
    def shaperange(self):
        """Lower case alias for ShapeRange"""
        return self.ShapeRange

    @property
    def SlideRange(self):
        return SlideRange(self.com_object.SlideRange)

    @property
    def sliderange(self):
        """Lower case alias for SlideRange"""
        return self.SlideRange

    @property
    def TextRange(self):
        return TextRange(self.com_object.TextRange)

    @property
    def textrange(self):
        """Lower case alias for TextRange"""
        return self.TextRange

    @property
    def TextRange2(self):
        return TextRange2(self.com_object.TextRange2)

    @property
    def textrange2(self):
        """Lower case alias for TextRange2"""
        return self.TextRange2

    @property
    def Type(self):
        return self.com_object.Type

    @property
    def type(self):
        """Lower case alias for Type"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def ApplyPictToEnd(self):
        return self.com_object.ApplyPictToEnd

    @ApplyPictToEnd.setter
    def ApplyPictToEnd(self, value):
        self.com_object.ApplyPictToEnd = value

    @property
    def applypicttoend(self):
        """Lower case alias for ApplyPictToEnd"""
        return self.ApplyPictToEnd

    @applypicttoend.setter
    def applypicttoend(self, value):
        """Lower case alias for ApplyPictToEnd.setter"""
        self.ApplyPictToEnd = value

    @property
    def ApplyPictToFront(self):
        return self.com_object.ApplyPictToFront

    @ApplyPictToFront.setter
    def ApplyPictToFront(self, value):
        self.com_object.ApplyPictToFront = value

    @property
    def applypicttofront(self):
        """Lower case alias for ApplyPictToFront"""
        return self.ApplyPictToFront

    @applypicttofront.setter
    def applypicttofront(self, value):
        """Lower case alias for ApplyPictToFront.setter"""
        self.ApplyPictToFront = value

    @property
    def ApplyPictToSides(self):
        return self.com_object.ApplyPictToSides

    @ApplyPictToSides.setter
    def ApplyPictToSides(self, value):
        self.com_object.ApplyPictToSides = value

    @property
    def applypicttosides(self):
        """Lower case alias for ApplyPictToSides"""
        return self.ApplyPictToSides

    @applypicttosides.setter
    def applypicttosides(self, value):
        """Lower case alias for ApplyPictToSides.setter"""
        self.ApplyPictToSides = value

    @property
    def AxisGroup(self):
        return XlAxisGroup(self.com_object.AxisGroup)

    @AxisGroup.setter
    def AxisGroup(self, value):
        self.com_object.AxisGroup = value

    @property
    def axisgroup(self):
        """Lower case alias for AxisGroup"""
        return self.AxisGroup

    @axisgroup.setter
    def axisgroup(self, value):
        """Lower case alias for AxisGroup.setter"""
        self.AxisGroup = value

    @property
    def BarShape(self):
        return XlBarShape(self.com_object.BarShape)

    @BarShape.setter
    def BarShape(self, value):
        self.com_object.BarShape = value

    @property
    def barshape(self):
        """Lower case alias for BarShape"""
        return self.BarShape

    @barshape.setter
    def barshape(self, value):
        """Lower case alias for BarShape.setter"""
        self.BarShape = value

    @property
    def BubbleSizes(self):
        return self.com_object.BubbleSizes

    @BubbleSizes.setter
    def BubbleSizes(self, value):
        self.com_object.BubbleSizes = value

    @property
    def bubblesizes(self):
        """Lower case alias for BubbleSizes"""
        return self.BubbleSizes

    @bubblesizes.setter
    def bubblesizes(self, value):
        """Lower case alias for BubbleSizes.setter"""
        self.BubbleSizes = value

    @property
    def ChartType(self):
        return self.com_object.ChartType

    @ChartType.setter
    def ChartType(self, value):
        self.com_object.ChartType = value

    @property
    def charttype(self):
        """Lower case alias for ChartType"""
        return self.ChartType

    @charttype.setter
    def charttype(self, value):
        """Lower case alias for ChartType.setter"""
        self.ChartType = value

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def ErrorBars(self):
        return ErrorBars(self.com_object.ErrorBars)

    @property
    def errorbars(self):
        """Lower case alias for ErrorBars"""
        return self.ErrorBars

    @property
    def Explosion(self):
        return self.com_object.Explosion

    @Explosion.setter
    def Explosion(self, value):
        self.com_object.Explosion = value

    @property
    def explosion(self):
        """Lower case alias for Explosion"""
        return self.Explosion

    @explosion.setter
    def explosion(self, value):
        """Lower case alias for Explosion.setter"""
        self.Explosion = value

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

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
    def FormulaLocal(self):
        return self.com_object.FormulaLocal

    @FormulaLocal.setter
    def FormulaLocal(self, value):
        self.com_object.FormulaLocal = value

    @property
    def formulalocal(self):
        """Lower case alias for FormulaLocal"""
        return self.FormulaLocal

    @formulalocal.setter
    def formulalocal(self, value):
        """Lower case alias for FormulaLocal.setter"""
        self.FormulaLocal = value

    @property
    def FormulaR1C1(self):
        return self.com_object.FormulaR1C1

    @FormulaR1C1.setter
    def FormulaR1C1(self, value):
        self.com_object.FormulaR1C1 = value

    @property
    def formular1c1(self):
        """Lower case alias for FormulaR1C1"""
        return self.FormulaR1C1

    @formular1c1.setter
    def formular1c1(self, value):
        """Lower case alias for FormulaR1C1.setter"""
        self.FormulaR1C1 = value

    @property
    def FormulaR1C1Local(self):
        return self.com_object.FormulaR1C1Local

    @FormulaR1C1Local.setter
    def FormulaR1C1Local(self, value):
        self.com_object.FormulaR1C1Local = value

    @property
    def formular1c1local(self):
        """Lower case alias for FormulaR1C1Local"""
        return self.FormulaR1C1Local

    @formular1c1local.setter
    def formular1c1local(self, value):
        """Lower case alias for FormulaR1C1Local.setter"""
        self.FormulaR1C1Local = value

    @property
    def Has3DEffect(self):
        return self.com_object.Has3DEffect

    @Has3DEffect.setter
    def Has3DEffect(self, value):
        self.com_object.Has3DEffect = value

    @property
    def has3deffect(self):
        """Lower case alias for Has3DEffect"""
        return self.Has3DEffect

    @has3deffect.setter
    def has3deffect(self, value):
        """Lower case alias for Has3DEffect.setter"""
        self.Has3DEffect = value

    @property
    def HasDataLabels(self):
        return self.com_object.HasDataLabels

    @HasDataLabels.setter
    def HasDataLabels(self, value):
        self.com_object.HasDataLabels = value

    @property
    def hasdatalabels(self):
        """Lower case alias for HasDataLabels"""
        return self.HasDataLabels

    @hasdatalabels.setter
    def hasdatalabels(self, value):
        """Lower case alias for HasDataLabels.setter"""
        self.HasDataLabels = value

    @property
    def HasErrorBars(self):
        return self.com_object.HasErrorBars

    @HasErrorBars.setter
    def HasErrorBars(self, value):
        self.com_object.HasErrorBars = value

    @property
    def haserrorbars(self):
        """Lower case alias for HasErrorBars"""
        return self.HasErrorBars

    @haserrorbars.setter
    def haserrorbars(self, value):
        """Lower case alias for HasErrorBars.setter"""
        self.HasErrorBars = value

    @property
    def HasLeaderLines(self):
        return self.com_object.HasLeaderLines

    @HasLeaderLines.setter
    def HasLeaderLines(self, value):
        self.com_object.HasLeaderLines = value

    @property
    def hasleaderlines(self):
        """Lower case alias for HasLeaderLines"""
        return self.HasLeaderLines

    @hasleaderlines.setter
    def hasleaderlines(self, value):
        """Lower case alias for HasLeaderLines.setter"""
        self.HasLeaderLines = value

    @property
    def InvertColor(self):
        return self.com_object.InvertColor

    @InvertColor.setter
    def InvertColor(self, value):
        self.com_object.InvertColor = value

    @property
    def invertcolor(self):
        """Lower case alias for InvertColor"""
        return self.InvertColor

    @invertcolor.setter
    def invertcolor(self, value):
        """Lower case alias for InvertColor.setter"""
        self.InvertColor = value

    @property
    def InvertColorIndex(self):
        return self.com_object.InvertColorIndex

    @InvertColorIndex.setter
    def InvertColorIndex(self, value):
        self.com_object.InvertColorIndex = value

    @property
    def invertcolorindex(self):
        """Lower case alias for InvertColorIndex"""
        return self.InvertColorIndex

    @invertcolorindex.setter
    def invertcolorindex(self, value):
        """Lower case alias for InvertColorIndex.setter"""
        self.InvertColorIndex = value

    @property
    def InvertIfNegative(self):
        return self.com_object.InvertIfNegative

    @InvertIfNegative.setter
    def InvertIfNegative(self, value):
        self.com_object.InvertIfNegative = value

    @property
    def invertifnegative(self):
        """Lower case alias for InvertIfNegative"""
        return self.InvertIfNegative

    @invertifnegative.setter
    def invertifnegative(self, value):
        """Lower case alias for InvertIfNegative.setter"""
        self.InvertIfNegative = value

    @property
    def isfiltered(self):
        return self.com_object.isfiltered

    @isfiltered.setter
    def isfiltered(self, value):
        self.com_object.isfiltered = value

    @property
    def LeaderLines(self):
        return LeaderLines(self.com_object.LeaderLines)

    @property
    def leaderlines(self):
        """Lower case alias for LeaderLines"""
        return self.LeaderLines

    @property
    def MarkerBackgroundColor(self):
        return self.com_object.MarkerBackgroundColor

    @MarkerBackgroundColor.setter
    def MarkerBackgroundColor(self, value):
        self.com_object.MarkerBackgroundColor = value

    @property
    def markerbackgroundcolor(self):
        """Lower case alias for MarkerBackgroundColor"""
        return self.MarkerBackgroundColor

    @markerbackgroundcolor.setter
    def markerbackgroundcolor(self, value):
        """Lower case alias for MarkerBackgroundColor.setter"""
        self.MarkerBackgroundColor = value

    @property
    def MarkerBackgroundColorIndex(self):
        return XlColorIndex(self.com_object.MarkerBackgroundColorIndex)

    @MarkerBackgroundColorIndex.setter
    def MarkerBackgroundColorIndex(self, value):
        self.com_object.MarkerBackgroundColorIndex = value

    @property
    def markerbackgroundcolorindex(self):
        """Lower case alias for MarkerBackgroundColorIndex"""
        return self.MarkerBackgroundColorIndex

    @markerbackgroundcolorindex.setter
    def markerbackgroundcolorindex(self, value):
        """Lower case alias for MarkerBackgroundColorIndex.setter"""
        self.MarkerBackgroundColorIndex = value

    @property
    def MarkerForegroundColor(self):
        return self.com_object.MarkerForegroundColor

    @MarkerForegroundColor.setter
    def MarkerForegroundColor(self, value):
        self.com_object.MarkerForegroundColor = value

    @property
    def markerforegroundcolor(self):
        """Lower case alias for MarkerForegroundColor"""
        return self.MarkerForegroundColor

    @markerforegroundcolor.setter
    def markerforegroundcolor(self, value):
        """Lower case alias for MarkerForegroundColor.setter"""
        self.MarkerForegroundColor = value

    @property
    def MarkerForegroundColorIndex(self):
        return XlColorIndex(self.com_object.MarkerForegroundColorIndex)

    @MarkerForegroundColorIndex.setter
    def MarkerForegroundColorIndex(self, value):
        self.com_object.MarkerForegroundColorIndex = value

    @property
    def markerforegroundcolorindex(self):
        """Lower case alias for MarkerForegroundColorIndex"""
        return self.MarkerForegroundColorIndex

    @markerforegroundcolorindex.setter
    def markerforegroundcolorindex(self, value):
        """Lower case alias for MarkerForegroundColorIndex.setter"""
        self.MarkerForegroundColorIndex = value

    @property
    def MarkerSize(self):
        return self.com_object.MarkerSize

    @MarkerSize.setter
    def MarkerSize(self, value):
        self.com_object.MarkerSize = value

    @property
    def markersize(self):
        """Lower case alias for MarkerSize"""
        return self.MarkerSize

    @markersize.setter
    def markersize(self, value):
        """Lower case alias for MarkerSize.setter"""
        self.MarkerSize = value

    @property
    def MarkerStyle(self):
        return XlMarkerStyle(self.com_object.MarkerStyle)

    @MarkerStyle.setter
    def MarkerStyle(self, value):
        self.com_object.MarkerStyle = value

    @property
    def markerstyle(self):
        """Lower case alias for MarkerStyle"""
        return self.MarkerStyle

    @markerstyle.setter
    def markerstyle(self, value):
        """Lower case alias for MarkerStyle.setter"""
        self.MarkerStyle = value

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
    def parentdatalabeloption(self):
        return self.com_object.parentdatalabeloption

    @parentdatalabeloption.setter
    def parentdatalabeloption(self, value):
        self.com_object.parentdatalabeloption = value

    @property
    def PictureType(self):
        return XlChartPictureType(self.com_object.PictureType)

    @PictureType.setter
    def PictureType(self, value):
        self.com_object.PictureType = value

    @property
    def picturetype(self):
        """Lower case alias for PictureType"""
        return self.PictureType

    @picturetype.setter
    def picturetype(self, value):
        """Lower case alias for PictureType.setter"""
        self.PictureType = value

    @property
    def PictureUnit2(self):
        return self.com_object.PictureUnit2

    @PictureUnit2.setter
    def PictureUnit2(self, value):
        self.com_object.PictureUnit2 = value

    @property
    def pictureunit2(self):
        """Lower case alias for PictureUnit2"""
        return self.PictureUnit2

    @pictureunit2.setter
    def pictureunit2(self, value):
        """Lower case alias for PictureUnit2.setter"""
        self.PictureUnit2 = value

    @property
    def PlotColorIndex(self):
        return self.com_object.PlotColorIndex

    @property
    def plotcolorindex(self):
        """Lower case alias for PlotColorIndex"""
        return self.PlotColorIndex

    @property
    def PlotOrder(self):
        return self.com_object.PlotOrder

    @PlotOrder.setter
    def PlotOrder(self, value):
        self.com_object.PlotOrder = value

    @property
    def plotorder(self):
        """Lower case alias for PlotOrder"""
        return self.PlotOrder

    @plotorder.setter
    def plotorder(self, value):
        """Lower case alias for PlotOrder.setter"""
        self.PlotOrder = value

    @property
    def quartilecalculationinclusivemedian(self):
        return self.com_object.quartilecalculationinclusivemedian

    @quartilecalculationinclusivemedian.setter
    def quartilecalculationinclusivemedian(self, value):
        self.com_object.quartilecalculationinclusivemedian = value

    @property
    def Shadow(self):
        return self.com_object.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.com_object.Shadow = value

    @property
    def shadow(self):
        """Lower case alias for Shadow"""
        return self.Shadow

    @shadow.setter
    def shadow(self, value):
        """Lower case alias for Shadow.setter"""
        self.Shadow = value

    @property
    def Smooth(self):
        return self.com_object.Smooth

    @Smooth.setter
    def Smooth(self, value):
        self.com_object.Smooth = value

    @property
    def smooth(self):
        """Lower case alias for Smooth"""
        return self.Smooth

    @smooth.setter
    def smooth(self, value):
        """Lower case alias for Smooth.setter"""
        self.Smooth = value

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
    def Values(self):
        return self.com_object.Values

    @Values.setter
    def Values(self, value):
        self.com_object.Values = value

    @property
    def values(self):
        """Lower case alias for Values"""
        return self.Values

    @values.setter
    def values(self, value):
        """Lower case alias for Values.setter"""
        self.Values = value

    @property
    def XValues(self):
        return self.com_object.XValues

    @XValues.setter
    def XValues(self, value):
        self.com_object.XValues = value

    @property
    def xvalues(self):
        """Lower case alias for XValues"""
        return self.XValues

    @xvalues.setter
    def xvalues(self, value):
        """Lower case alias for XValues.setter"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Count(self):
        return self.com_object.Count

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Border(self):
        return ChartBorder(self.com_object.Border)

    @property
    def border(self):
        """Lower case alias for Border"""
        return self.Border

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
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

    @property
    def to(self):
        """Lower case alias for To"""
        return self.To

    @to.setter
    def to(self, value):
        """Lower case alias for To.setter"""
        self.To = value


class ShadowFormat:

    def __init__(self, shadowformat=None):
        self.com_object= shadowformat

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Blur(self):
        return self.com_object.Blur

    @Blur.setter
    def Blur(self, value):
        self.com_object.Blur = value

    @property
    def blur(self):
        """Lower case alias for Blur"""
        return self.Blur

    @blur.setter
    def blur(self, value):
        """Lower case alias for Blur.setter"""
        self.Blur = value

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def ForeColor(self):
        return ColorFormat(self.com_object.ForeColor)

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
    def Obscured(self):
        return self.com_object.Obscured

    @Obscured.setter
    def Obscured(self, value):
        self.com_object.Obscured = value

    @property
    def obscured(self):
        """Lower case alias for Obscured"""
        return self.Obscured

    @obscured.setter
    def obscured(self, value):
        """Lower case alias for Obscured.setter"""
        self.Obscured = value

    @property
    def OffsetX(self):
        return self.com_object.OffsetX

    @OffsetX.setter
    def OffsetX(self, value):
        self.com_object.OffsetX = value

    @property
    def offsetx(self):
        """Lower case alias for OffsetX"""
        return self.OffsetX

    @offsetx.setter
    def offsetx(self, value):
        """Lower case alias for OffsetX.setter"""
        self.OffsetX = value

    @property
    def OffsetY(self):
        return self.com_object.OffsetY

    @OffsetY.setter
    def OffsetY(self, value):
        self.com_object.OffsetY = value

    @property
    def offsety(self):
        """Lower case alias for OffsetY"""
        return self.OffsetY

    @offsety.setter
    def offsety(self, value):
        """Lower case alias for OffsetY.setter"""
        self.OffsetY = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def RotateWithShape(self):
        return self.com_object.RotateWithShape

    @RotateWithShape.setter
    def RotateWithShape(self, value):
        self.com_object.RotateWithShape = value

    @property
    def rotatewithshape(self):
        """Lower case alias for RotateWithShape"""
        return self.RotateWithShape

    @rotatewithshape.setter
    def rotatewithshape(self, value):
        """Lower case alias for RotateWithShape.setter"""
        self.RotateWithShape = value

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
    def Style(self):
        return self.com_object.Style

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
    def Transparency(self):
        return self.com_object.Transparency

    @Transparency.setter
    def Transparency(self, value):
        self.com_object.Transparency = value

    @property
    def transparency(self):
        """Lower case alias for Transparency"""
        return self.Transparency

    @transparency.setter
    def transparency(self, value):
        """Lower case alias for Transparency.setter"""
        self.Transparency = value

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

    @property
    def actionsettings(self):
        """Lower case alias for ActionSettings"""
        return self.ActionSettings

    @property
    def Adjustments(self):
        return Adjustments(self.com_object.Adjustments)

    @property
    def adjustments(self):
        """Lower case alias for Adjustments"""
        return self.Adjustments

    @property
    def AlternativeText(self):
        return self.com_object.AlternativeText

    @AlternativeText.setter
    def AlternativeText(self, value):
        self.com_object.AlternativeText = value

    @property
    def alternativetext(self):
        """Lower case alias for AlternativeText"""
        return self.AlternativeText

    @alternativetext.setter
    def alternativetext(self, value):
        """Lower case alias for AlternativeText.setter"""
        self.AlternativeText = value

    @property
    def AnimationSettings(self):
        return AnimationSettings(self.com_object.AnimationSettings)

    @property
    def animationsettings(self):
        """Lower case alias for AnimationSettings"""
        return self.AnimationSettings

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def AutoShapeType(self):
        return Shape(self.com_object.AutoShapeType)

    @AutoShapeType.setter
    def AutoShapeType(self, value):
        self.com_object.AutoShapeType = value

    @property
    def autoshapetype(self):
        """Lower case alias for AutoShapeType"""
        return self.AutoShapeType

    @autoshapetype.setter
    def autoshapetype(self, value):
        """Lower case alias for AutoShapeType.setter"""
        self.AutoShapeType = value

    @property
    def BackgroundStyle(self):
        return self.com_object.BackgroundStyle

    @BackgroundStyle.setter
    def BackgroundStyle(self, value):
        self.com_object.BackgroundStyle = value

    @property
    def backgroundstyle(self):
        """Lower case alias for BackgroundStyle"""
        return self.BackgroundStyle

    @backgroundstyle.setter
    def backgroundstyle(self, value):
        """Lower case alias for BackgroundStyle.setter"""
        self.BackgroundStyle = value

    @property
    def BlackWhiteMode(self):
        return self.com_object.BlackWhiteMode

    @BlackWhiteMode.setter
    def BlackWhiteMode(self, value):
        self.com_object.BlackWhiteMode = value

    @property
    def blackwhitemode(self):
        """Lower case alias for BlackWhiteMode"""
        return self.BlackWhiteMode

    @blackwhitemode.setter
    def blackwhitemode(self, value):
        """Lower case alias for BlackWhiteMode.setter"""
        self.BlackWhiteMode = value

    @property
    def Callout(self):
        return CalloutFormat(self.com_object.Callout)

    @property
    def callout(self):
        """Lower case alias for Callout"""
        return self.Callout

    @property
    def Chart(self):
        return Chart(self.com_object.Chart)

    @property
    def chart(self):
        """Lower case alias for Chart"""
        return self.Chart

    @property
    def Child(self):
        return self.com_object.Child

    @property
    def child(self):
        """Lower case alias for Child"""
        return self.Child

    @property
    def ConnectionSiteCount(self):
        return self.com_object.ConnectionSiteCount

    @property
    def connectionsitecount(self):
        """Lower case alias for ConnectionSiteCount"""
        return self.ConnectionSiteCount

    @property
    def Connector(self):
        return self.com_object.Connector

    @property
    def connector(self):
        """Lower case alias for Connector"""
        return self.Connector

    @property
    def ConnectorFormat(self):
        return ConnectorFormat(self.com_object.ConnectorFormat)

    @property
    def connectorformat(self):
        """Lower case alias for ConnectorFormat"""
        return self.ConnectorFormat

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def CustomerData(self):
        return CustomerData(self.com_object.CustomerData)

    @property
    def customerdata(self):
        """Lower case alias for CustomerData"""
        return self.CustomerData

    @property
    def Decorative(self):
        return self.com_object.Decorative

    @Decorative.setter
    def Decorative(self, value):
        self.com_object.Decorative = value

    @property
    def decorative(self):
        """Lower case alias for Decorative"""
        return self.Decorative

    @decorative.setter
    def decorative(self, value):
        """Lower case alias for Decorative.setter"""
        self.Decorative = value

    @property
    def Fill(self):
        return FillFormat(self.com_object.Fill)

    @property
    def fill(self):
        """Lower case alias for Fill"""
        return self.Fill

    @property
    def Glow(self):
        return self.com_object.Glow

    @property
    def glow(self):
        """Lower case alias for Glow"""
        return self.Glow

    @property
    def GraphicStyle(self):
        return self.com_object.GraphicStyle

    @GraphicStyle.setter
    def GraphicStyle(self, value):
        self.com_object.GraphicStyle = value

    @property
    def graphicstyle(self):
        """Lower case alias for GraphicStyle"""
        return self.GraphicStyle

    @graphicstyle.setter
    def graphicstyle(self, value):
        """Lower case alias for GraphicStyle.setter"""
        self.GraphicStyle = value

    @property
    def GroupItems(self):
        return GroupShapes(self.com_object.GroupItems)

    @property
    def groupitems(self):
        """Lower case alias for GroupItems"""
        return self.GroupItems

    @property
    def HasChart(self):
        return self.com_object.HasChart

    @property
    def haschart(self):
        """Lower case alias for HasChart"""
        return self.HasChart

    @property
    def hasinkxml(self):
        return self.com_object.hasinkxml

    @property
    def HasSmartArt(self):
        return self.com_object.HasSmartArt

    @property
    def hassmartart(self):
        """Lower case alias for HasSmartArt"""
        return self.HasSmartArt

    @property
    def HasTable(self):
        return self.com_object.HasTable

    @property
    def hastable(self):
        """Lower case alias for HasTable"""
        return self.HasTable

    @property
    def HasTextFrame(self):
        return self.com_object.HasTextFrame

    @property
    def hastextframe(self):
        """Lower case alias for HasTextFrame"""
        return self.HasTextFrame

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
    def HorizontalFlip(self):
        return self.com_object.HorizontalFlip

    @property
    def horizontalflip(self):
        """Lower case alias for HorizontalFlip"""
        return self.HorizontalFlip

    @property
    def Id(self):
        return self.com_object.Id

    @property
    def id(self):
        """Lower case alias for Id"""
        return self.Id

    @property
    def inkxml(self):
        return self.com_object.inkxml

    @property
    def isnarration(self):
        return self.com_object.isnarration

    @isnarration.setter
    def isnarration(self, value):
        self.com_object.isnarration = value

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
    def Line(self):
        return LineFormat(self.com_object.Line)

    @property
    def line(self):
        """Lower case alias for Line"""
        return self.Line

    @property
    def LinkFormat(self):
        return LinkFormat(self.com_object.LinkFormat)

    @property
    def linkformat(self):
        """Lower case alias for LinkFormat"""
        return self.LinkFormat

    @property
    def LockAspectRatio(self):
        return self.com_object.LockAspectRatio

    @LockAspectRatio.setter
    def LockAspectRatio(self, value):
        self.com_object.LockAspectRatio = value

    @property
    def lockaspectratio(self):
        """Lower case alias for LockAspectRatio"""
        return self.LockAspectRatio

    @lockaspectratio.setter
    def lockaspectratio(self, value):
        """Lower case alias for LockAspectRatio.setter"""
        self.LockAspectRatio = value

    @property
    def MediaFormat(self):
        return self.com_object.MediaFormat

    @property
    def mediaformat(self):
        """Lower case alias for MediaFormat"""
        return self.MediaFormat

    @property
    def MediaType(self):
        return self.com_object.MediaType

    @property
    def mediatype(self):
        """Lower case alias for MediaType"""
        return self.MediaType

    @property
    def Model3D(self):
        return Model3DFormat(self.com_object.Model3D)

    @property
    def model3d(self):
        """Lower case alias for Model3D"""
        return self.Model3D

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
    def Nodes(self):
        return ShapeNodes(self.com_object.Nodes)

    @property
    def nodes(self):
        """Lower case alias for Nodes"""
        return self.Nodes

    @property
    def OLEFormat(self):
        return OLEFormat(self.com_object.OLEFormat)

    @property
    def oleformat(self):
        """Lower case alias for OLEFormat"""
        return self.OLEFormat

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def ParentGroup(self):
        return Shape(self.com_object.ParentGroup)

    @property
    def parentgroup(self):
        """Lower case alias for ParentGroup"""
        return self.ParentGroup

    @property
    def PictureFormat(self):
        return PictureFormat(self.com_object.PictureFormat)

    @property
    def pictureformat(self):
        """Lower case alias for PictureFormat"""
        return self.PictureFormat

    @property
    def PlaceholderFormat(self):
        return PlaceholderFormat(self.com_object.PlaceholderFormat)

    @property
    def placeholderformat(self):
        """Lower case alias for PlaceholderFormat"""
        return self.PlaceholderFormat

    @property
    def Reflection(self):
        return self.com_object.Reflection

    @property
    def reflection(self):
        """Lower case alias for Reflection"""
        return self.Reflection

    @property
    def Rotation(self):
        return self.com_object.Rotation

    @Rotation.setter
    def Rotation(self, value):
        self.com_object.Rotation = value

    @property
    def rotation(self):
        """Lower case alias for Rotation"""
        return self.Rotation

    @rotation.setter
    def rotation(self, value):
        """Lower case alias for Rotation.setter"""
        self.Rotation = value

    @property
    def Shadow(self):
        return ShadowFormat(self.com_object.Shadow)

    @property
    def shadow(self):
        """Lower case alias for Shadow"""
        return self.Shadow

    @property
    def ShapeStyle(self):
        return self.com_object.ShapeStyle

    @ShapeStyle.setter
    def ShapeStyle(self, value):
        self.com_object.ShapeStyle = value

    @property
    def shapestyle(self):
        """Lower case alias for ShapeStyle"""
        return self.ShapeStyle

    @shapestyle.setter
    def shapestyle(self, value):
        """Lower case alias for ShapeStyle.setter"""
        self.ShapeStyle = value

    @property
    def SmartArt(self):
        return Shape(self.com_object.SmartArt)

    @property
    def smartart(self):
        """Lower case alias for SmartArt"""
        return self.SmartArt

    @property
    def SoftEdge(self):
        return self.com_object.SoftEdge

    @property
    def softedge(self):
        """Lower case alias for SoftEdge"""
        return self.SoftEdge

    @property
    def Table(self):
        return Table(self.com_object.Table)

    @property
    def table(self):
        """Lower case alias for Table"""
        return self.Table

    @property
    def Tags(self):
        return Tags(self.com_object.Tags)

    @property
    def tags(self):
        """Lower case alias for Tags"""
        return self.Tags

    @property
    def TextEffect(self):
        return TextEffectFormat(self.com_object.TextEffect)

    @property
    def texteffect(self):
        """Lower case alias for TextEffect"""
        return self.TextEffect

    @property
    def TextFrame(self):
        return TextFrame(self.com_object.TextFrame)

    @property
    def textframe(self):
        """Lower case alias for TextFrame"""
        return self.TextFrame

    @property
    def TextFrame2(self):
        return TextFrame2(self.com_object.TextFrame2)

    @property
    def textframe2(self):
        """Lower case alias for TextFrame2"""
        return self.TextFrame2

    @property
    def ThreeD(self):
        return ThreeDFormat(self.com_object.ThreeD)

    @property
    def threed(self):
        """Lower case alias for ThreeD"""
        return self.ThreeD

    @property
    def Title(self):
        return Shape(self.com_object.Title)

    @property
    def title(self):
        """Lower case alias for Title"""
        return self.Title

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
    def Type(self):
        return self.com_object.Type

    @property
    def type(self):
        """Lower case alias for Type"""
        return self.Type

    @property
    def VerticalFlip(self):
        return self.com_object.VerticalFlip

    @property
    def verticalflip(self):
        """Lower case alias for VerticalFlip"""
        return self.VerticalFlip

    @property
    def Vertices(self):
        return self.com_object.Vertices

    @property
    def vertices(self):
        """Lower case alias for Vertices"""
        return self.Vertices

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
    def ZOrderPosition(self):
        return self.com_object.ZOrderPosition

    @property
    def zorderposition(self):
        """Lower case alias for ZOrderPosition"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def EditingType(self):
        return self.com_object.EditingType

    @property
    def editingtype(self):
        """Lower case alias for EditingType"""
        return self.EditingType

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Points(self):
        return self.com_object.Points

    @property
    def points(self):
        """Lower case alias for Points"""
        return self.Points

    @property
    def SegmentType(self):
        return self.com_object.SegmentType

    @property
    def segmenttype(self):
        """Lower case alias for SegmentType"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Count(self):
        return self.com_object.Count

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
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

    @property
    def actionsettings(self):
        """Lower case alias for ActionSettings"""
        return self.ActionSettings

    @property
    def Adjustments(self):
        return Adjustments(self.com_object.Adjustments)

    @property
    def adjustments(self):
        """Lower case alias for Adjustments"""
        return self.Adjustments

    @property
    def AlternativeText(self):
        return self.com_object.AlternativeText

    @AlternativeText.setter
    def AlternativeText(self, value):
        self.com_object.AlternativeText = value

    @property
    def alternativetext(self):
        """Lower case alias for AlternativeText"""
        return self.AlternativeText

    @alternativetext.setter
    def alternativetext(self, value):
        """Lower case alias for AlternativeText.setter"""
        self.AlternativeText = value

    @property
    def AnimationSettings(self):
        return AnimationSettings(self.com_object.AnimationSettings)

    @property
    def animationsettings(self):
        """Lower case alias for AnimationSettings"""
        return self.AnimationSettings

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def AutoShapeType(self):
        return ShapeRange(self.com_object.AutoShapeType)

    @AutoShapeType.setter
    def AutoShapeType(self, value):
        self.com_object.AutoShapeType = value

    @property
    def autoshapetype(self):
        """Lower case alias for AutoShapeType"""
        return self.AutoShapeType

    @autoshapetype.setter
    def autoshapetype(self, value):
        """Lower case alias for AutoShapeType.setter"""
        self.AutoShapeType = value

    @property
    def BackgroundStyle(self):
        return self.com_object.BackgroundStyle

    @BackgroundStyle.setter
    def BackgroundStyle(self, value):
        self.com_object.BackgroundStyle = value

    @property
    def backgroundstyle(self):
        """Lower case alias for BackgroundStyle"""
        return self.BackgroundStyle

    @backgroundstyle.setter
    def backgroundstyle(self, value):
        """Lower case alias for BackgroundStyle.setter"""
        self.BackgroundStyle = value

    @property
    def BlackWhiteMode(self):
        return self.com_object.BlackWhiteMode

    @BlackWhiteMode.setter
    def BlackWhiteMode(self, value):
        self.com_object.BlackWhiteMode = value

    @property
    def blackwhitemode(self):
        """Lower case alias for BlackWhiteMode"""
        return self.BlackWhiteMode

    @blackwhitemode.setter
    def blackwhitemode(self, value):
        """Lower case alias for BlackWhiteMode.setter"""
        self.BlackWhiteMode = value

    @property
    def Callout(self):
        return CalloutFormat(self.com_object.Callout)

    @property
    def callout(self):
        """Lower case alias for Callout"""
        return self.Callout

    @property
    def Chart(self):
        return Chart(self.com_object.Chart)

    @property
    def chart(self):
        """Lower case alias for Chart"""
        return self.Chart

    @property
    def Child(self):
        return self.com_object.Child

    @property
    def child(self):
        """Lower case alias for Child"""
        return self.Child

    @property
    def ConnectionSiteCount(self):
        return self.com_object.ConnectionSiteCount

    @property
    def connectionsitecount(self):
        """Lower case alias for ConnectionSiteCount"""
        return self.ConnectionSiteCount

    @property
    def Connector(self):
        return self.com_object.Connector

    @property
    def connector(self):
        """Lower case alias for Connector"""
        return self.Connector

    @property
    def ConnectorFormat(self):
        return ConnectorFormat(self.com_object.ConnectorFormat)

    @property
    def connectorformat(self):
        """Lower case alias for ConnectorFormat"""
        return self.ConnectorFormat

    @property
    def Count(self):
        return self.com_object.Count

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def CustomerData(self):
        return CustomerData(self.com_object.CustomerData)

    @property
    def customerdata(self):
        """Lower case alias for CustomerData"""
        return self.CustomerData

    @property
    def Decorative(self):
        return self.com_object.Decorative

    @Decorative.setter
    def Decorative(self, value):
        self.com_object.Decorative = value

    @property
    def decorative(self):
        """Lower case alias for Decorative"""
        return self.Decorative

    @decorative.setter
    def decorative(self, value):
        """Lower case alias for Decorative.setter"""
        self.Decorative = value

    @property
    def Fill(self):
        return FillFormat(self.com_object.Fill)

    @property
    def fill(self):
        """Lower case alias for Fill"""
        return self.Fill

    @property
    def Glow(self):
        return self.com_object.Glow

    @property
    def glow(self):
        """Lower case alias for Glow"""
        return self.Glow

    @property
    def GraphicStyle(self):
        return self.com_object.GraphicStyle

    @GraphicStyle.setter
    def GraphicStyle(self, value):
        self.com_object.GraphicStyle = value

    @property
    def graphicstyle(self):
        """Lower case alias for GraphicStyle"""
        return self.GraphicStyle

    @graphicstyle.setter
    def graphicstyle(self, value):
        """Lower case alias for GraphicStyle.setter"""
        self.GraphicStyle = value

    @property
    def GroupItems(self):
        return GroupShapes(self.com_object.GroupItems)

    @property
    def groupitems(self):
        """Lower case alias for GroupItems"""
        return self.GroupItems

    @property
    def HasChart(self):
        return self.com_object.HasChart

    @property
    def haschart(self):
        """Lower case alias for HasChart"""
        return self.HasChart

    @property
    def hasinkxml(self):
        return self.com_object.hasinkxml

    @property
    def HasSmartArt(self):
        return self.com_object.HasSmartArt

    @property
    def hassmartart(self):
        """Lower case alias for HasSmartArt"""
        return self.HasSmartArt

    @property
    def HasTable(self):
        return self.com_object.HasTable

    @property
    def hastable(self):
        """Lower case alias for HasTable"""
        return self.HasTable

    @property
    def HasTextFrame(self):
        return self.com_object.HasTextFrame

    @property
    def hastextframe(self):
        """Lower case alias for HasTextFrame"""
        return self.HasTextFrame

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
    def HorizontalFlip(self):
        return self.com_object.HorizontalFlip

    @property
    def horizontalflip(self):
        """Lower case alias for HorizontalFlip"""
        return self.HorizontalFlip

    @property
    def Id(self):
        return self.com_object.Id

    @property
    def id(self):
        """Lower case alias for Id"""
        return self.Id

    @property
    def inkxml(self):
        return self.com_object.inkxml

    @property
    def isnarration(self):
        return self.com_object.isnarration

    @isnarration.setter
    def isnarration(self, value):
        self.com_object.isnarration = value

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
    def Line(self):
        return LineFormat(self.com_object.Line)

    @property
    def line(self):
        """Lower case alias for Line"""
        return self.Line

    @property
    def LinkFormat(self):
        return LinkFormat(self.com_object.LinkFormat)

    @property
    def linkformat(self):
        """Lower case alias for LinkFormat"""
        return self.LinkFormat

    @property
    def LockAspectRatio(self):
        return self.com_object.LockAspectRatio

    @LockAspectRatio.setter
    def LockAspectRatio(self, value):
        self.com_object.LockAspectRatio = value

    @property
    def lockaspectratio(self):
        """Lower case alias for LockAspectRatio"""
        return self.LockAspectRatio

    @lockaspectratio.setter
    def lockaspectratio(self, value):
        """Lower case alias for LockAspectRatio.setter"""
        self.LockAspectRatio = value

    @property
    def MediaFormat(self):
        return MediaFormat(self.com_object.MediaFormat)

    @property
    def mediaformat(self):
        """Lower case alias for MediaFormat"""
        return self.MediaFormat

    @property
    def MediaType(self):
        return self.com_object.MediaType

    @property
    def mediatype(self):
        """Lower case alias for MediaType"""
        return self.MediaType

    @property
    def Model3D(self):
        return Model3DFormat(self.com_object.Model3D)

    @property
    def model3d(self):
        """Lower case alias for Model3D"""
        return self.Model3D

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
    def Nodes(self):
        return ShapeNodes(self.com_object.Nodes)

    @property
    def nodes(self):
        """Lower case alias for Nodes"""
        return self.Nodes

    @property
    def OLEFormat(self):
        return OLEFormat(self.com_object.OLEFormat)

    @property
    def oleformat(self):
        """Lower case alias for OLEFormat"""
        return self.OLEFormat

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def ParentGroup(self):
        return Shape(self.com_object.ParentGroup)

    @property
    def parentgroup(self):
        """Lower case alias for ParentGroup"""
        return self.ParentGroup

    @property
    def PictureFormat(self):
        return PictureFormat(self.com_object.PictureFormat)

    @property
    def pictureformat(self):
        """Lower case alias for PictureFormat"""
        return self.PictureFormat

    @property
    def PlaceholderFormat(self):
        return PlaceholderFormat(self.com_object.PlaceholderFormat)

    @property
    def placeholderformat(self):
        """Lower case alias for PlaceholderFormat"""
        return self.PlaceholderFormat

    @property
    def Reflection(self):
        return self.com_object.Reflection

    @property
    def reflection(self):
        """Lower case alias for Reflection"""
        return self.Reflection

    @property
    def Rotation(self):
        return self.com_object.Rotation

    @Rotation.setter
    def Rotation(self, value):
        self.com_object.Rotation = value

    @property
    def rotation(self):
        """Lower case alias for Rotation"""
        return self.Rotation

    @rotation.setter
    def rotation(self, value):
        """Lower case alias for Rotation.setter"""
        self.Rotation = value

    @property
    def Shadow(self):
        return ShadowFormat(self.com_object.Shadow)

    @property
    def shadow(self):
        """Lower case alias for Shadow"""
        return self.Shadow

    @property
    def ShapeStyle(self):
        return self.com_object.ShapeStyle

    @property
    def shapestyle(self):
        """Lower case alias for ShapeStyle"""
        return self.ShapeStyle

    @property
    def SmartArt(self):
        return ShapeRange(self.com_object.SmartArt)

    @property
    def smartart(self):
        """Lower case alias for SmartArt"""
        return self.SmartArt

    @property
    def SoftEdge(self):
        return self.com_object.SoftEdge

    @property
    def softedge(self):
        """Lower case alias for SoftEdge"""
        return self.SoftEdge

    @property
    def Table(self):
        return Table(self.com_object.Table)

    @property
    def table(self):
        """Lower case alias for Table"""
        return self.Table

    @property
    def Tags(self):
        return Tags(self.com_object.Tags)

    @property
    def tags(self):
        """Lower case alias for Tags"""
        return self.Tags

    @property
    def TextEffect(self):
        return TextEffectFormat(self.com_object.TextEffect)

    @property
    def texteffect(self):
        """Lower case alias for TextEffect"""
        return self.TextEffect

    @property
    def TextFrame(self):
        return TextFrame(self.com_object.TextFrame)

    @property
    def textframe(self):
        """Lower case alias for TextFrame"""
        return self.TextFrame

    @property
    def TextFrame2(self):
        return TextFrame2(self.com_object.TextFrame2)

    @property
    def textframe2(self):
        """Lower case alias for TextFrame2"""
        return self.TextFrame2

    @property
    def ThreeD(self):
        return ThreeDFormat(self.com_object.ThreeD)

    @property
    def threed(self):
        """Lower case alias for ThreeD"""
        return self.ThreeD

    @property
    def Title(self):
        return Shape(self.com_object.Title)

    @property
    def title(self):
        """Lower case alias for Title"""
        return self.Title

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
    def Type(self):
        return self.com_object.Type

    @property
    def type(self):
        """Lower case alias for Type"""
        return self.Type

    @property
    def VerticalFlip(self):
        return self.com_object.VerticalFlip

    @property
    def verticalflip(self):
        """Lower case alias for VerticalFlip"""
        return self.VerticalFlip

    @property
    def Vertices(self):
        return self.com_object.Vertices

    @property
    def vertices(self):
        """Lower case alias for Vertices"""
        return self.Vertices

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
    def ZOrderPosition(self):
        return self.com_object.ZOrderPosition

    @property
    def zorderposition(self):
        """Lower case alias for ZOrderPosition"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Count(self):
        return self.com_object.Count

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def HasTitle(self):
        return self.com_object.HasTitle

    @property
    def hastitle(self):
        """Lower case alias for HasTitle"""
        return self.HasTitle

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Placeholders(self):
        return Placeholders(self.com_object.Placeholders)

    @property
    def placeholders(self):
        """Lower case alias for Placeholders"""
        return self.Placeholders

    @property
    def Title(self):
        return Shape(self.com_object.Title)

    @property
    def title(self):
        """Lower case alias for Title"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Background(self):
        return ShapeRange(self.com_object.Background)

    @property
    def background(self):
        """Lower case alias for Background"""
        return self.Background

    @property
    def BackgroundStyle(self):
        return self.com_object.BackgroundStyle

    @BackgroundStyle.setter
    def BackgroundStyle(self, value):
        self.com_object.BackgroundStyle = value

    @property
    def backgroundstyle(self):
        """Lower case alias for BackgroundStyle"""
        return self.BackgroundStyle

    @backgroundstyle.setter
    def backgroundstyle(self, value):
        """Lower case alias for BackgroundStyle.setter"""
        self.BackgroundStyle = value

    @property
    def ColorScheme(self):
        return ColorScheme(self.com_object.ColorScheme)

    @ColorScheme.setter
    def ColorScheme(self, value):
        self.com_object.ColorScheme = value

    @property
    def colorscheme(self):
        """Lower case alias for ColorScheme"""
        return self.ColorScheme

    @colorscheme.setter
    def colorscheme(self, value):
        """Lower case alias for ColorScheme.setter"""
        self.ColorScheme = value

    @property
    def Comments(self):
        return Comments(self.com_object.Comments)

    @property
    def comments(self):
        """Lower case alias for Comments"""
        return self.Comments

    @property
    def CustomerData(self):
        return CustomerData(self.com_object.CustomerData)

    @property
    def customerdata(self):
        """Lower case alias for CustomerData"""
        return self.CustomerData

    @property
    def CustomLayout(self):
        return CustomLayout(self.com_object.CustomLayout)

    @property
    def customlayout(self):
        """Lower case alias for CustomLayout"""
        return self.CustomLayout

    @property
    def Design(self):
        return Design(self.com_object.Design)

    @property
    def design(self):
        """Lower case alias for Design"""
        return self.Design

    @property
    def DisplayMasterShapes(self):
        return self.com_object.DisplayMasterShapes

    @DisplayMasterShapes.setter
    def DisplayMasterShapes(self, value):
        self.com_object.DisplayMasterShapes = value

    @property
    def displaymastershapes(self):
        """Lower case alias for DisplayMasterShapes"""
        return self.DisplayMasterShapes

    @displaymastershapes.setter
    def displaymastershapes(self, value):
        """Lower case alias for DisplayMasterShapes.setter"""
        self.DisplayMasterShapes = value

    @property
    def FollowMasterBackground(self):
        return self.com_object.FollowMasterBackground

    @FollowMasterBackground.setter
    def FollowMasterBackground(self, value):
        self.com_object.FollowMasterBackground = value

    @property
    def followmasterbackground(self):
        """Lower case alias for FollowMasterBackground"""
        return self.FollowMasterBackground

    @followmasterbackground.setter
    def followmasterbackground(self, value):
        """Lower case alias for FollowMasterBackground.setter"""
        self.FollowMasterBackground = value

    @property
    def HasNotesPage(self):
        return self.com_object.HasNotesPage

    @property
    def hasnotespage(self):
        """Lower case alias for HasNotesPage"""
        return self.HasNotesPage

    @property
    def HeadersFooters(self):
        return HeadersFooters(self.com_object.HeadersFooters)

    @property
    def headersfooters(self):
        """Lower case alias for HeadersFooters"""
        return self.HeadersFooters

    @property
    def Hyperlinks(self):
        return Hyperlinks(self.com_object.Hyperlinks)

    @property
    def hyperlinks(self):
        """Lower case alias for Hyperlinks"""
        return self.Hyperlinks

    @property
    def Layout(self):
        return PpSlideLayout(self.com_object.Layout)

    @Layout.setter
    def Layout(self, value):
        self.com_object.Layout = value

    @property
    def layout(self):
        """Lower case alias for Layout"""
        return self.Layout

    @layout.setter
    def layout(self, value):
        """Lower case alias for Layout.setter"""
        self.Layout = value

    @property
    def Master(self):
        return Master(self.com_object.Master)

    @property
    def master(self):
        """Lower case alias for Master"""
        return self.Master

    @property
    def Name(self):
        return self.com_object.Name

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def NotesPage(self):
        return SlideRange(self.com_object.NotesPage)

    @property
    def notespage(self):
        """Lower case alias for NotesPage"""
        return self.NotesPage

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PrintSteps(self):
        return self.com_object.PrintSteps

    @property
    def printsteps(self):
        """Lower case alias for PrintSteps"""
        return self.PrintSteps

    @property
    def sectionIndex(self):
        return Slide(self.com_object.sectionIndex)

    @property
    def sectionindex(self):
        """Lower case alias for sectionIndex"""
        return self.sectionIndex

    @property
    def Shapes(self):
        return Shapes(self.com_object.Shapes)

    @property
    def shapes(self):
        """Lower case alias for Shapes"""
        return self.Shapes

    @property
    def SlideID(self):
        return self.com_object.SlideID

    @property
    def slideid(self):
        """Lower case alias for SlideID"""
        return self.SlideID

    @property
    def SlideIndex(self):
        return Slides(self.com_object.SlideIndex)

    @property
    def slideindex(self):
        """Lower case alias for SlideIndex"""
        return self.SlideIndex

    @property
    def SlideNumber(self):
        return self.com_object.SlideNumber

    @property
    def slidenumber(self):
        """Lower case alias for SlideNumber"""
        return self.SlideNumber

    @property
    def SlideShowTransition(self):
        return SlideShowTransition(self.com_object.SlideShowTransition)

    @property
    def slideshowtransition(self):
        """Lower case alias for SlideShowTransition"""
        return self.SlideShowTransition

    @property
    def Tags(self):
        return Tags(self.com_object.Tags)

    @property
    def tags(self):
        """Lower case alias for Tags"""
        return self.Tags

    @property
    def ThemeColorScheme(self):
        return self.com_object.ThemeColorScheme

    @property
    def themecolorscheme(self):
        """Lower case alias for ThemeColorScheme"""
        return self.ThemeColorScheme

    @property
    def TimeLine(self):
        return TimeLine(self.com_object.TimeLine)

    @property
    def timeline(self):
        """Lower case alias for TimeLine"""
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


class slidenavigation:

    def __init__(self, slidenavigation=None):
        self.com_object= slidenavigation

    @property
    def application(self):
        return self.com_object.application

    @property
    def parent(self):
        return self.com_object.parent

    @property
    def visible(self):
        return self.com_object.visible

    @visible.setter
    def visible(self, value):
        self.com_object.visible = value


class SlideRange:

    def __init__(self, sliderange=None):
        self.com_object= sliderange

    def __call__(self, item):
        return SlideRange(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Background(self):
        return ShapeRange(self.com_object.Background)

    @property
    def background(self):
        """Lower case alias for Background"""
        return self.Background

    @property
    def BackgroundStyle(self):
        return self.com_object.BackgroundStyle

    @BackgroundStyle.setter
    def BackgroundStyle(self, value):
        self.com_object.BackgroundStyle = value

    @property
    def backgroundstyle(self):
        """Lower case alias for BackgroundStyle"""
        return self.BackgroundStyle

    @backgroundstyle.setter
    def backgroundstyle(self, value):
        """Lower case alias for BackgroundStyle.setter"""
        self.BackgroundStyle = value

    @property
    def ColorScheme(self):
        return ColorScheme(self.com_object.ColorScheme)

    @ColorScheme.setter
    def ColorScheme(self, value):
        self.com_object.ColorScheme = value

    @property
    def colorscheme(self):
        """Lower case alias for ColorScheme"""
        return self.ColorScheme

    @colorscheme.setter
    def colorscheme(self, value):
        """Lower case alias for ColorScheme.setter"""
        self.ColorScheme = value

    @property
    def Comments(self):
        return Comments(self.com_object.Comments)

    @property
    def comments(self):
        """Lower case alias for Comments"""
        return self.Comments

    @property
    def Count(self):
        return self.com_object.Count

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def CustomerData(self):
        return CustomerData(self.com_object.CustomerData)

    @property
    def customerdata(self):
        """Lower case alias for CustomerData"""
        return self.CustomerData

    @property
    def CustomLayout(self):
        return CustomLayout(self.com_object.CustomLayout)

    @property
    def customlayout(self):
        """Lower case alias for CustomLayout"""
        return self.CustomLayout

    @property
    def Design(self):
        return Design(self.com_object.Design)

    @property
    def design(self):
        """Lower case alias for Design"""
        return self.Design

    @property
    def DisplayMasterShapes(self):
        return self.com_object.DisplayMasterShapes

    @DisplayMasterShapes.setter
    def DisplayMasterShapes(self, value):
        self.com_object.DisplayMasterShapes = value

    @property
    def displaymastershapes(self):
        """Lower case alias for DisplayMasterShapes"""
        return self.DisplayMasterShapes

    @displaymastershapes.setter
    def displaymastershapes(self, value):
        """Lower case alias for DisplayMasterShapes.setter"""
        self.DisplayMasterShapes = value

    @property
    def FollowMasterBackground(self):
        return self.com_object.FollowMasterBackground

    @FollowMasterBackground.setter
    def FollowMasterBackground(self, value):
        self.com_object.FollowMasterBackground = value

    @property
    def followmasterbackground(self):
        """Lower case alias for FollowMasterBackground"""
        return self.FollowMasterBackground

    @followmasterbackground.setter
    def followmasterbackground(self, value):
        """Lower case alias for FollowMasterBackground.setter"""
        self.FollowMasterBackground = value

    @property
    def HasNotesPage(self):
        return self.com_object.HasNotesPage

    @property
    def hasnotespage(self):
        """Lower case alias for HasNotesPage"""
        return self.HasNotesPage

    @property
    def HeadersFooters(self):
        return HeadersFooters(self.com_object.HeadersFooters)

    @property
    def headersfooters(self):
        """Lower case alias for HeadersFooters"""
        return self.HeadersFooters

    @property
    def Hyperlinks(self):
        return Hyperlinks(self.com_object.Hyperlinks)

    @property
    def hyperlinks(self):
        """Lower case alias for Hyperlinks"""
        return self.Hyperlinks

    @property
    def Layout(self):
        return PpSlideLayout(self.com_object.Layout)

    @Layout.setter
    def Layout(self, value):
        self.com_object.Layout = value

    @property
    def layout(self):
        """Lower case alias for Layout"""
        return self.Layout

    @layout.setter
    def layout(self, value):
        """Lower case alias for Layout.setter"""
        self.Layout = value

    @property
    def Master(self):
        return Master(self.com_object.Master)

    @property
    def master(self):
        """Lower case alias for Master"""
        return self.Master

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
    def NotesPage(self):
        return SlideRange(self.com_object.NotesPage)

    @property
    def notespage(self):
        """Lower case alias for NotesPage"""
        return self.NotesPage

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PrintSteps(self):
        return self.com_object.PrintSteps

    @property
    def printsteps(self):
        """Lower case alias for PrintSteps"""
        return self.PrintSteps

    @property
    def sectionIndex(self):
        return SlideRange(self.com_object.sectionIndex)

    @property
    def sectionindex(self):
        """Lower case alias for sectionIndex"""
        return self.sectionIndex

    @property
    def Shapes(self):
        return Shapes(self.com_object.Shapes)

    @property
    def shapes(self):
        """Lower case alias for Shapes"""
        return self.Shapes

    @property
    def SlideID(self):
        return self.com_object.SlideID

    @property
    def slideid(self):
        """Lower case alias for SlideID"""
        return self.SlideID

    @property
    def SlideIndex(self):
        return Slides(self.com_object.SlideIndex)

    @property
    def slideindex(self):
        """Lower case alias for SlideIndex"""
        return self.SlideIndex

    @property
    def SlideNumber(self):
        return self.com_object.SlideNumber

    @property
    def slidenumber(self):
        """Lower case alias for SlideNumber"""
        return self.SlideNumber

    @property
    def SlideShowTransition(self):
        return SlideShowTransition(self.com_object.SlideShowTransition)

    @property
    def slideshowtransition(self):
        """Lower case alias for SlideShowTransition"""
        return self.SlideShowTransition

    @property
    def Tags(self):
        return Tags(self.com_object.Tags)

    @property
    def tags(self):
        """Lower case alias for Tags"""
        return self.Tags

    @property
    def ThemeColorScheme(self):
        return self.com_object.ThemeColorScheme

    @property
    def themecolorscheme(self):
        """Lower case alias for ThemeColorScheme"""
        return self.ThemeColorScheme

    @property
    def TimeLine(self):
        return TimeLine(self.com_object.TimeLine)

    @property
    def timeline(self):
        """Lower case alias for TimeLine"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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

    def Add(self, Index=None, Layout=None):
        arguments = com_arguments([unwrap(a) for a in [Index, Layout]])
        return Slide(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Index=None, Layout=None):
        arguments = [Index, Layout]
        return self.Add(*arguments)

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

    @property
    def advancemode(self):
        """Lower case alias for AdvanceMode"""
        return self.AdvanceMode

    @advancemode.setter
    def advancemode(self, value):
        """Lower case alias for AdvanceMode.setter"""
        self.AdvanceMode = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def EndingSlide(self):
        return self.com_object.EndingSlide

    @EndingSlide.setter
    def EndingSlide(self, value):
        self.com_object.EndingSlide = value

    @property
    def endingslide(self):
        """Lower case alias for EndingSlide"""
        return self.EndingSlide

    @endingslide.setter
    def endingslide(self, value):
        """Lower case alias for EndingSlide.setter"""
        self.EndingSlide = value

    @property
    def LoopUntilStopped(self):
        return self.com_object.LoopUntilStopped

    @LoopUntilStopped.setter
    def LoopUntilStopped(self, value):
        self.com_object.LoopUntilStopped = value

    @property
    def loopuntilstopped(self):
        """Lower case alias for LoopUntilStopped"""
        return self.LoopUntilStopped

    @loopuntilstopped.setter
    def loopuntilstopped(self, value):
        """Lower case alias for LoopUntilStopped.setter"""
        self.LoopUntilStopped = value

    @property
    def NamedSlideShows(self):
        return NamedSlideShows(self.com_object.NamedSlideShows)

    @property
    def namedslideshows(self):
        """Lower case alias for NamedSlideShows"""
        return self.NamedSlideShows

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PointerColor(self):
        return ColorFormat(self.com_object.PointerColor)

    @property
    def pointercolor(self):
        """Lower case alias for PointerColor"""
        return self.PointerColor

    @property
    def RangeType(self):
        return self.com_object.RangeType

    @RangeType.setter
    def RangeType(self, value):
        self.com_object.RangeType = value

    @property
    def rangetype(self):
        """Lower case alias for RangeType"""
        return self.RangeType

    @rangetype.setter
    def rangetype(self, value):
        """Lower case alias for RangeType.setter"""
        self.RangeType = value

    @property
    def ShowMediaControls(self):
        return self.com_object.ShowMediaControls

    @ShowMediaControls.setter
    def ShowMediaControls(self, value):
        self.com_object.ShowMediaControls = value

    @property
    def showmediacontrols(self):
        """Lower case alias for ShowMediaControls"""
        return self.ShowMediaControls

    @showmediacontrols.setter
    def showmediacontrols(self, value):
        """Lower case alias for ShowMediaControls.setter"""
        self.ShowMediaControls = value

    @property
    def ShowPresenterView(self):
        return SlideShowSettings(self.com_object.ShowPresenterView)

    @ShowPresenterView.setter
    def ShowPresenterView(self, value):
        self.com_object.ShowPresenterView = value

    @property
    def showpresenterview(self):
        """Lower case alias for ShowPresenterView"""
        return self.ShowPresenterView

    @showpresenterview.setter
    def showpresenterview(self, value):
        """Lower case alias for ShowPresenterView.setter"""
        self.ShowPresenterView = value

    @property
    def ShowScrollbar(self):
        return self.com_object.ShowScrollbar

    @ShowScrollbar.setter
    def ShowScrollbar(self, value):
        self.com_object.ShowScrollbar = value

    @property
    def showscrollbar(self):
        """Lower case alias for ShowScrollbar"""
        return self.ShowScrollbar

    @showscrollbar.setter
    def showscrollbar(self, value):
        """Lower case alias for ShowScrollbar.setter"""
        self.ShowScrollbar = value

    @property
    def ShowType(self):
        return self.com_object.ShowType

    @ShowType.setter
    def ShowType(self, value):
        self.com_object.ShowType = value

    @property
    def showtype(self):
        """Lower case alias for ShowType"""
        return self.ShowType

    @showtype.setter
    def showtype(self, value):
        """Lower case alias for ShowType.setter"""
        self.ShowType = value

    @property
    def ShowWithAnimation(self):
        return self.com_object.ShowWithAnimation

    @ShowWithAnimation.setter
    def ShowWithAnimation(self, value):
        self.com_object.ShowWithAnimation = value

    @property
    def showwithanimation(self):
        """Lower case alias for ShowWithAnimation"""
        return self.ShowWithAnimation

    @showwithanimation.setter
    def showwithanimation(self, value):
        """Lower case alias for ShowWithAnimation.setter"""
        self.ShowWithAnimation = value

    @property
    def ShowWithNarration(self):
        return self.com_object.ShowWithNarration

    @ShowWithNarration.setter
    def ShowWithNarration(self, value):
        self.com_object.ShowWithNarration = value

    @property
    def showwithnarration(self):
        """Lower case alias for ShowWithNarration"""
        return self.ShowWithNarration

    @showwithnarration.setter
    def showwithnarration(self, value):
        """Lower case alias for ShowWithNarration.setter"""
        self.ShowWithNarration = value

    @property
    def SlideShowName(self):
        return self.com_object.SlideShowName

    @SlideShowName.setter
    def SlideShowName(self, value):
        self.com_object.SlideShowName = value

    @property
    def slideshowname(self):
        """Lower case alias for SlideShowName"""
        return self.SlideShowName

    @slideshowname.setter
    def slideshowname(self, value):
        """Lower case alias for SlideShowName.setter"""
        self.SlideShowName = value

    @property
    def StartingSlide(self):
        return self.com_object.StartingSlide

    @StartingSlide.setter
    def StartingSlide(self, value):
        self.com_object.StartingSlide = value

    @property
    def startingslide(self):
        """Lower case alias for StartingSlide"""
        return self.StartingSlide

    @startingslide.setter
    def startingslide(self, value):
        """Lower case alias for StartingSlide.setter"""
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

    @property
    def advanceonclick(self):
        """Lower case alias for AdvanceOnClick"""
        return self.AdvanceOnClick

    @advanceonclick.setter
    def advanceonclick(self, value):
        """Lower case alias for AdvanceOnClick.setter"""
        self.AdvanceOnClick = value

    @property
    def AdvanceOnTime(self):
        return self.com_object.AdvanceOnTime

    @AdvanceOnTime.setter
    def AdvanceOnTime(self, value):
        self.com_object.AdvanceOnTime = value

    @property
    def advanceontime(self):
        """Lower case alias for AdvanceOnTime"""
        return self.AdvanceOnTime

    @advanceontime.setter
    def advanceontime(self, value):
        """Lower case alias for AdvanceOnTime.setter"""
        self.AdvanceOnTime = value

    @property
    def AdvanceTime(self):
        return self.com_object.AdvanceTime

    @AdvanceTime.setter
    def AdvanceTime(self, value):
        self.com_object.AdvanceTime = value

    @property
    def advancetime(self):
        """Lower case alias for AdvanceTime"""
        return self.AdvanceTime

    @advancetime.setter
    def advancetime(self, value):
        """Lower case alias for AdvanceTime.setter"""
        self.AdvanceTime = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Duration(self):
        return self.com_object.Duration

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
    def EntryEffect(self):
        return self.com_object.EntryEffect

    @EntryEffect.setter
    def EntryEffect(self, value):
        self.com_object.EntryEffect = value

    @property
    def entryeffect(self):
        """Lower case alias for EntryEffect"""
        return self.EntryEffect

    @entryeffect.setter
    def entryeffect(self, value):
        """Lower case alias for EntryEffect.setter"""
        self.EntryEffect = value

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
    def LoopSoundUntilNext(self):
        return self.com_object.LoopSoundUntilNext

    @LoopSoundUntilNext.setter
    def LoopSoundUntilNext(self, value):
        self.com_object.LoopSoundUntilNext = value

    @property
    def loopsounduntilnext(self):
        """Lower case alias for LoopSoundUntilNext"""
        return self.LoopSoundUntilNext

    @loopsounduntilnext.setter
    def loopsounduntilnext(self, value):
        """Lower case alias for LoopSoundUntilNext.setter"""
        self.LoopSoundUntilNext = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def SoundEffect(self):
        return SoundEffect(self.com_object.SoundEffect)

    @property
    def soundeffect(self):
        """Lower case alias for SoundEffect"""
        return self.SoundEffect

    @property
    def Speed(self):
        return self.com_object.Speed

    @Speed.setter
    def Speed(self, value):
        self.com_object.Speed = value

    @property
    def speed(self):
        """Lower case alias for Speed"""
        return self.Speed

    @speed.setter
    def speed(self, value):
        """Lower case alias for Speed.setter"""
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

    @property
    def acceleratorsenabled(self):
        """Lower case alias for AcceleratorsEnabled"""
        return self.AcceleratorsEnabled

    @acceleratorsenabled.setter
    def acceleratorsenabled(self, value):
        """Lower case alias for AcceleratorsEnabled.setter"""
        self.AcceleratorsEnabled = value

    @property
    def AdvanceMode(self):
        return self.com_object.AdvanceMode

    @property
    def advancemode(self):
        """Lower case alias for AdvanceMode"""
        return self.AdvanceMode

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def CurrentShowPosition(self):
        return self.com_object.CurrentShowPosition

    @property
    def currentshowposition(self):
        """Lower case alias for CurrentShowPosition"""
        return self.CurrentShowPosition

    @property
    def IsNamedShow(self):
        return self.com_object.IsNamedShow

    @property
    def isnamedshow(self):
        """Lower case alias for IsNamedShow"""
        return self.IsNamedShow

    @property
    def laserpointerenabled(self):
        return self.com_object.laserpointerenabled

    @laserpointerenabled.setter
    def laserpointerenabled(self, value):
        self.com_object.laserpointerenabled = value

    @property
    def LastSlideViewed(self):
        return Slide(self.com_object.LastSlideViewed)

    @property
    def lastslideviewed(self):
        """Lower case alias for LastSlideViewed"""
        return self.LastSlideViewed

    @property
    def MediaControlsHeight(self):
        return self.com_object.MediaControlsHeight

    @property
    def mediacontrolsheight(self):
        """Lower case alias for MediaControlsHeight"""
        return self.MediaControlsHeight

    @property
    def MediaControlsLeft(self):
        return Slide(self.com_object.MediaControlsLeft)

    @property
    def mediacontrolsleft(self):
        """Lower case alias for MediaControlsLeft"""
        return self.MediaControlsLeft

    @property
    def MediaControlsTop(self):
        return Slide(self.com_object.MediaControlsTop)

    @property
    def mediacontrolstop(self):
        """Lower case alias for MediaControlsTop"""
        return self.MediaControlsTop

    @property
    def MediaControlsVisible(self):
        return self.com_object.MediaControlsVisible

    @property
    def mediacontrolsvisible(self):
        """Lower case alias for MediaControlsVisible"""
        return self.MediaControlsVisible

    @property
    def MediaControlsWidth(self):
        return self.com_object.MediaControlsWidth

    @property
    def mediacontrolswidth(self):
        """Lower case alias for MediaControlsWidth"""
        return self.MediaControlsWidth

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PointerColor(self):
        return ColorFormat(self.com_object.PointerColor)

    @property
    def pointercolor(self):
        """Lower case alias for PointerColor"""
        return self.PointerColor

    @property
    def PointerType(self):
        return self.com_object.PointerType

    @PointerType.setter
    def PointerType(self, value):
        self.com_object.PointerType = value

    @property
    def pointertype(self):
        """Lower case alias for PointerType"""
        return self.PointerType

    @pointertype.setter
    def pointertype(self, value):
        """Lower case alias for PointerType.setter"""
        self.PointerType = value

    @property
    def PresentationElapsedTime(self):
        return self.com_object.PresentationElapsedTime

    @property
    def presentationelapsedtime(self):
        """Lower case alias for PresentationElapsedTime"""
        return self.PresentationElapsedTime

    @property
    def Slide(self):
        return Slide(self.com_object.Slide)

    @property
    def slide(self):
        """Lower case alias for Slide"""
        return self.Slide

    @property
    def SlideElapsedTime(self):
        return self.com_object.SlideElapsedTime

    @SlideElapsedTime.setter
    def SlideElapsedTime(self, value):
        self.com_object.SlideElapsedTime = value

    @property
    def slideelapsedtime(self):
        """Lower case alias for SlideElapsedTime"""
        return self.SlideElapsedTime

    @slideelapsedtime.setter
    def slideelapsedtime(self, value):
        """Lower case alias for SlideElapsedTime.setter"""
        self.SlideElapsedTime = value

    @property
    def SlideShowName(self):
        return self.com_object.SlideShowName

    @property
    def slideshowname(self):
        """Lower case alias for SlideShowName"""
        return self.SlideShowName

    @property
    def State(self):
        return self.com_object.State

    @State.setter
    def State(self, value):
        self.com_object.State = value

    @property
    def state(self):
        """Lower case alias for State"""
        return self.State

    @state.setter
    def state(self, value):
        """Lower case alias for State.setter"""
        self.State = value

    @property
    def Zoom(self):
        return self.com_object.Zoom

    @property
    def zoom(self):
        """Lower case alias for Zoom"""
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

    @property
    def active(self):
        """Lower case alias for Active"""
        return self.Active

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def IsFullScreen(self):
        return self.com_object.IsFullScreen

    @property
    def isfullscreen(self):
        """Lower case alias for IsFullScreen"""
        return self.IsFullScreen

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
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Presentation(self):
        return Presentation(self.com_object.Presentation)

    @property
    def presentation(self):
        """Lower case alias for Presentation"""
        return self.Presentation

    @property
    def slidenavigation(self):
        return SlideNavigation(self.com_object.slidenavigation)

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
    def View(self):
        return SlideShowView(self.com_object.View)

    @property
    def view(self):
        """Lower case alias for View"""
        return self.View

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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

    @property
    def alternativetext(self):
        """Lower case alias for AlternativeText"""
        return self.AlternativeText

    @alternativetext.setter
    def alternativetext(self, value):
        """Lower case alias for AlternativeText.setter"""
        self.AlternativeText = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Background(self):
        return TableBackground(self.com_object.Background)

    @property
    def background(self):
        """Lower case alias for Background"""
        return self.Background

    @property
    def Columns(self):
        return Columns(self.com_object.Columns)

    @property
    def columns(self):
        """Lower case alias for Columns"""
        return self.Columns

    @property
    def FirstCol(self):
        return self.com_object.FirstCol

    @FirstCol.setter
    def FirstCol(self, value):
        self.com_object.FirstCol = value

    @property
    def firstcol(self):
        """Lower case alias for FirstCol"""
        return self.FirstCol

    @firstcol.setter
    def firstcol(self, value):
        """Lower case alias for FirstCol.setter"""
        self.FirstCol = value

    @property
    def FirstRow(self):
        return self.com_object.FirstRow

    @FirstRow.setter
    def FirstRow(self, value):
        self.com_object.FirstRow = value

    @property
    def firstrow(self):
        """Lower case alias for FirstRow"""
        return self.FirstRow

    @firstrow.setter
    def firstrow(self, value):
        """Lower case alias for FirstRow.setter"""
        self.FirstRow = value

    @property
    def HorizBanding(self):
        return self.com_object.HorizBanding

    @HorizBanding.setter
    def HorizBanding(self, value):
        self.com_object.HorizBanding = value

    @property
    def horizbanding(self):
        """Lower case alias for HorizBanding"""
        return self.HorizBanding

    @horizbanding.setter
    def horizbanding(self, value):
        """Lower case alias for HorizBanding.setter"""
        self.HorizBanding = value

    @property
    def LastCol(self):
        return self.com_object.LastCol

    @LastCol.setter
    def LastCol(self, value):
        self.com_object.LastCol = value

    @property
    def lastcol(self):
        """Lower case alias for LastCol"""
        return self.LastCol

    @lastcol.setter
    def lastcol(self, value):
        """Lower case alias for LastCol.setter"""
        self.LastCol = value

    @property
    def LastRow(self):
        return self.com_object.LastRow

    @LastRow.setter
    def LastRow(self, value):
        self.com_object.LastRow = value

    @property
    def lastrow(self):
        """Lower case alias for LastRow"""
        return self.LastRow

    @lastrow.setter
    def lastrow(self, value):
        """Lower case alias for LastRow.setter"""
        self.LastRow = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Rows(self):
        return Rows(self.com_object.Rows)

    @property
    def rows(self):
        """Lower case alias for Rows"""
        return self.Rows

    @property
    def Style(self):
        return TableStyle(self.com_object.Style)

    @property
    def style(self):
        """Lower case alias for Style"""
        return self.Style

    @property
    def TableDirection(self):
        return self.com_object.TableDirection

    @TableDirection.setter
    def TableDirection(self, value):
        self.com_object.TableDirection = value

    @property
    def tabledirection(self):
        """Lower case alias for TableDirection"""
        return self.TableDirection

    @tabledirection.setter
    def tabledirection(self, value):
        """Lower case alias for TableDirection.setter"""
        self.TableDirection = value

    @property
    def Title(self):
        return Table(self.com_object.Title)

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
    def VertBanding(self):
        return self.com_object.VertBanding

    @VertBanding.setter
    def VertBanding(self, value):
        self.com_object.VertBanding = value

    @property
    def vertbanding(self):
        """Lower case alias for VertBanding"""
        return self.VertBanding

    @vertbanding.setter
    def vertbanding(self, value):
        """Lower case alias for VertBanding.setter"""
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

    @property
    def fill(self):
        """Lower case alias for Fill"""
        return self.Fill

    @property
    def Picture(self):
        return PictureFormat(self.com_object.Picture)

    @property
    def picture(self):
        """Lower case alias for Picture"""
        return self.Picture

    @property
    def Reflection(self):
        return self.com_object.Reflection

    @property
    def reflection(self):
        """Lower case alias for Reflection"""
        return self.Reflection

    @property
    def Shadow(self):
        return ShadowFormat(self.com_object.Shadow)

    @property
    def shadow(self):
        """Lower case alias for Shadow"""
        return self.Shadow


class TableStyle:

    def __init__(self, tablestyle=None):
        self.com_object= tablestyle

    @property
    def Id(self):
        return self.com_object.Id

    @property
    def id(self):
        """Lower case alias for Id"""
        return self.Id

    @property
    def Name(self):
        return self.com_object.Name

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name


class TabStop:

    def __init__(self, tabstop=None):
        self.com_object= tabstop

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Count(self):
        return self.com_object.Count

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def DefaultSpacing(self):
        return self.com_object.DefaultSpacing

    @DefaultSpacing.setter
    def DefaultSpacing(self, value):
        self.com_object.DefaultSpacing = value

    @property
    def defaultspacing(self):
        """Lower case alias for DefaultSpacing"""
        return self.DefaultSpacing

    @defaultspacing.setter
    def defaultspacing(self, value):
        """Lower case alias for DefaultSpacing.setter"""
        self.DefaultSpacing = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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

    @property
    def alignment(self):
        """Lower case alias for Alignment"""
        return self.Alignment

    @alignment.setter
    def alignment(self, value):
        """Lower case alias for Alignment.setter"""
        self.Alignment = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def FontBold(self):
        return self.com_object.FontBold

    @FontBold.setter
    def FontBold(self, value):
        self.com_object.FontBold = value

    @property
    def fontbold(self):
        """Lower case alias for FontBold"""
        return self.FontBold

    @fontbold.setter
    def fontbold(self, value):
        """Lower case alias for FontBold.setter"""
        self.FontBold = value

    @property
    def FontItalic(self):
        return self.com_object.FontItalic

    @FontItalic.setter
    def FontItalic(self, value):
        self.com_object.FontItalic = value

    @property
    def fontitalic(self):
        """Lower case alias for FontItalic"""
        return self.FontItalic

    @fontitalic.setter
    def fontitalic(self, value):
        """Lower case alias for FontItalic.setter"""
        self.FontItalic = value

    @property
    def FontName(self):
        return self.com_object.FontName

    @FontName.setter
    def FontName(self, value):
        self.com_object.FontName = value

    @property
    def fontname(self):
        """Lower case alias for FontName"""
        return self.FontName

    @fontname.setter
    def fontname(self, value):
        """Lower case alias for FontName.setter"""
        self.FontName = value

    @property
    def FontSize(self):
        return self.com_object.FontSize

    @FontSize.setter
    def FontSize(self, value):
        self.com_object.FontSize = value

    @property
    def fontsize(self):
        """Lower case alias for FontSize"""
        return self.FontSize

    @fontsize.setter
    def fontsize(self, value):
        """Lower case alias for FontSize.setter"""
        self.FontSize = value

    @property
    def KernedPairs(self):
        return self.com_object.KernedPairs

    @KernedPairs.setter
    def KernedPairs(self, value):
        self.com_object.KernedPairs = value

    @property
    def kernedpairs(self):
        """Lower case alias for KernedPairs"""
        return self.KernedPairs

    @kernedpairs.setter
    def kernedpairs(self, value):
        """Lower case alias for KernedPairs.setter"""
        self.KernedPairs = value

    @property
    def NormalizedHeight(self):
        return self.com_object.NormalizedHeight

    @NormalizedHeight.setter
    def NormalizedHeight(self, value):
        self.com_object.NormalizedHeight = value

    @property
    def normalizedheight(self):
        """Lower case alias for NormalizedHeight"""
        return self.NormalizedHeight

    @normalizedheight.setter
    def normalizedheight(self, value):
        """Lower case alias for NormalizedHeight.setter"""
        self.NormalizedHeight = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PresetShape(self):
        return self.com_object.PresetShape

    @PresetShape.setter
    def PresetShape(self, value):
        self.com_object.PresetShape = value

    @property
    def presetshape(self):
        """Lower case alias for PresetShape"""
        return self.PresetShape

    @presetshape.setter
    def presetshape(self, value):
        """Lower case alias for PresetShape.setter"""
        self.PresetShape = value

    @property
    def PresetTextEffect(self):
        return self.com_object.PresetTextEffect

    @PresetTextEffect.setter
    def PresetTextEffect(self, value):
        self.com_object.PresetTextEffect = value

    @property
    def presettexteffect(self):
        """Lower case alias for PresetTextEffect"""
        return self.PresetTextEffect

    @presettexteffect.setter
    def presettexteffect(self, value):
        """Lower case alias for PresetTextEffect.setter"""
        self.PresetTextEffect = value

    @property
    def RotatedChars(self):
        return self.com_object.RotatedChars

    @RotatedChars.setter
    def RotatedChars(self, value):
        self.com_object.RotatedChars = value

    @property
    def rotatedchars(self):
        """Lower case alias for RotatedChars"""
        return self.RotatedChars

    @rotatedchars.setter
    def rotatedchars(self, value):
        """Lower case alias for RotatedChars.setter"""
        self.RotatedChars = value

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
    def Tracking(self):
        return self.com_object.Tracking

    @Tracking.setter
    def Tracking(self, value):
        self.com_object.Tracking = value

    @property
    def tracking(self):
        """Lower case alias for Tracking"""
        return self.Tracking

    @tracking.setter
    def tracking(self, value):
        """Lower case alias for Tracking.setter"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def HasText(self):
        return self.com_object.HasText

    @property
    def hastext(self):
        """Lower case alias for HasText"""
        return self.HasText

    @property
    def HorizontalAnchor(self):
        return self.com_object.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.com_object.HorizontalAnchor = value

    @property
    def horizontalanchor(self):
        """Lower case alias for HorizontalAnchor"""
        return self.HorizontalAnchor

    @horizontalanchor.setter
    def horizontalanchor(self, value):
        """Lower case alias for HorizontalAnchor.setter"""
        self.HorizontalAnchor = value

    @property
    def MarginBottom(self):
        return self.com_object.MarginBottom

    @MarginBottom.setter
    def MarginBottom(self, value):
        self.com_object.MarginBottom = value

    @property
    def marginbottom(self):
        """Lower case alias for MarginBottom"""
        return self.MarginBottom

    @marginbottom.setter
    def marginbottom(self, value):
        """Lower case alias for MarginBottom.setter"""
        self.MarginBottom = value

    @property
    def MarginLeft(self):
        return self.com_object.MarginLeft

    @MarginLeft.setter
    def MarginLeft(self, value):
        self.com_object.MarginLeft = value

    @property
    def marginleft(self):
        """Lower case alias for MarginLeft"""
        return self.MarginLeft

    @marginleft.setter
    def marginleft(self, value):
        """Lower case alias for MarginLeft.setter"""
        self.MarginLeft = value

    @property
    def MarginRight(self):
        return self.com_object.MarginRight

    @MarginRight.setter
    def MarginRight(self, value):
        self.com_object.MarginRight = value

    @property
    def marginright(self):
        """Lower case alias for MarginRight"""
        return self.MarginRight

    @marginright.setter
    def marginright(self, value):
        """Lower case alias for MarginRight.setter"""
        self.MarginRight = value

    @property
    def MarginTop(self):
        return self.com_object.MarginTop

    @MarginTop.setter
    def MarginTop(self, value):
        self.com_object.MarginTop = value

    @property
    def margintop(self):
        """Lower case alias for MarginTop"""
        return self.MarginTop

    @margintop.setter
    def margintop(self, value):
        """Lower case alias for MarginTop.setter"""
        self.MarginTop = value

    @property
    def Orientation(self):
        return self.com_object.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.com_object.Orientation = value

    @property
    def orientation(self):
        """Lower case alias for Orientation"""
        return self.Orientation

    @orientation.setter
    def orientation(self, value):
        """Lower case alias for Orientation.setter"""
        self.Orientation = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Ruler(self):
        return Ruler(self.com_object.Ruler)

    @property
    def ruler(self):
        """Lower case alias for Ruler"""
        return self.Ruler

    @property
    def TextRange(self):
        return TextRange(self.com_object.TextRange)

    @property
    def textrange(self):
        """Lower case alias for TextRange"""
        return self.TextRange

    @property
    def VerticalAnchor(self):
        return self.com_object.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.com_object.VerticalAnchor = value

    @property
    def verticalanchor(self):
        """Lower case alias for VerticalAnchor"""
        return self.VerticalAnchor

    @verticalanchor.setter
    def verticalanchor(self, value):
        """Lower case alias for VerticalAnchor.setter"""
        self.VerticalAnchor = value

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def Column(self):
        return Column(self.com_object.Column)

    @property
    def column(self):
        """Lower case alias for Column"""
        return self.Column

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def HasText(self):
        return self.com_object.HasText

    @property
    def hastext(self):
        """Lower case alias for HasText"""
        return self.HasText

    @property
    def HorizontalAnchor(self):
        return self.com_object.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.com_object.HorizontalAnchor = value

    @property
    def horizontalanchor(self):
        """Lower case alias for HorizontalAnchor"""
        return self.HorizontalAnchor

    @horizontalanchor.setter
    def horizontalanchor(self, value):
        """Lower case alias for HorizontalAnchor.setter"""
        self.HorizontalAnchor = value

    @property
    def MarginBottom(self):
        return self.com_object.MarginBottom

    @MarginBottom.setter
    def MarginBottom(self, value):
        self.com_object.MarginBottom = value

    @property
    def marginbottom(self):
        """Lower case alias for MarginBottom"""
        return self.MarginBottom

    @marginbottom.setter
    def marginbottom(self, value):
        """Lower case alias for MarginBottom.setter"""
        self.MarginBottom = value

    @property
    def MarginLeft(self):
        return self.com_object.MarginLeft

    @MarginLeft.setter
    def MarginLeft(self, value):
        self.com_object.MarginLeft = value

    @property
    def marginleft(self):
        """Lower case alias for MarginLeft"""
        return self.MarginLeft

    @marginleft.setter
    def marginleft(self, value):
        """Lower case alias for MarginLeft.setter"""
        self.MarginLeft = value

    @property
    def MarginRight(self):
        return self.com_object.MarginRight

    @MarginRight.setter
    def MarginRight(self, value):
        self.com_object.MarginRight = value

    @property
    def marginright(self):
        """Lower case alias for MarginRight"""
        return self.MarginRight

    @marginright.setter
    def marginright(self, value):
        """Lower case alias for MarginRight.setter"""
        self.MarginRight = value

    @property
    def MarginTop(self):
        return self.com_object.MarginTop

    @MarginTop.setter
    def MarginTop(self, value):
        self.com_object.MarginTop = value

    @property
    def margintop(self):
        """Lower case alias for MarginTop"""
        return self.MarginTop

    @margintop.setter
    def margintop(self, value):
        """Lower case alias for MarginTop.setter"""
        self.MarginTop = value

    @property
    def NoTextRotation(self):
        return self.com_object.NoTextRotation

    @NoTextRotation.setter
    def NoTextRotation(self, value):
        self.com_object.NoTextRotation = value

    @property
    def notextrotation(self):
        """Lower case alias for NoTextRotation"""
        return self.NoTextRotation

    @notextrotation.setter
    def notextrotation(self, value):
        """Lower case alias for NoTextRotation.setter"""
        self.NoTextRotation = value

    @property
    def Orientation(self):
        return self.com_object.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.com_object.Orientation = value

    @property
    def orientation(self):
        """Lower case alias for Orientation"""
        return self.Orientation

    @orientation.setter
    def orientation(self, value):
        """Lower case alias for Orientation.setter"""
        self.Orientation = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PathFormat(self):
        return self.com_object.PathFormat

    @PathFormat.setter
    def PathFormat(self, value):
        self.com_object.PathFormat = value

    @property
    def pathformat(self):
        """Lower case alias for PathFormat"""
        return self.PathFormat

    @pathformat.setter
    def pathformat(self, value):
        """Lower case alias for PathFormat.setter"""
        self.PathFormat = value

    @property
    def Ruler(self):
        return self.com_object.Ruler

    @property
    def ruler(self):
        """Lower case alias for Ruler"""
        return self.Ruler

    @property
    def TextRange(self):
        return self.com_object.TextRange

    @property
    def textrange(self):
        """Lower case alias for TextRange"""
        return self.TextRange

    @property
    def ThreeD(self):
        return ThreeDFormat(self.com_object.ThreeD)

    @property
    def threed(self):
        """Lower case alias for ThreeD"""
        return self.ThreeD

    @property
    def VerticalAnchor(self):
        return self.com_object.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.com_object.VerticalAnchor = value

    @property
    def verticalanchor(self):
        """Lower case alias for VerticalAnchor"""
        return self.VerticalAnchor

    @verticalanchor.setter
    def verticalanchor(self, value):
        """Lower case alias for VerticalAnchor.setter"""
        self.VerticalAnchor = value

    @property
    def WarpFormat(self):
        return self.com_object.WarpFormat

    @WarpFormat.setter
    def WarpFormat(self, value):
        self.com_object.WarpFormat = value

    @property
    def warpformat(self):
        """Lower case alias for WarpFormat"""
        return self.WarpFormat

    @warpformat.setter
    def warpformat(self, value):
        """Lower case alias for WarpFormat.setter"""
        self.WarpFormat = value

    @property
    def WordArtFormat(self):
        return self.com_object.WordArtFormat

    @WordArtFormat.setter
    def WordArtFormat(self, value):
        self.com_object.WordArtFormat = value

    @property
    def wordartformat(self):
        """Lower case alias for WordArtFormat"""
        return self.WordArtFormat

    @wordartformat.setter
    def wordartformat(self, value):
        """Lower case alias for WordArtFormat.setter"""
        self.WordArtFormat = value

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

    @property
    def actionsettings(self):
        """Lower case alias for ActionSettings"""
        return self.ActionSettings

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def BoundHeight(self):
        return self.com_object.BoundHeight

    @property
    def boundheight(self):
        """Lower case alias for BoundHeight"""
        return self.BoundHeight

    @property
    def BoundLeft(self):
        return self.com_object.BoundLeft

    @property
    def boundleft(self):
        """Lower case alias for BoundLeft"""
        return self.BoundLeft

    @property
    def BoundTop(self):
        return self.com_object.BoundTop

    @property
    def boundtop(self):
        """Lower case alias for BoundTop"""
        return self.BoundTop

    @property
    def BoundWidth(self):
        return self.com_object.BoundWidth

    @property
    def boundwidth(self):
        """Lower case alias for BoundWidth"""
        return self.BoundWidth

    @property
    def Count(self):
        return self.com_object.Count

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Font(self):
        return Font(self.com_object.Font)

    @property
    def font(self):
        """Lower case alias for Font"""
        return self.Font

    @property
    def IndentLevel(self):
        return self.com_object.IndentLevel

    @IndentLevel.setter
    def IndentLevel(self, value):
        self.com_object.IndentLevel = value

    @property
    def indentlevel(self):
        """Lower case alias for IndentLevel"""
        return self.IndentLevel

    @indentlevel.setter
    def indentlevel(self, value):
        """Lower case alias for IndentLevel.setter"""
        self.IndentLevel = value

    @property
    def LanguageID(self):
        return self.com_object.LanguageID

    @LanguageID.setter
    def LanguageID(self, value):
        self.com_object.LanguageID = value

    @property
    def languageid(self):
        """Lower case alias for LanguageID"""
        return self.LanguageID

    @languageid.setter
    def languageid(self, value):
        """Lower case alias for LanguageID.setter"""
        self.LanguageID = value

    @property
    def Length(self):
        return self.com_object.Length

    @property
    def length(self):
        """Lower case alias for Length"""
        return self.Length

    @property
    def ParagraphFormat(self):
        return ParagraphFormat(self.com_object.ParagraphFormat)

    @property
    def paragraphformat(self):
        """Lower case alias for ParagraphFormat"""
        return self.ParagraphFormat

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Start(self):
        return self.com_object.Start

    @property
    def start(self):
        """Lower case alias for Start"""
        return self.Start

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


class textrange2:

    def __init__(self, textrange2=None):
        self.com_object= textrange2

    @property
    def application(self):
        return self.com_object.application

    @property
    def boundheight(self):
        return self.com_object.boundheight

    @property
    def boundleft(self):
        return self.com_object.boundleft

    @property
    def boundtop(self):
        return self.com_object.boundtop

    @property
    def boundwidth(self):
        return self.com_object.boundwidth

    def characters(self, Start=None, Length=None):
        arguments = com_arguments([unwrap(a) for a in [Start, Length]])
        if hasattr(self.com_object, "Getcharacters"):
            return self.com_object.Getcharacters(*arguments)
        else:
            return self.com_object.characters(*arguments)

    @property
    def count(self):
        return self.com_object.count

    @property
    def creator(self):
        return self.com_object.creator

    @property
    def font(self):
        return Font(self.com_object.font)

    @property
    def languageid(self):
        return self.com_object.languageid

    @languageid.setter
    def languageid(self, value):
        self.com_object.languageid = value

    @property
    def length(self):
        return self.com_object.length

    def lines(self, Start=None, Length=None):
        arguments = com_arguments([unwrap(a) for a in [Start, Length]])
        if hasattr(self.com_object, "Getlines"):
            return self.com_object.Getlines(*arguments)
        else:
            return self.com_object.lines(*arguments)

    def mathzones(self, Start=None, Length=None):
        arguments = com_arguments([unwrap(a) for a in [Start, Length]])
        if hasattr(self.com_object, "Getmathzones"):
            return self.com_object.Getmathzones(*arguments)
        else:
            return self.com_object.mathzones(*arguments)

    @property
    def paragraphformat(self):
        return ParagraphFormat(self.com_object.paragraphformat)

    def paragraphs(self, Start=None, Length=None):
        arguments = com_arguments([unwrap(a) for a in [Start, Length]])
        if hasattr(self.com_object, "Getparagraphs"):
            return self.com_object.Getparagraphs(*arguments)
        else:
            return self.com_object.paragraphs(*arguments)

    @property
    def parent(self):
        return self.com_object.parent

    def runs(self, Start=None, Length=None):
        arguments = com_arguments([unwrap(a) for a in [Start, Length]])
        if hasattr(self.com_object, "Getruns"):
            return self.com_object.Getruns(*arguments)
        else:
            return self.com_object.runs(*arguments)

    def sentences(self, Start=None, Length=None):
        arguments = com_arguments([unwrap(a) for a in [Start, Length]])
        if hasattr(self.com_object, "Getsentences"):
            return TextRange2(self.com_object.Getsentences(*arguments))
        else:
            return TextRange2(self.com_object.sentences(*arguments))

    @property
    def start(self):
        return self.com_object.start

    @property
    def text(self):
        return self.com_object.text

    @text.setter
    def text(self, value):
        self.com_object.text = value

    def words(self, Start=None, Length=None):
        arguments = com_arguments([unwrap(a) for a in [Start, Length]])
        if hasattr(self.com_object, "Getwords"):
            return self.com_object.Getwords(*arguments)
        else:
            return self.com_object.words(*arguments)


class TextStyle:

    def __init__(self, textstyle=None):
        self.com_object= textstyle

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Levels(self):
        return TextStyleLevels(self.com_object.Levels)

    @property
    def levels(self):
        """Lower case alias for Levels"""
        return self.Levels

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Ruler(self):
        return Ruler(self.com_object.Ruler)

    @property
    def ruler(self):
        """Lower case alias for Ruler"""
        return self.Ruler

    @property
    def TextFrame(self):
        return TextFrame(self.com_object.TextFrame)

    @property
    def textframe(self):
        """Lower case alias for TextFrame"""
        return self.TextFrame


class TextStyleLevel:

    def __init__(self, textstylelevel=None):
        self.com_object= textstylelevel

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Font(self):
        return Font(self.com_object.Font)

    @property
    def font(self):
        """Lower case alias for Font"""
        return self.Font

    @property
    def ParagraphFormat(self):
        return ParagraphFormat(self.com_object.ParagraphFormat)

    @property
    def paragraphformat(self):
        """Lower case alias for ParagraphFormat"""
        return self.ParagraphFormat

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

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

    def Item(self, Type=None):
        arguments = com_arguments([unwrap(a) for a in [Type]])
        return TextStyle(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Type=None):
        arguments = [Type]
        return self.Item(*arguments)


class theme:

    def __init__(self, theme=None):
        self.com_object= theme

    @property
    def application(self):
        return self.com_object.application

    @property
    def parent(self):
        return self.com_object.parent

    @property
    def themevariants(self):
        return ThemeVariants(self.com_object.themevariants)


class themevariant:

    def __init__(self, themevariant=None):
        self.com_object= themevariant

    @property
    def application(self):
        return self.com_object.application

    @property
    def height(self):
        return self.com_object.height

    @property
    def id(self):
        return self.com_object.id

    @property
    def name(self):
        return self.com_object.name

    @property
    def parent(self):
        return self.com_object.parent

    @property
    def width(self):
        return self.com_object.width


class themevariants:

    def __init__(self, themevariants=None):
        self.com_object= themevariants

    def __call__(self, item):
        return themevariant(self.com_object(item))

    @property
    def application(self):
        return self.com_object.application

    @property
    def count(self):
        return self.com_object.count

    @property
    def parent(self):
        return self.com_object.parent


class ThreeDFormat:

    def __init__(self, threedformat=None):
        self.com_object= threedformat

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def BevelBottomDepth(self):
        return ThreeDFormat(self.com_object.BevelBottomDepth)

    @BevelBottomDepth.setter
    def BevelBottomDepth(self, value):
        self.com_object.BevelBottomDepth = value

    @property
    def bevelbottomdepth(self):
        """Lower case alias for BevelBottomDepth"""
        return self.BevelBottomDepth

    @bevelbottomdepth.setter
    def bevelbottomdepth(self, value):
        """Lower case alias for BevelBottomDepth.setter"""
        self.BevelBottomDepth = value

    @property
    def BevelBottomInset(self):
        return ThreeDFormat(self.com_object.BevelBottomInset)

    @BevelBottomInset.setter
    def BevelBottomInset(self, value):
        self.com_object.BevelBottomInset = value

    @property
    def bevelbottominset(self):
        """Lower case alias for BevelBottomInset"""
        return self.BevelBottomInset

    @bevelbottominset.setter
    def bevelbottominset(self, value):
        """Lower case alias for BevelBottomInset.setter"""
        self.BevelBottomInset = value

    @property
    def BevelBottomType(self):
        return self.com_object.BevelBottomType

    @BevelBottomType.setter
    def BevelBottomType(self, value):
        self.com_object.BevelBottomType = value

    @property
    def bevelbottomtype(self):
        """Lower case alias for BevelBottomType"""
        return self.BevelBottomType

    @bevelbottomtype.setter
    def bevelbottomtype(self, value):
        """Lower case alias for BevelBottomType.setter"""
        self.BevelBottomType = value

    @property
    def BevelTopDepth(self):
        return ThreeDFormat(self.com_object.BevelTopDepth)

    @BevelTopDepth.setter
    def BevelTopDepth(self, value):
        self.com_object.BevelTopDepth = value

    @property
    def beveltopdepth(self):
        """Lower case alias for BevelTopDepth"""
        return self.BevelTopDepth

    @beveltopdepth.setter
    def beveltopdepth(self, value):
        """Lower case alias for BevelTopDepth.setter"""
        self.BevelTopDepth = value

    @property
    def BevelTopInset(self):
        return ThreeDFormat(self.com_object.BevelTopInset)

    @BevelTopInset.setter
    def BevelTopInset(self, value):
        self.com_object.BevelTopInset = value

    @property
    def beveltopinset(self):
        """Lower case alias for BevelTopInset"""
        return self.BevelTopInset

    @beveltopinset.setter
    def beveltopinset(self, value):
        """Lower case alias for BevelTopInset.setter"""
        self.BevelTopInset = value

    @property
    def BevelTopType(self):
        return self.com_object.BevelTopType

    @BevelTopType.setter
    def BevelTopType(self, value):
        self.com_object.BevelTopType = value

    @property
    def beveltoptype(self):
        """Lower case alias for BevelTopType"""
        return self.BevelTopType

    @beveltoptype.setter
    def beveltoptype(self, value):
        """Lower case alias for BevelTopType.setter"""
        self.BevelTopType = value

    @property
    def ContourColor(self):
        return ColorFormat(self.com_object.ContourColor)

    @property
    def contourcolor(self):
        """Lower case alias for ContourColor"""
        return self.ContourColor

    @property
    def ContourWidth(self):
        return ThreeDFormat(self.com_object.ContourWidth)

    @ContourWidth.setter
    def ContourWidth(self, value):
        self.com_object.ContourWidth = value

    @property
    def contourwidth(self):
        """Lower case alias for ContourWidth"""
        return self.ContourWidth

    @contourwidth.setter
    def contourwidth(self, value):
        """Lower case alias for ContourWidth.setter"""
        self.ContourWidth = value

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Depth(self):
        return self.com_object.Depth

    @Depth.setter
    def Depth(self, value):
        self.com_object.Depth = value

    @property
    def depth(self):
        """Lower case alias for Depth"""
        return self.Depth

    @depth.setter
    def depth(self, value):
        """Lower case alias for Depth.setter"""
        self.Depth = value

    @property
    def ExtrusionColor(self):
        return ColorFormat(self.com_object.ExtrusionColor)

    @property
    def extrusioncolor(self):
        """Lower case alias for ExtrusionColor"""
        return self.ExtrusionColor

    @property
    def ExtrusionColorType(self):
        return self.com_object.ExtrusionColorType

    @ExtrusionColorType.setter
    def ExtrusionColorType(self, value):
        self.com_object.ExtrusionColorType = value

    @property
    def extrusioncolortype(self):
        """Lower case alias for ExtrusionColorType"""
        return self.ExtrusionColorType

    @extrusioncolortype.setter
    def extrusioncolortype(self, value):
        """Lower case alias for ExtrusionColorType.setter"""
        self.ExtrusionColorType = value

    @property
    def FieldOfView(self):
        return ThreeDFormat(self.com_object.FieldOfView)

    @FieldOfView.setter
    def FieldOfView(self, value):
        self.com_object.FieldOfView = value

    @property
    def fieldofview(self):
        """Lower case alias for FieldOfView"""
        return self.FieldOfView

    @fieldofview.setter
    def fieldofview(self, value):
        """Lower case alias for FieldOfView.setter"""
        self.FieldOfView = value

    @property
    def LightAngle(self):
        return self.com_object.LightAngle

    @LightAngle.setter
    def LightAngle(self, value):
        self.com_object.LightAngle = value

    @property
    def lightangle(self):
        """Lower case alias for LightAngle"""
        return self.LightAngle

    @lightangle.setter
    def lightangle(self, value):
        """Lower case alias for LightAngle.setter"""
        self.LightAngle = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Perspective(self):
        return self.com_object.Perspective

    @Perspective.setter
    def Perspective(self, value):
        self.com_object.Perspective = value

    @property
    def perspective(self):
        """Lower case alias for Perspective"""
        return self.Perspective

    @perspective.setter
    def perspective(self, value):
        """Lower case alias for Perspective.setter"""
        self.Perspective = value

    @property
    def PresetCamera(self):
        return ThreeDFormat(self.com_object.PresetCamera)

    @property
    def presetcamera(self):
        """Lower case alias for PresetCamera"""
        return self.PresetCamera

    @property
    def PresetExtrusionDirection(self):
        return self.com_object.PresetExtrusionDirection

    @property
    def presetextrusiondirection(self):
        """Lower case alias for PresetExtrusionDirection"""
        return self.PresetExtrusionDirection

    @property
    def PresetLighting(self):
        return ThreeDFormat(self.com_object.PresetLighting)

    @PresetLighting.setter
    def PresetLighting(self, value):
        self.com_object.PresetLighting = value

    @property
    def presetlighting(self):
        """Lower case alias for PresetLighting"""
        return self.PresetLighting

    @presetlighting.setter
    def presetlighting(self, value):
        """Lower case alias for PresetLighting.setter"""
        self.PresetLighting = value

    @property
    def PresetLightingDirection(self):
        return self.com_object.PresetLightingDirection

    @PresetLightingDirection.setter
    def PresetLightingDirection(self, value):
        self.com_object.PresetLightingDirection = value

    @property
    def presetlightingdirection(self):
        """Lower case alias for PresetLightingDirection"""
        return self.PresetLightingDirection

    @presetlightingdirection.setter
    def presetlightingdirection(self, value):
        """Lower case alias for PresetLightingDirection.setter"""
        self.PresetLightingDirection = value

    @property
    def PresetLightingSoftness(self):
        return self.com_object.PresetLightingSoftness

    @PresetLightingSoftness.setter
    def PresetLightingSoftness(self, value):
        self.com_object.PresetLightingSoftness = value

    @property
    def presetlightingsoftness(self):
        """Lower case alias for PresetLightingSoftness"""
        return self.PresetLightingSoftness

    @presetlightingsoftness.setter
    def presetlightingsoftness(self, value):
        """Lower case alias for PresetLightingSoftness.setter"""
        self.PresetLightingSoftness = value

    @property
    def PresetMaterial(self):
        return self.com_object.PresetMaterial

    @PresetMaterial.setter
    def PresetMaterial(self, value):
        self.com_object.PresetMaterial = value

    @property
    def presetmaterial(self):
        """Lower case alias for PresetMaterial"""
        return self.PresetMaterial

    @presetmaterial.setter
    def presetmaterial(self, value):
        """Lower case alias for PresetMaterial.setter"""
        self.PresetMaterial = value

    @property
    def PresetThreeDFormat(self):
        return self.com_object.PresetThreeDFormat

    @property
    def presetthreedformat(self):
        """Lower case alias for PresetThreeDFormat"""
        return self.PresetThreeDFormat

    @property
    def ProjectText(self):
        return self.com_object.ProjectText

    @ProjectText.setter
    def ProjectText(self, value):
        self.com_object.ProjectText = value

    @property
    def projecttext(self):
        """Lower case alias for ProjectText"""
        return self.ProjectText

    @projecttext.setter
    def projecttext(self, value):
        """Lower case alias for ProjectText.setter"""
        self.ProjectText = value

    @property
    def RotationX(self):
        return self.com_object.RotationX

    @RotationX.setter
    def RotationX(self, value):
        self.com_object.RotationX = value

    @property
    def rotationx(self):
        """Lower case alias for RotationX"""
        return self.RotationX

    @rotationx.setter
    def rotationx(self, value):
        """Lower case alias for RotationX.setter"""
        self.RotationX = value

    @property
    def RotationY(self):
        return self.com_object.RotationY

    @RotationY.setter
    def RotationY(self, value):
        self.com_object.RotationY = value

    @property
    def rotationy(self):
        """Lower case alias for RotationY"""
        return self.RotationY

    @rotationy.setter
    def rotationy(self, value):
        """Lower case alias for RotationY.setter"""
        self.RotationY = value

    @property
    def RotationZ(self):
        return ThreeDFormat(self.com_object.RotationZ)

    @RotationZ.setter
    def RotationZ(self, value):
        self.com_object.RotationZ = value

    @property
    def rotationz(self):
        """Lower case alias for RotationZ"""
        return self.RotationZ

    @rotationz.setter
    def rotationz(self, value):
        """Lower case alias for RotationZ.setter"""
        self.RotationZ = value

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

    @property
    def Z(self):
        return ThreeDFormat(self.com_object.Z)

    @Z.setter
    def Z(self, value):
        self.com_object.Z = value

    @property
    def z(self):
        """Lower case alias for Z"""
        return self.Z

    @z.setter
    def z(self, value):
        """Lower case alias for Z.setter"""
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

    @property
    def alignment(self):
        """Lower case alias for Alignment"""
        return self.Alignment

    @alignment.setter
    def alignment(self, value):
        """Lower case alias for Alignment.setter"""
        self.Alignment = value

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Depth(self):
        return self.com_object.Depth

    @property
    def depth(self):
        """Lower case alias for Depth"""
        return self.Depth

    @property
    def Font(self):
        return ChartFont(self.com_object.Font)

    @property
    def font(self):
        """Lower case alias for Font"""
        return self.Font

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

    @property
    def MultiLevel(self):
        return self.com_object.MultiLevel

    @MultiLevel.setter
    def MultiLevel(self, value):
        self.com_object.MultiLevel = value

    @property
    def multilevel(self):
        """Lower case alias for MultiLevel"""
        return self.MultiLevel

    @multilevel.setter
    def multilevel(self, value):
        """Lower case alias for MultiLevel.setter"""
        self.MultiLevel = value

    @property
    def Name(self):
        return self.com_object.Name

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def NumberFormat(self):
        return self.com_object.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value):
        self.com_object.NumberFormat = value

    @property
    def numberformat(self):
        """Lower case alias for NumberFormat"""
        return self.NumberFormat

    @numberformat.setter
    def numberformat(self, value):
        """Lower case alias for NumberFormat.setter"""
        self.NumberFormat = value

    @property
    def NumberFormatLinked(self):
        return self.com_object.NumberFormatLinked

    @NumberFormatLinked.setter
    def NumberFormatLinked(self, value):
        self.com_object.NumberFormatLinked = value

    @property
    def numberformatlinked(self):
        """Lower case alias for NumberFormatLinked"""
        return self.NumberFormatLinked

    @numberformatlinked.setter
    def numberformatlinked(self, value):
        """Lower case alias for NumberFormatLinked.setter"""
        self.NumberFormatLinked = value

    @property
    def NumberFormatLocal(self):
        return self.com_object.NumberFormatLocal

    @NumberFormatLocal.setter
    def NumberFormatLocal(self, value):
        self.com_object.NumberFormatLocal = value

    @property
    def numberformatlocal(self):
        """Lower case alias for NumberFormatLocal"""
        return self.NumberFormatLocal

    @numberformatlocal.setter
    def numberformatlocal(self, value):
        """Lower case alias for NumberFormatLocal.setter"""
        self.NumberFormatLocal = value

    @property
    def Offset(self):
        return self.com_object.Offset

    @Offset.setter
    def Offset(self, value):
        self.com_object.Offset = value

    @property
    def offset(self):
        """Lower case alias for Offset"""
        return self.Offset

    @offset.setter
    def offset(self, value):
        """Lower case alias for Offset.setter"""
        self.Offset = value

    @property
    def Orientation(self):
        return self.com_object.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.com_object.Orientation = value

    @property
    def orientation(self):
        """Lower case alias for Orientation"""
        return self.Orientation

    @orientation.setter
    def orientation(self, value):
        """Lower case alias for Orientation.setter"""
        self.Orientation = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def ReadingOrder(self):
        return XlReadingOrder(self.com_object.ReadingOrder)

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.com_object.ReadingOrder = value

    @property
    def readingorder(self):
        """Lower case alias for ReadingOrder"""
        return self.ReadingOrder

    @readingorder.setter
    def readingorder(self, value):
        """Lower case alias for ReadingOrder.setter"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def InteractiveSequences(self):
        return Sequences(self.com_object.InteractiveSequences)

    @property
    def interactivesequences(self):
        """Lower case alias for InteractiveSequences"""
        return self.InteractiveSequences

    @property
    def MainSequence(self):
        return Sequence(self.com_object.MainSequence)

    @property
    def mainsequence(self):
        """Lower case alias for MainSequence"""
        return self.MainSequence

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
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

    @property
    def accelerate(self):
        """Lower case alias for Accelerate"""
        return self.Accelerate

    @accelerate.setter
    def accelerate(self, value):
        """Lower case alias for Accelerate.setter"""
        self.Accelerate = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def AutoReverse(self):
        return self.com_object.AutoReverse

    @AutoReverse.setter
    def AutoReverse(self, value):
        self.com_object.AutoReverse = value

    @property
    def autoreverse(self):
        """Lower case alias for AutoReverse"""
        return self.AutoReverse

    @autoreverse.setter
    def autoreverse(self, value):
        """Lower case alias for AutoReverse.setter"""
        self.AutoReverse = value

    @property
    def BounceEnd(self):
        return self.com_object.BounceEnd

    @BounceEnd.setter
    def BounceEnd(self, value):
        self.com_object.BounceEnd = value

    @property
    def bounceend(self):
        """Lower case alias for BounceEnd"""
        return self.BounceEnd

    @bounceend.setter
    def bounceend(self, value):
        """Lower case alias for BounceEnd.setter"""
        self.BounceEnd = value

    @property
    def BounceEndIntensity(self):
        return self.com_object.BounceEndIntensity

    @BounceEndIntensity.setter
    def BounceEndIntensity(self, value):
        self.com_object.BounceEndIntensity = value

    @property
    def bounceendintensity(self):
        """Lower case alias for BounceEndIntensity"""
        return self.BounceEndIntensity

    @bounceendintensity.setter
    def bounceendintensity(self, value):
        """Lower case alias for BounceEndIntensity.setter"""
        self.BounceEndIntensity = value

    @property
    def Decelerate(self):
        return self.com_object.Decelerate

    @Decelerate.setter
    def Decelerate(self, value):
        self.com_object.Decelerate = value

    @property
    def decelerate(self):
        """Lower case alias for Decelerate"""
        return self.Decelerate

    @decelerate.setter
    def decelerate(self, value):
        """Lower case alias for Decelerate.setter"""
        self.Decelerate = value

    @property
    def duration(self):
        return self.com_object.duration

    @duration.setter
    def duration(self, value):
        self.com_object.duration = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def RepeatCount(self):
        return self.com_object.RepeatCount

    @RepeatCount.setter
    def RepeatCount(self, value):
        self.com_object.RepeatCount = value

    @property
    def repeatcount(self):
        """Lower case alias for RepeatCount"""
        return self.RepeatCount

    @repeatcount.setter
    def repeatcount(self, value):
        """Lower case alias for RepeatCount.setter"""
        self.RepeatCount = value

    @property
    def RepeatDuration(self):
        return self.com_object.RepeatDuration

    @RepeatDuration.setter
    def RepeatDuration(self, value):
        self.com_object.RepeatDuration = value

    @property
    def repeatduration(self):
        """Lower case alias for RepeatDuration"""
        return self.RepeatDuration

    @repeatduration.setter
    def repeatduration(self, value):
        """Lower case alias for RepeatDuration.setter"""
        self.RepeatDuration = value

    @property
    def Restart(self):
        return self.com_object.Restart

    @Restart.setter
    def Restart(self, value):
        self.com_object.Restart = value

    @property
    def restart(self):
        """Lower case alias for Restart"""
        return self.Restart

    @restart.setter
    def restart(self, value):
        """Lower case alias for Restart.setter"""
        self.Restart = value

    @property
    def RewindAtEnd(self):
        return self.com_object.RewindAtEnd

    @RewindAtEnd.setter
    def RewindAtEnd(self, value):
        self.com_object.RewindAtEnd = value

    @property
    def rewindatend(self):
        """Lower case alias for RewindAtEnd"""
        return self.RewindAtEnd

    @rewindatend.setter
    def rewindatend(self, value):
        """Lower case alias for RewindAtEnd.setter"""
        self.RewindAtEnd = value

    @property
    def SmoothEnd(self):
        return self.com_object.SmoothEnd

    @SmoothEnd.setter
    def SmoothEnd(self, value):
        self.com_object.SmoothEnd = value

    @property
    def smoothend(self):
        """Lower case alias for SmoothEnd"""
        return self.SmoothEnd

    @smoothend.setter
    def smoothend(self, value):
        """Lower case alias for SmoothEnd.setter"""
        self.SmoothEnd = value

    @property
    def SmoothStart(self):
        return self.com_object.SmoothStart

    @SmoothStart.setter
    def SmoothStart(self, value):
        self.com_object.SmoothStart = value

    @property
    def smoothstart(self):
        """Lower case alias for SmoothStart"""
        return self.SmoothStart

    @smoothstart.setter
    def smoothstart(self, value):
        """Lower case alias for SmoothStart.setter"""
        self.SmoothStart = value

    @property
    def Speed(self):
        return self.com_object.Speed

    @Speed.setter
    def Speed(self, value):
        self.com_object.Speed = value

    @property
    def speed(self):
        """Lower case alias for Speed"""
        return self.Speed

    @speed.setter
    def speed(self, value):
        """Lower case alias for Speed.setter"""
        self.Speed = value

    @property
    def triggerBookmark(self):
        return self.com_object.triggerBookmark

    @triggerBookmark.setter
    def triggerBookmark(self, value):
        self.com_object.triggerBookmark = value

    @property
    def triggerbookmark(self):
        """Lower case alias for triggerBookmark"""
        return self.triggerBookmark

    @triggerbookmark.setter
    def triggerbookmark(self, value):
        """Lower case alias for triggerBookmark.setter"""
        self.triggerBookmark = value

    @property
    def TriggerDelayTime(self):
        return self.com_object.TriggerDelayTime

    @TriggerDelayTime.setter
    def TriggerDelayTime(self, value):
        self.com_object.TriggerDelayTime = value

    @property
    def triggerdelaytime(self):
        """Lower case alias for TriggerDelayTime"""
        return self.TriggerDelayTime

    @triggerdelaytime.setter
    def triggerdelaytime(self, value):
        """Lower case alias for TriggerDelayTime.setter"""
        self.TriggerDelayTime = value

    @property
    def TriggerShape(self):
        return self.com_object.TriggerShape

    @TriggerShape.setter
    def TriggerShape(self, value):
        self.com_object.TriggerShape = value

    @property
    def triggershape(self):
        """Lower case alias for TriggerShape"""
        return self.TriggerShape

    @triggershape.setter
    def triggershape(self, value):
        """Lower case alias for TriggerShape.setter"""
        self.TriggerShape = value

    @property
    def TriggerType(self):
        return self.com_object.TriggerType

    @TriggerType.setter
    def TriggerType(self, value):
        self.com_object.TriggerType = value

    @property
    def triggertype(self):
        """Lower case alias for TriggerType"""
        return self.TriggerType

    @triggertype.setter
    def triggertype(self, value):
        """Lower case alias for TriggerType.setter"""
        self.TriggerType = value


class Trendline:

    def __init__(self, trendline=None):
        self.com_object= trendline

    @property
    def Application(self):
        return self.com_object.Application

    @property
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Backward2(self):
        return self.com_object.Backward2

    @Backward2.setter
    def Backward2(self, value):
        self.com_object.Backward2 = value

    @property
    def backward2(self):
        """Lower case alias for Backward2"""
        return self.Backward2

    @backward2.setter
    def backward2(self, value):
        """Lower case alias for Backward2.setter"""
        self.Backward2 = value

    @property
    def Border(self):
        return ChartBorder(self.com_object.Border)

    @property
    def border(self):
        """Lower case alias for Border"""
        return self.Border

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def DataLabel(self):
        return DataLabel(self.com_object.DataLabel)

    @property
    def datalabel(self):
        """Lower case alias for DataLabel"""
        return self.DataLabel

    @property
    def DisplayEquation(self):
        return self.com_object.DisplayEquation

    @DisplayEquation.setter
    def DisplayEquation(self, value):
        self.com_object.DisplayEquation = value

    @property
    def displayequation(self):
        """Lower case alias for DisplayEquation"""
        return self.DisplayEquation

    @displayequation.setter
    def displayequation(self, value):
        """Lower case alias for DisplayEquation.setter"""
        self.DisplayEquation = value

    @property
    def DisplayRSquared(self):
        return self.com_object.DisplayRSquared

    @DisplayRSquared.setter
    def DisplayRSquared(self, value):
        self.com_object.DisplayRSquared = value

    @property
    def displayrsquared(self):
        """Lower case alias for DisplayRSquared"""
        return self.DisplayRSquared

    @displayrsquared.setter
    def displayrsquared(self, value):
        """Lower case alias for DisplayRSquared.setter"""
        self.DisplayRSquared = value

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

    @property
    def Forward2(self):
        return self.com_object.Forward2

    @Forward2.setter
    def Forward2(self, value):
        self.com_object.Forward2 = value

    @property
    def forward2(self):
        """Lower case alias for Forward2"""
        return self.Forward2

    @forward2.setter
    def forward2(self, value):
        """Lower case alias for Forward2.setter"""
        self.Forward2 = value

    @property
    def Index(self):
        return self.com_object.Index

    @property
    def index(self):
        """Lower case alias for Index"""
        return self.Index

    @property
    def Intercept(self):
        return self.com_object.Intercept

    @Intercept.setter
    def Intercept(self, value):
        self.com_object.Intercept = value

    @property
    def intercept(self):
        """Lower case alias for Intercept"""
        return self.Intercept

    @intercept.setter
    def intercept(self, value):
        """Lower case alias for Intercept.setter"""
        self.Intercept = value

    @property
    def InterceptIsAuto(self):
        return self.com_object.InterceptIsAuto

    @InterceptIsAuto.setter
    def InterceptIsAuto(self, value):
        self.com_object.InterceptIsAuto = value

    @property
    def interceptisauto(self):
        """Lower case alias for InterceptIsAuto"""
        return self.InterceptIsAuto

    @interceptisauto.setter
    def interceptisauto(self, value):
        """Lower case alias for InterceptIsAuto.setter"""
        self.InterceptIsAuto = value

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
    def NameIsAuto(self):
        return self.com_object.NameIsAuto

    @NameIsAuto.setter
    def NameIsAuto(self, value):
        self.com_object.NameIsAuto = value

    @property
    def nameisauto(self):
        """Lower case alias for NameIsAuto"""
        return self.NameIsAuto

    @nameisauto.setter
    def nameisauto(self, value):
        """Lower case alias for NameIsAuto.setter"""
        self.NameIsAuto = value

    @property
    def Order(self):
        return self.com_object.Order

    @Order.setter
    def Order(self, value):
        self.com_object.Order = value

    @property
    def order(self):
        """Lower case alias for Order"""
        return self.Order

    @order.setter
    def order(self, value):
        """Lower case alias for Order.setter"""
        self.Order = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Period(self):
        return self.com_object.Period

    @Period.setter
    def Period(self, value):
        self.com_object.Period = value

    @property
    def period(self):
        """Lower case alias for Period"""
        return self.Period

    @period.setter
    def period(self, value):
        """Lower case alias for Period.setter"""
        self.Period = value

    @property
    def Type(self):
        return XlTrendlineType(self.com_object.Type)

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Count(self):
        return self.com_object.Count

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Fill(self):
        return FillFormat(self.com_object.Fill)

    @property
    def fill(self):
        """Lower case alias for Fill"""
        return self.Fill

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def DisplaySlideMiniature(self):
        return self.com_object.DisplaySlideMiniature

    @DisplaySlideMiniature.setter
    def DisplaySlideMiniature(self, value):
        self.com_object.DisplaySlideMiniature = value

    @property
    def displayslideminiature(self):
        """Lower case alias for DisplaySlideMiniature"""
        return self.DisplaySlideMiniature

    @displayslideminiature.setter
    def displayslideminiature(self, value):
        """Lower case alias for DisplaySlideMiniature.setter"""
        self.DisplaySlideMiniature = value

    @property
    def MediaControlsHeight(self):
        return self.com_object.MediaControlsHeight

    @property
    def mediacontrolsheight(self):
        """Lower case alias for MediaControlsHeight"""
        return self.MediaControlsHeight

    @property
    def MediaControlsLeft(self):
        return self.com_object.MediaControlsLeft

    @property
    def mediacontrolsleft(self):
        """Lower case alias for MediaControlsLeft"""
        return self.MediaControlsLeft

    @property
    def MediaControlsTop(self):
        return self.com_object.MediaControlsTop

    @property
    def mediacontrolstop(self):
        """Lower case alias for MediaControlsTop"""
        return self.MediaControlsTop

    @property
    def MediaControlsVisible(self):
        return self.com_object.MediaControlsVisible

    @property
    def mediacontrolsvisible(self):
        """Lower case alias for MediaControlsVisible"""
        return self.MediaControlsVisible

    @property
    def MediaControlsWidth(self):
        return self.com_object.MediaControlsWidth

    @property
    def mediacontrolswidth(self):
        """Lower case alias for MediaControlsWidth"""
        return self.MediaControlsWidth

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PrintOptions(self):
        return PrintOptions(self.com_object.PrintOptions)

    @property
    def printoptions(self):
        """Lower case alias for PrintOptions"""
        return self.PrintOptions

    @property
    def Slide(self):
        return Slide(self.com_object.Slide)

    @Slide.setter
    def Slide(self, value):
        self.com_object.Slide = value

    @property
    def slide(self):
        """Lower case alias for Slide"""
        return self.Slide

    @slide.setter
    def slide(self, value):
        """Lower case alias for Slide.setter"""
        self.Slide = value

    @property
    def Type(self):
        return self.com_object.Type

    @property
    def type(self):
        """Lower case alias for Type"""
        return self.Type

    @property
    def Zoom(self):
        return self.com_object.Zoom

    @Zoom.setter
    def Zoom(self, value):
        self.com_object.Zoom = value

    @property
    def zoom(self):
        """Lower case alias for Zoom"""
        return self.Zoom

    @zoom.setter
    def zoom(self, value):
        """Lower case alias for Zoom.setter"""
        self.Zoom = value

    @property
    def ZoomToFit(self):
        return self.com_object.ZoomToFit

    @ZoomToFit.setter
    def ZoomToFit(self, value):
        self.com_object.ZoomToFit = value

    @property
    def zoomtofit(self):
        """Lower case alias for ZoomToFit"""
        return self.ZoomToFit

    @zoomtofit.setter
    def zoomtofit(self, value):
        """Lower case alias for ZoomToFit.setter"""
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
    def application(self):
        """Lower case alias for Application"""
        return self.Application

    @property
    def Creator(self):
        return self.com_object.Creator

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @property
    def Format(self):
        return ChartFormat(self.com_object.Format)

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

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
    def PictureType(self):
        return self.com_object.PictureType

    @PictureType.setter
    def PictureType(self, value):
        self.com_object.PictureType = value

    @property
    def picturetype(self):
        """Lower case alias for PictureType"""
        return self.PictureType

    @picturetype.setter
    def picturetype(self, value):
        """Lower case alias for PictureType.setter"""
        self.PictureType = value

    @property
    def PictureUnit(self):
        return self.com_object.PictureUnit

    @PictureUnit.setter
    def PictureUnit(self, value):
        self.com_object.PictureUnit = value

    @property
    def pictureunit(self):
        """Lower case alias for PictureUnit"""
        return self.PictureUnit

    @pictureunit.setter
    def pictureunit(self, value):
        """Lower case alias for PictureUnit.setter"""
        self.PictureUnit = value

    @property
    def Thickness(self):
        return self.com_object.Thickness

    @Thickness.setter
    def Thickness(self, value):
        self.com_object.Thickness = value

    @property
    def thickness(self):
        """Lower case alias for Thickness"""
        return self.Thickness

    @thickness.setter
    def thickness(self, value):
        """Lower case alias for Thickness.setter"""
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

# xlbinstype enumeration
xlBinsTypeAutomatic = 0
xlBinsTypeCategorical = 1
xlBinsTypeManual = 2
xlBinsTypeBinSize = 3
xlBinsTypeBinCount = 4

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

# xlparentdatalabeloptions enumeration
xlParentDataLabelOptionsNone = 0
xlParentDataLabelOptionsBanner = 1
xlParentDataLabelOptionsOverlapping = 2

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

# xlpiesliceindex enumeration
xlCenterPoint = 5
xlInnerCenterPoint = 8
xlInnerClockwisePoint = 7
xlInnerCounterClockwisePoint = 9
xlMidClockwiseRadiusPoint = 4
xlMidCounterClockwiseRadiusPoint = 6
xlOuterCenterPoint = 2
xlOuterClockwisePoint = 3
xlOuterCounterClockwisePoint = 1

# xlpieslicelocation enumeration
xlCenterPoint = 5
xlInnerCenterPoint = 8
xlInnerClockwisePoint = 7
xlInnerCounterClockwisePoint = 9
xlMidClockwiseRadiusPoint = 4
xlMidCounterClockwiseRadiusPoint = 6
xlOuterCenterPoint = 2
xlOuterClockwisePoint = 3
xlOuterCounterClockwisePoint = 1

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
