import win32com.client

class AccessObject:

    def __init__(self, accessobject=None):
        self.accessobject = accessobject

    @property
    def CurrentView(self):
        return AcCurrentView(self.accessobject.CurrentView)

    @property
    def DateCreated(self):
        return self.accessobject.DateCreated

    @property
    def DateModified(self):
        return self.accessobject.DateModified

    @property
    def FullName(self):
        return self.accessobject.FullName

    @FullName.setter
    def FullName(self, value):
        self.accessobject.FullName = value

    @property
    def IsLoaded(self):
        return self.accessobject.IsLoaded

    @property
    def IsWeb(self):
        return self.accessobject.IsWeb

    @property
    def Name(self):
        return self.accessobject.Name

    @property
    def Parent(self):
        return self.accessobject.Parent

    @property
    def Properties(self):
        return AccessObjectProperties(self.accessobject.Properties)

    @property
    def Type(self):
        return AcObjectType(self.accessobject.Type)

    def GetDependencyInfo(self):
        return self.accessobject.GetDependencyInfo()

    def IsDependentUpon(self, *args, ObjectType=None, ObjectName=None):
        arguments = {"ObjectType": ObjectType, "ObjectName": ObjectName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.accessobject.IsDependentUpon(*args, **arguments)

class AccessObjectProperties:

    def __init__(self, accessobjectproperties=None):
        self.accessobjectproperties = accessobjectproperties

    @property
    def Application(self):
        return self.accessobjectproperties.Application

    @property
    def Count(self):
        return self.accessobjectproperties.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.accessobjectproperties.Item):
            return self.accessobjectproperties.Item(*args, **arguments)
        else:
            return self.accessobjectproperties.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.accessobjectproperties.Parent

    def Add(self, *args, PropertyName=None, Value=None):
        arguments = {"PropertyName": PropertyName, "Value": Value}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.accessobjectproperties.Add(*args, **arguments)

    def Remove(self, *args, Item=None):
        arguments = {"Item": Item}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.accessobjectproperties.Remove(*args, **arguments)

class AccessObjectProperty:

    def __init__(self, accessobjectproperty=None):
        self.accessobjectproperty = accessobjectproperty

    @property
    def Name(self):
        return self.accessobjectproperty.Name

    @property
    def Value(self):
        return self.accessobjectproperty.Value

    @Value.setter
    def Value(self, value):
        self.accessobjectproperty.Value = value

class AdditionalData:

    def __init__(self, additionaldata=None):
        self.additionaldata = additionaldata

    @property
    def Count(self):
        return self.additionaldata.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.additionaldata.Item):
            return self.additionaldata.Item(*args, **arguments)
        else:
            return self.additionaldata.GetItem(*args, **arguments)

    @property
    def Name(self):
        return self.additionaldata.Name

    @Name.setter
    def Name(self, value):
        self.additionaldata.Name = value

    def Add(self, *args, var=None):
        arguments = {"var": var}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.additionaldata.Add(*args, **arguments)

class AllDatabaseDiagrams:

    def __init__(self, alldatabasediagrams=None):
        self.alldatabasediagrams = alldatabasediagrams

    @property
    def Application(self):
        return self.alldatabasediagrams.Application

    @property
    def Count(self):
        return self.alldatabasediagrams.Count

    def Item(self, *args, var=None):
        arguments = {"var": var}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.alldatabasediagrams.Item):
            return self.alldatabasediagrams.Item(*args, **arguments)
        else:
            return self.alldatabasediagrams.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.alldatabasediagrams.Parent

class AllForms:

    def __init__(self, allforms=None):
        self.allforms = allforms

    @property
    def Application(self):
        return self.allforms.Application

    @property
    def Count(self):
        return self.allforms.Count

    def Item(self, *args, var=None):
        arguments = {"var": var}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.allforms.Item):
            return self.allforms.Item(*args, **arguments)
        else:
            return self.allforms.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.allforms.Parent

class AllFunctions:

    def __init__(self, allfunctions=None):
        self.allfunctions = allfunctions

    @property
    def Application(self):
        return self.allfunctions.Application

    @property
    def Count(self):
        return self.allfunctions.Count

    def Item(self, *args, var=None):
        arguments = {"var": var}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.allfunctions.Item):
            return self.allfunctions.Item(*args, **arguments)
        else:
            return self.allfunctions.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.allfunctions.Parent

class AllModules:

    def __init__(self, allmodules=None):
        self.allmodules = allmodules

    @property
    def Application(self):
        return self.allmodules.Application

    @property
    def Count(self):
        return self.allmodules.Count

    def Item(self, *args, var=None):
        arguments = {"var": var}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.allmodules.Item):
            return self.allmodules.Item(*args, **arguments)
        else:
            return self.allmodules.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.allmodules.Parent

class AllQueries:

    def __init__(self, allqueries=None):
        self.allqueries = allqueries

    @property
    def Application(self):
        return self.allqueries.Application

    @property
    def Count(self):
        return self.allqueries.Count

    def Item(self, *args, var=None):
        arguments = {"var": var}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.allqueries.Item):
            return self.allqueries.Item(*args, **arguments)
        else:
            return self.allqueries.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.allqueries.Parent

class AllReports:

    def __init__(self, allreports=None):
        self.allreports = allreports

    @property
    def Application(self):
        return self.allreports.Application

    @property
    def Count(self):
        return self.allreports.Count

    def Item(self, *args, var=None):
        arguments = {"var": var}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.allreports.Item):
            return self.allreports.Item(*args, **arguments)
        else:
            return self.allreports.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.allreports.Parent

class AllStoredProcedures:

    def __init__(self, allstoredprocedures=None):
        self.allstoredprocedures = allstoredprocedures

    @property
    def Application(self):
        return self.allstoredprocedures.Application

    @property
    def Count(self):
        return self.allstoredprocedures.Count

    def Item(self, *args, var=None):
        arguments = {"var": var}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.allstoredprocedures.Item):
            return self.allstoredprocedures.Item(*args, **arguments)
        else:
            return self.allstoredprocedures.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.allstoredprocedures.Parent

class AllTables:

    def __init__(self, alltables=None):
        self.alltables = alltables

    @property
    def Application(self):
        return self.alltables.Application

    @property
    def Count(self):
        return self.alltables.Count

    def Item(self, *args, var=None):
        arguments = {"var": var}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.alltables.Item):
            return self.alltables.Item(*args, **arguments)
        else:
            return self.alltables.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.alltables.Parent

class AllViews:

    def __init__(self, allviews=None):
        self.allviews = allviews

    @property
    def Application(self):
        return self.allviews.Application

    @property
    def Count(self):
        return self.allviews.Count

    def Item(self, *args, var=None):
        arguments = {"var": var}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.allviews.Item):
            return self.allviews.Item(*args, **arguments)
        else:
            return self.allviews.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.allviews.Parent

class Application:

    def __init__(self, application=None):
        self.application = application

    def new(self):
        self.application = win32com.client.Dispatch("Access.Application")
        return self

    @property
    def AppIcon(self):
        return self.application.AppIcon

    @property
    def Application(self):
        return self.application.Application

    @property
    def AppTitle(self):
        return self.application.AppTitle

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
    def BrokenReference(self):
        return self.application.BrokenReference

    @property
    def Build(self):
        return self.application.Build

    @property
    def CodeContextObject(self):
        return self.application.CodeContextObject

    @property
    def CodeData(self):
        return self.application.CodeData

    @property
    def CodeProject(self):
        return self.application.CodeProject

    @property
    def COMAddIns(self):
        return self.application.COMAddIns

    @property
    def CommandBars(self):
        return self.application.CommandBars

    @property
    def CurrentData(self):
        return self.application.CurrentData

    @property
    def CurrentObjectName(self):
        return self.application.CurrentObjectName

    @property
    def CurrentObjectType(self):
        return self.application.CurrentObjectType

    @property
    def CurrentProject(self):
        return self.application.CurrentProject

    @property
    def DBEngine(self):
        return self.application.DBEngine

    @property
    def DoCmd(self):
        return self.application.DoCmd

    @property
    def FeatureInstall(self):
        return self.application.FeatureInstall

    @FeatureInstall.setter
    def FeatureInstall(self, value):
        self.application.FeatureInstall = value

    def FileDialog(self, *args, dialogType=None):
        arguments = {"dialogType": dialogType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.application.FileDialog):
            return self.application.FileDialog(*args, **arguments)
        else:
            return self.application.GetFileDialog(*args, **arguments)

    @property
    def Forms(self):
        return self.application.Forms

    @property
    def IsCompiled(self):
        return self.application.IsCompiled

    @property
    def LanguageSettings(self):
        return self.application.LanguageSettings

    @property
    def MacroError(self):
        return MacroError(self.application.MacroError)

    @property
    def MenuBar(self):
        return self.application.MenuBar

    @MenuBar.setter
    def MenuBar(self, value):
        self.application.MenuBar = value

    @property
    def Modules(self):
        return self.application.Modules

    @property
    def Name(self):
        return self.application.Name

    @property
    def NewFileTaskPane(self):
        return self.application.NewFileTaskPane

    @property
    def Parent(self):
        return self.application.Parent

    @property
    def Printer(self):
        return Printer(self.application.Printer)

    @Printer.setter
    def Printer(self, value):
        self.application.Printer = value

    @property
    def Printers(self):
        return Printers(self.application.Printers)

    @property
    def ProductCode(self):
        return self.application.ProductCode

    @property
    def References(self):
        return self.application.References

    @property
    def Reports(self):
        return self.application.Reports

    @property
    def Screen(self):
        return self.application.Screen

    @property
    def ShortcutMenuBar(self):
        return self.application.ShortcutMenuBar

    @ShortcutMenuBar.setter
    def ShortcutMenuBar(self, value):
        self.application.ShortcutMenuBar = value

    @property
    def TempVars(self):
        return TempVar(self.application.TempVars)

    @property
    def UserControl(self):
        return self.application.UserControl

    @UserControl.setter
    def UserControl(self, value):
        self.application.UserControl = value

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
    def WebServices(self):
        return self.application.WebServices

    def AccessError(self, *args, ErrorNumber=None):
        arguments = {"ErrorNumber": ErrorNumber}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.AccessError(*args, **arguments)

    def AddToFavorites(self):
        self.application.AddToFavorites()

    def BuildCriteria(self, *args, Field=None, FieldType=None, Expression=None):
        arguments = {"Field": Field, "FieldType": FieldType, "Expression": Expression}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.BuildCriteria(*args, **arguments)

    def CloseCurrentDatabase(self):
        return self.application.CloseCurrentDatabase()

    def CodeDb(self):
        return self.application.CodeDb()

    def ColumnHistory(self, *args, TableName=None, ColumnName=None, queryString=None):
        arguments = {"TableName": TableName, "ColumnName": ColumnName, "queryString": queryString}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.ColumnHistory(*args, **arguments)

    def ConvertAccessProject(self, *args, SourceFilename=None, DestinationFilename=None, DestinationFileFormat=None):
        arguments = {"SourceFilename": SourceFilename, "DestinationFilename": DestinationFilename, "DestinationFileFormat": DestinationFileFormat}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.ConvertAccessProject(*args, **arguments)

    def CreateAccessProject(self, *args, filepath=None, Connect=None):
        arguments = {"filepath": filepath, "Connect": Connect}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.CreateAccessProject(*args, **arguments)

    def CreateAdditionalData(self):
        return self.application.CreateAdditionalData()

    def CreateControl(self, *args, FormName=None, ControlType=None, Section=None, Parent=None, ColumnName=None, Left=None, Top=None, Width=None, Height=None):
        arguments = {"FormName": FormName, "ControlType": ControlType, "Section": Section, "Parent": Parent, "ColumnName": ColumnName, "Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.CreateControl(*args, **arguments)

    def CreateForm(self, *args, Database=None, FormTemplate=None):
        arguments = {"Database": Database, "FormTemplate": FormTemplate}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.CreateForm(*args, **arguments)

    def CreateGroupLevel(self, *args, ReportName=None, Expression=None, Header=None, Footer=None):
        arguments = {"ReportName": ReportName, "Expression": Expression, "Header": Header, "Footer": Footer}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.CreateGroupLevel(*args, **arguments)

    def CreateReport(self, *args, Database=None, ReportTemplate=None):
        arguments = {"Database": Database, "ReportTemplate": ReportTemplate}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.CreateReport(*args, **arguments)

    def CreateReportControl(self, *args, ReportName=None, ControlType=None, Section=None, Parent=None, ColumnName=None, Left=None, Top=None, Width=None, Height=None):
        arguments = {"ReportName": ReportName, "ControlType": ControlType, "Section": Section, "Parent": Parent, "ColumnName": ColumnName, "Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.CreateReportControl(*args, **arguments)

    def CurrentDb(self):
        return self.application.CurrentDb()

    def CurrentUser(self):
        return self.application.CurrentUser()

    def CurrentWebUser(self, *args, DisplayOption=None):
        arguments = {"DisplayOption": DisplayOption}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.CurrentWebUser(*args, **arguments)

    def CurrentWebUserGroups(self, *args, DisplayOption=None):
        arguments = {"DisplayOption": DisplayOption}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.CurrentWebUserGroups(*args, **arguments)

    def DCount(self, *args, Expr=None, Domain=None, Criteria=None):
        arguments = {"Expr": Expr, "Domain": Domain, "Criteria": Criteria}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.DCount(*args, **arguments)

    def DDEExecute(self, *args, ChanNum=None, Command=None):
        arguments = {"ChanNum": ChanNum, "Command": Command}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.DDEExecute(*args, **arguments)

    def DDEInitiate(self, *args, Application=None, Topic=None):
        arguments = {"Application": Application, "Topic": Topic}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.DDEInitiate(*args, **arguments)

    def DDEPoke(self, *args, ChanNum=None, Item=None, Data=None):
        arguments = {"ChanNum": ChanNum, "Item": Item, "Data": Data}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.DDEPoke(*args, **arguments)

    def DDERequest(self, *args, ChanNum=None, Item=None):
        arguments = {"ChanNum": ChanNum, "Item": Item}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.DDERequest(*args, **arguments)

    def DDETerminate(self, *args, ChanNum=None):
        arguments = {"ChanNum": ChanNum}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.DDETerminate(*args, **arguments)

    def DDETerminateAll(self):
        return self.application.DDETerminateAll()

    def DefaultWorkspaceClone(self):
        return self.application.DefaultWorkspaceClone()

    def DeleteControl(self, *args, FormName=None, ControlName=None):
        arguments = {"FormName": FormName, "ControlName": ControlName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.DeleteControl(*args, **arguments)

    def DeleteReportControl(self, *args, ReportName=None, ControlName=None):
        arguments = {"ReportName": ReportName, "ControlName": ControlName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.DeleteReportControl(*args, **arguments)

    def DFirst(self, *args, Expr=None, Domain=None, Criteria=None):
        arguments = {"Expr": Expr, "Domain": Domain, "Criteria": Criteria}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.DFirst(*args, **arguments)

    def DirtyObject(self, *args, ObjectType=None, ObjectName=None):
        arguments = {"ObjectType": ObjectType, "ObjectName": ObjectName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.DirtyObject(*args, **arguments)

    def DLast(self, *args, Expr=None, Domain=None, Criteria=None):
        arguments = {"Expr": Expr, "Domain": Domain, "Criteria": Criteria}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.DLast(*args, **arguments)

    def DLookup(self, *args, Expr=None, Domain=None, Criteria=None):
        arguments = {"Expr": Expr, "Domain": Domain, "Criteria": Criteria}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.DLookup(*args, **arguments)

    def DMax(self, *args, Expr=None, Domain=None, Criteria=None):
        arguments = {"Expr": Expr, "Domain": Domain, "Criteria": Criteria}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.DMax(*args, **arguments)

    def DMin(self, *args, Expr=None, Domain=None, Criteria=None):
        arguments = {"Expr": Expr, "Domain": Domain, "Criteria": Criteria}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.DMin(*args, **arguments)

    def DStDev(self, *args, Expr=None, Domain=None, Criteria=None):
        arguments = {"Expr": Expr, "Domain": Domain, "Criteria": Criteria}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.DStDev(*args, **arguments)

    def DStDevP(self, *args, Expr=None, Domain=None, Criteria=None):
        arguments = {"Expr": Expr, "Domain": Domain, "Criteria": Criteria}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.DStDevP(*args, **arguments)

    def DSum(self, *args, Expr=None, Domain=None, Criteria=None):
        arguments = {"Expr": Expr, "Domain": Domain, "Criteria": Criteria}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.DSum(*args, **arguments)

    def DVar(self, *args, Expr=None, Domain=None, Criteria=None):
        arguments = {"Expr": Expr, "Domain": Domain, "Criteria": Criteria}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.DVar(*args, **arguments)

    def DVarP(self, *args, Expr=None, Domain=None, Criteria=None):
        arguments = {"Expr": Expr, "Domain": Domain, "Criteria": Criteria}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.DVarP(*args, **arguments)

    def Echo(self, *args, EchoOn=None, bstrStatusBarText=None):
        arguments = {"EchoOn": EchoOn, "bstrStatusBarText": bstrStatusBarText}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.Echo(*args, **arguments)

    def EuroConvert(self, *args, Number=None, SourceCurrency=None, TargetCurrency=None, FullPrecision=None, TriangulationPrecision=None):
        arguments = {"Number": Number, "SourceCurrency": SourceCurrency, "TargetCurrency": TargetCurrency, "FullPrecision": FullPrecision, "TriangulationPrecision": TriangulationPrecision}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.EuroConvert(*args, **arguments)

    def Eval(self, *args, StringExpr=None):
        arguments = {"StringExpr": StringExpr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.Eval(*args, **arguments)

    def ExportNavigationPane(self, *args, Path=None):
        arguments = {"Path": Path}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.ExportNavigationPane(*args, **arguments)

    def ExportXML(self, *args, ObjectType=None, DataSource=None, DataTarget=None, SchemaTarget=None, PresentationTarget=None, ImageTarget=None, Encoding=None, OtherFlags=None, WhereCondition=None, AdditionalData=None):
        arguments = {"ObjectType": ObjectType, "DataSource": DataSource, "DataTarget": DataTarget, "SchemaTarget": SchemaTarget, "PresentationTarget": PresentationTarget, "ImageTarget": ImageTarget, "Encoding": Encoding, "OtherFlags": OtherFlags, "WhereCondition": WhereCondition, "AdditionalData": AdditionalData}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.ExportXML(*args, **arguments)

    def FollowHyperlink(self, *args, Address=None, SubAddress=None, NewWindow=None, AddHistory=None, ExtraInfo=None, Method=None, HeaderInfo=None):
        arguments = {"Address": Address, "SubAddress": SubAddress, "NewWindow": NewWindow, "AddHistory": AddHistory, "ExtraInfo": ExtraInfo, "Method": Method, "HeaderInfo": HeaderInfo}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.FollowHyperlink(*args, **arguments)

    def GetHiddenAttribute(self, *args, ObjectType=None, ObjectName=None):
        arguments = {"ObjectType": ObjectType, "ObjectName": ObjectName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.GetHiddenAttribute(*args, **arguments)

    def GetOption(self, *args, OptionName=None):
        arguments = {"OptionName": OptionName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.GetOption(*args, **arguments)

    def GUIDFromString(self, *args, String=None):
        arguments = {"String": String}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.GUIDFromString(*args, **arguments)

    def HtmlEncode(self, *args, PlainText=None, Length=None):
        arguments = {"PlainText": PlainText, "Length": Length}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.HtmlEncode(*args, **arguments)

    def hWndAccessApp(self):
        return self.application.hWndAccessApp()

    def HyperlinkPart(self, *args, Hyperlink=None, Part=None):
        arguments = {"Hyperlink": Hyperlink, "Part": Part}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.HyperlinkPart(*args, **arguments)

    def ImportNavigationPane(self, *args, Path=None, fAppendOnly=None):
        arguments = {"Path": Path, "fAppendOnly": fAppendOnly}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.ImportNavigationPane(*args, **arguments)

    def ImportXML(self, *args, DataSource=None, ImportOptions=None):
        arguments = {"DataSource": DataSource, "ImportOptions": ImportOptions}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.ImportXML(*args, **arguments)

    def InstantiateTemplate(self, *args, Path=None):
        arguments = {"Path": Path}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.InstantiateTemplate(*args, **arguments)

    def IsCurrentWebUserInGroup(self, *args, GroupNameOrID=None):
        arguments = {"GroupNameOrID": GroupNameOrID}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.IsCurrentWebUserInGroup(*args, **arguments)

    def LoadCustomUI(self, *args, CustomUIName=None, CustomUIXML=None):
        arguments = {"CustomUIName": CustomUIName, "CustomUIXML": CustomUIXML}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.LoadCustomUI(*args, **arguments)

    def LoadFromAXL(self, *args, ObjectType=None, ObjectName=None, FileName=None):
        arguments = {"ObjectType": ObjectType, "ObjectName": ObjectName, "FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.LoadFromAXL(*args, **arguments)

    def LoadPicture(self, *args, FileName=None):
        arguments = {"FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.LoadPicture(*args, **arguments)

    def NewAccessProject(self, *args, filepath=None, Connect=None):
        arguments = {"filepath": filepath, "Connect": Connect}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.NewAccessProject(*args, **arguments)

    def NewCurrentDatabase(self, *args, filepath=None, FileFormat=None, Template=None, SiteAddress=None, ListID=None):
        arguments = {"filepath": filepath, "FileFormat": FileFormat, "Template": Template, "SiteAddress": SiteAddress, "ListID": ListID}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.NewCurrentDatabase(*args, **arguments)

    def Nz(self, *args, Value=None, ValueIfNull=None):
        arguments = {"Value": Value, "ValueIfNull": ValueIfNull}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.Nz(*args, **arguments)

    def OpenAccessProject(self, *args, filepath=None, Exclusive=None):
        arguments = {"filepath": filepath, "Exclusive": Exclusive}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.OpenAccessProject(*args, **arguments)

    def OpenCurrentDatabase(self, *args, filepath=None, Exclusive=None, bstrPassword=None):
        arguments = {"filepath": filepath, "Exclusive": Exclusive, "bstrPassword": bstrPassword}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.OpenCurrentDatabase(*args, **arguments)

    def PlainText(self, *args, RichText=None, Length=None):
        arguments = {"RichText": RichText, "Length": Length}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.PlainText(*args, **arguments)

    def Quit(self, *args, Option=None):
        arguments = {"Option": Option}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.Quit(*args, **arguments)

    def RefreshDatabaseWindow(self):
        return self.application.RefreshDatabaseWindow()

    def RefreshTitleBar(self):
        return self.application.RefreshTitleBar()

    def Run(self, *args, Procedure=None, Arg1=None, Arg2=None, Arg3=None, Arg4=None, Arg5=None, Arg6=None, Arg7=None, Arg8=None, Arg9=None, Arg10=None, Arg11=None, Arg12=None, Arg13=None, Arg14=None, Arg15=None, Arg16=None, Arg17=None, Arg18=None, Arg19=None, Arg20=None, Arg21=None, Arg22=None, Arg23=None, Arg24=None, Arg25=None, Arg26=None, Arg27=None, Arg28=None, Arg29=None, Arg30=None):
        arguments = {"Procedure": Procedure, "Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4, "Arg5": Arg5, "Arg6": Arg6, "Arg7": Arg7, "Arg8": Arg8, "Arg9": Arg9, "Arg10": Arg10, "Arg11": Arg11, "Arg12": Arg12, "Arg13": Arg13, "Arg14": Arg14, "Arg15": Arg15, "Arg16": Arg16, "Arg17": Arg17, "Arg18": Arg18, "Arg19": Arg19, "Arg20": Arg20, "Arg21": Arg21, "Arg22": Arg22, "Arg23": Arg23, "Arg24": Arg24, "Arg25": Arg25, "Arg26": Arg26, "Arg27": Arg27, "Arg28": Arg28, "Arg29": Arg29, "Arg30": Arg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.Run(*args, **arguments)

    def RunCommand(self, *args, Command=None):
        arguments = {"Command": Command}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.RunCommand(*args, **arguments)

    def SaveAsAXL(self, *args, ObjectType=None, ObjectName=None, FileName=None):
        arguments = {"ObjectType": ObjectType, "ObjectName": ObjectName, "FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.SaveAsAXL(*args, **arguments)

    def SaveAsTemplate(self, *args, Path=None, Title=None, IconPath=None, CoreTable=None, Category=None, PreviewPath=None, Description=None, InstantiationForm=None, ApplicationPart=None, IncludeData=None):
        arguments = {"Path": Path, "Title": Title, "IconPath": IconPath, "CoreTable": CoreTable, "Category": Category, "PreviewPath": PreviewPath, "Description": Description, "InstantiationForm": InstantiationForm, "ApplicationPart": ApplicationPart, "IncludeData": IncludeData}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.SaveAsTemplate(*args, **arguments)

    def SetDefaultWorkgroupFile(self, *args, Path=None):
        arguments = {"Path": Path}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.SetDefaultWorkgroupFile(*args, **arguments)

    def SetHiddenAttribute(self, *args, ObjectType=None, ObjectName=None, fHidden=None):
        arguments = {"ObjectType": ObjectType, "ObjectName": ObjectName, "fHidden": fHidden}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.SetHiddenAttribute(*args, **arguments)

    def SetOption(self, *args, OptionName=None, Setting=None):
        arguments = {"OptionName": OptionName, "Setting": Setting}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.SetOption(*args, **arguments)

    def StringFromGUID(self, *args, Guid=None):
        arguments = {"Guid": Guid}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.StringFromGUID(*args, **arguments)

    def SysCmd(self, *args, Action=None, Argument2=None, Argument3=None):
        arguments = {"Action": Action, "Argument2": Argument2, "Argument3": Argument3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.SysCmd(*args, **arguments)

    def TransformXML(self, *args, DataSource=None, TransformSource=None, OutputTarget=None, WellFormedXMLOutput=None, ScriptOption=None):
        arguments = {"DataSource": DataSource, "TransformSource": TransformSource, "OutputTarget": OutputTarget, "WellFormedXMLOutput": WellFormedXMLOutput, "ScriptOption": ScriptOption}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.TransformXML(*args, **arguments)

class Attachment:

    def __init__(self, attachment=None):
        self.attachment = attachment

    @property
    def AddColon(self):
        return self.attachment.AddColon

    @AddColon.setter
    def AddColon(self, value):
        self.attachment.AddColon = value

    @property
    def Application(self):
        return self.attachment.Application

    @property
    def AttachmentCount(self):
        return self.attachment.AttachmentCount

    @property
    def AutoLabel(self):
        return self.attachment.AutoLabel

    @AutoLabel.setter
    def AutoLabel(self, value):
        self.attachment.AutoLabel = value

    @property
    def BackColor(self):
        return self.attachment.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.attachment.BackColor = value

    @property
    def BackShade(self):
        return self.attachment.BackShade

    @BackShade.setter
    def BackShade(self, value):
        self.attachment.BackShade = value

    @property
    def BackStyle(self):
        return self.attachment.BackStyle

    @BackStyle.setter
    def BackStyle(self, value):
        self.attachment.BackStyle = value

    @property
    def BackThemeColorIndex(self):
        return self.attachment.BackThemeColorIndex

    @BackThemeColorIndex.setter
    def BackThemeColorIndex(self, value):
        self.attachment.BackThemeColorIndex = value

    @property
    def BackTint(self):
        return self.attachment.BackTint

    @BackTint.setter
    def BackTint(self, value):
        self.attachment.BackTint = value

    @property
    def BorderColor(self):
        return self.attachment.BorderColor

    @BorderColor.setter
    def BorderColor(self, value):
        self.attachment.BorderColor = value

    @property
    def BorderShade(self):
        return self.attachment.BorderShade

    @BorderShade.setter
    def BorderShade(self, value):
        self.attachment.BorderShade = value

    @property
    def BorderStyle(self):
        return self.attachment.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.attachment.BorderStyle = value

    @property
    def BorderThemeColorIndex(self):
        return self.attachment.BorderThemeColorIndex

    @BorderThemeColorIndex.setter
    def BorderThemeColorIndex(self, value):
        self.attachment.BorderThemeColorIndex = value

    @property
    def BorderTint(self):
        return self.attachment.BorderTint

    @BorderTint.setter
    def BorderTint(self, value):
        self.attachment.BorderTint = value

    @property
    def BorderWidth(self):
        return self.attachment.BorderWidth

    @BorderWidth.setter
    def BorderWidth(self, value):
        self.attachment.BorderWidth = value

    @property
    def BottomPadding(self):
        return self.attachment.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.attachment.BottomPadding = value

    @property
    def ColumnHidden(self):
        return self.attachment.ColumnHidden

    @ColumnHidden.setter
    def ColumnHidden(self, value):
        self.attachment.ColumnHidden = value

    @property
    def ColumnOrder(self):
        return self.attachment.ColumnOrder

    @ColumnOrder.setter
    def ColumnOrder(self, value):
        self.attachment.ColumnOrder = value

    @property
    def ColumnWidth(self):
        return self.attachment.ColumnWidth

    @ColumnWidth.setter
    def ColumnWidth(self, value):
        self.attachment.ColumnWidth = value

    @property
    def Controls(self):
        return Controls(self.attachment.Controls)

    @property
    def ControlSource(self):
        return self.attachment.ControlSource

    @ControlSource.setter
    def ControlSource(self, value):
        self.attachment.ControlSource = value

    @property
    def ControlTipText(self):
        return self.attachment.ControlTipText

    @ControlTipText.setter
    def ControlTipText(self, value):
        self.attachment.ControlTipText = value

    @property
    def ControlType(self):
        return self.attachment.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.attachment.ControlType = value

    @property
    def CurrentAttachment(self):
        return self.attachment.CurrentAttachment

    @CurrentAttachment.setter
    def CurrentAttachment(self, value):
        self.attachment.CurrentAttachment = value

    @property
    def DefaultPicture(self):
        return self.attachment.DefaultPicture

    @DefaultPicture.setter
    def DefaultPicture(self, value):
        self.attachment.DefaultPicture = value

    @property
    def DefaultPictureType(self):
        return self.attachment.DefaultPictureType

    @DefaultPictureType.setter
    def DefaultPictureType(self, value):
        self.attachment.DefaultPictureType = value

    @property
    def DisplayAs(self):
        return self.attachment.DisplayAs

    @DisplayAs.setter
    def DisplayAs(self, value):
        self.attachment.DisplayAs = value

    @property
    def DisplayWhen(self):
        return self.attachment.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.attachment.DisplayWhen = value

    @property
    def Enabled(self):
        return self.attachment.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.attachment.Enabled = value

    @property
    def EventProcPrefix(self):
        return self.attachment.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.attachment.EventProcPrefix = value

    def FileName(self, *args, var=None):
        arguments = {"var": var}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.attachment.FileName):
            return self.attachment.FileName(*args, **arguments)
        else:
            return self.attachment.GetFileName(*args, **arguments)

    def FileType(self, *args, var=None):
        arguments = {"var": var}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.attachment.FileType):
            return self.attachment.FileType(*args, **arguments)
        else:
            return self.attachment.GetFileType(*args, **arguments)

    def FileURL(self, *args, var=None):
        arguments = {"var": var}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.attachment.FileURL):
            return self.attachment.FileURL(*args, **arguments)
        else:
            return self.attachment.GetFileURL(*args, **arguments)

    @property
    def GridlineColor(self):
        return self.attachment.GridlineColor

    @GridlineColor.setter
    def GridlineColor(self, value):
        self.attachment.GridlineColor = value

    @property
    def GridlineShade(self):
        return self.attachment.GridlineShade

    @GridlineShade.setter
    def GridlineShade(self, value):
        self.attachment.GridlineShade = value

    @property
    def GridlineStyleBottom(self):
        return self.attachment.GridlineStyleBottom

    @GridlineStyleBottom.setter
    def GridlineStyleBottom(self, value):
        self.attachment.GridlineStyleBottom = value

    @property
    def GridlineStyleLeft(self):
        return self.attachment.GridlineStyleLeft

    @GridlineStyleLeft.setter
    def GridlineStyleLeft(self, value):
        self.attachment.GridlineStyleLeft = value

    @property
    def GridlineStyleRight(self):
        return self.attachment.GridlineStyleRight

    @GridlineStyleRight.setter
    def GridlineStyleRight(self, value):
        self.attachment.GridlineStyleRight = value

    @property
    def GridlineStyleTop(self):
        return self.attachment.GridlineStyleTop

    @GridlineStyleTop.setter
    def GridlineStyleTop(self, value):
        self.attachment.GridlineStyleTop = value

    @property
    def GridlineThemeColorIndex(self):
        return self.attachment.GridlineThemeColorIndex

    @GridlineThemeColorIndex.setter
    def GridlineThemeColorIndex(self, value):
        self.attachment.GridlineThemeColorIndex = value

    @property
    def GridlineTint(self):
        return self.attachment.GridlineTint

    @GridlineTint.setter
    def GridlineTint(self, value):
        self.attachment.GridlineTint = value

    @property
    def GridlineWidthBottom(self):
        return self.attachment.GridlineWidthBottom

    @GridlineWidthBottom.setter
    def GridlineWidthBottom(self, value):
        self.attachment.GridlineWidthBottom = value

    @property
    def GridlineWidthLeft(self):
        return self.attachment.GridlineWidthLeft

    @GridlineWidthLeft.setter
    def GridlineWidthLeft(self, value):
        self.attachment.GridlineWidthLeft = value

    @property
    def GridlineWidthRight(self):
        return self.attachment.GridlineWidthRight

    @GridlineWidthRight.setter
    def GridlineWidthRight(self, value):
        self.attachment.GridlineWidthRight = value

    @property
    def GridlineWidthTop(self):
        return self.attachment.GridlineWidthTop

    @GridlineWidthTop.setter
    def GridlineWidthTop(self, value):
        self.attachment.GridlineWidthTop = value

    @property
    def Height(self):
        return self.attachment.Height

    @Height.setter
    def Height(self, value):
        self.attachment.Height = value

    @property
    def HelpContextId(self):
        return self.attachment.HelpContextId

    @HelpContextId.setter
    def HelpContextId(self, value):
        self.attachment.HelpContextId = value

    @property
    def HorizontalAnchor(self):
        return self.attachment.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.attachment.HorizontalAnchor = value

    @property
    def InSelection(self):
        return self.attachment.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.attachment.InSelection = value

    @property
    def IsVisible(self):
        return self.attachment.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.attachment.IsVisible = value

    @property
    def LabelAlign(self):
        return self.attachment.LabelAlign

    @LabelAlign.setter
    def LabelAlign(self, value):
        self.attachment.LabelAlign = value

    @property
    def LabelX(self):
        return self.attachment.LabelX

    @LabelX.setter
    def LabelX(self, value):
        self.attachment.LabelX = value

    @property
    def LabelY(self):
        return self.attachment.LabelY

    @LabelY.setter
    def LabelY(self, value):
        self.attachment.LabelY = value

    @property
    def Layout(self):
        return AcLayoutType(self.attachment.Layout)

    @property
    def LayoutID(self):
        return self.attachment.LayoutID

    @property
    def Left(self):
        return self.attachment.Left

    @Left.setter
    def Left(self, value):
        self.attachment.Left = value

    @property
    def LeftPadding(self):
        return self.attachment.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.attachment.LeftPadding = value

    @property
    def Locked(self):
        return self.attachment.Locked

    @Locked.setter
    def Locked(self, value):
        self.attachment.Locked = value

    @property
    def Name(self):
        return self.attachment.Name

    @Name.setter
    def Name(self, value):
        self.attachment.Name = value

    @property
    def OldBorderStyle(self):
        return self.attachment.OldBorderStyle

    @OldBorderStyle.setter
    def OldBorderStyle(self, value):
        self.attachment.OldBorderStyle = value

    @property
    def OldValue(self):
        return self.attachment.OldValue

    @property
    def OnAttachmentCurrent(self):
        return self.attachment.OnAttachmentCurrent

    @OnAttachmentCurrent.setter
    def OnAttachmentCurrent(self, value):
        self.attachment.OnAttachmentCurrent = value

    @property
    def OnChange(self):
        return self.attachment.OnChange

    @OnChange.setter
    def OnChange(self, value):
        self.attachment.OnChange = value

    @property
    def OnClick(self):
        return self.attachment.OnClick

    @OnClick.setter
    def OnClick(self, value):
        self.attachment.OnClick = value

    @property
    def OnDblClick(self):
        return self.attachment.OnDblClick

    @OnDblClick.setter
    def OnDblClick(self, value):
        self.attachment.OnDblClick = value

    @property
    def OnDirty(self):
        return self.attachment.OnDirty

    @OnDirty.setter
    def OnDirty(self, value):
        self.attachment.OnDirty = value

    @property
    def OnEnter(self):
        return self.attachment.OnEnter

    @OnEnter.setter
    def OnEnter(self, value):
        self.attachment.OnEnter = value

    @property
    def OnExit(self):
        return self.attachment.OnExit

    @OnExit.setter
    def OnExit(self, value):
        self.attachment.OnExit = value

    @property
    def OnGotFocus(self):
        return self.attachment.OnGotFocus

    @OnGotFocus.setter
    def OnGotFocus(self, value):
        self.attachment.OnGotFocus = value

    @property
    def OnKeyDown(self):
        return self.attachment.OnKeyDown

    @OnKeyDown.setter
    def OnKeyDown(self, value):
        self.attachment.OnKeyDown = value

    @property
    def OnKeyPress(self):
        return self.attachment.OnKeyPress

    @OnKeyPress.setter
    def OnKeyPress(self, value):
        self.attachment.OnKeyPress = value

    @property
    def OnKeyUp(self):
        return self.attachment.OnKeyUp

    @OnKeyUp.setter
    def OnKeyUp(self, value):
        self.attachment.OnKeyUp = value

    @property
    def OnLostFocus(self):
        return self.attachment.OnLostFocus

    @OnLostFocus.setter
    def OnLostFocus(self, value):
        self.attachment.OnLostFocus = value

    @property
    def OnMouseDown(self):
        return self.attachment.OnMouseDown

    @OnMouseDown.setter
    def OnMouseDown(self, value):
        self.attachment.OnMouseDown = value

    @property
    def OnMouseMove(self):
        return self.attachment.OnMouseMove

    @OnMouseMove.setter
    def OnMouseMove(self, value):
        self.attachment.OnMouseMove = value

    @property
    def OnMouseUp(self):
        return self.attachment.OnMouseUp

    @OnMouseUp.setter
    def OnMouseUp(self, value):
        self.attachment.OnMouseUp = value

    @property
    def Parent(self):
        return self.attachment.Parent

    @property
    def PictureAlignment(self):
        return self.attachment.PictureAlignment

    @PictureAlignment.setter
    def PictureAlignment(self, value):
        self.attachment.PictureAlignment = value

    @property
    def PictureSizeMode(self):
        return self.attachment.PictureSizeMode

    @PictureSizeMode.setter
    def PictureSizeMode(self, value):
        self.attachment.PictureSizeMode = value

    @property
    def PictureTiling(self):
        return self.attachment.PictureTiling

    @PictureTiling.setter
    def PictureTiling(self, value):
        self.attachment.PictureTiling = value

    @property
    def Properties(self):
        return Properties(self.attachment.Properties)

    @property
    def RightPadding(self):
        return self.attachment.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.attachment.RightPadding = value

    @property
    def Section(self):
        return self.attachment.Section

    @Section.setter
    def Section(self, value):
        self.attachment.Section = value

    @property
    def ShortcutMenuBar(self):
        return self.attachment.ShortcutMenuBar

    @ShortcutMenuBar.setter
    def ShortcutMenuBar(self, value):
        self.attachment.ShortcutMenuBar = value

    @property
    def SpecialEffect(self):
        return self.attachment.SpecialEffect

    @SpecialEffect.setter
    def SpecialEffect(self, value):
        self.attachment.SpecialEffect = value

    @property
    def StatusBarText(self):
        return self.attachment.StatusBarText

    @StatusBarText.setter
    def StatusBarText(self, value):
        self.attachment.StatusBarText = value

    @property
    def TabIndex(self):
        return self.attachment.TabIndex

    @TabIndex.setter
    def TabIndex(self, value):
        self.attachment.TabIndex = value

    @property
    def TabStop(self):
        return self.attachment.TabStop

    @TabStop.setter
    def TabStop(self, value):
        self.attachment.TabStop = value

    @property
    def Tag(self):
        return self.attachment.Tag

    @Tag.setter
    def Tag(self, value):
        self.attachment.Tag = value

    @property
    def Top(self):
        return self.attachment.Top

    @Top.setter
    def Top(self, value):
        self.attachment.Top = value

    @property
    def TopPadding(self):
        return self.attachment.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.attachment.TopPadding = value

    @property
    def VerticalAnchor(self):
        return self.attachment.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.attachment.VerticalAnchor = value

    @property
    def Visible(self):
        return self.attachment.Visible

    @Visible.setter
    def Visible(self, value):
        self.attachment.Visible = value

    @property
    def Width(self):
        return self.attachment.Width

    @Width.setter
    def Width(self, value):
        self.attachment.Width = value

    def Back(self):
        self.attachment.Back()

    def Forward(self):
        self.attachment.Forward()

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.attachment.Move(*args, **arguments)

    def Requery(self):
        self.attachment.Requery()

    def SetFocus(self):
        self.attachment.SetFocus()

    def SizeToFit(self):
        self.attachment.SizeToFit()

class AutoCorrect:

    def __init__(self, autocorrect=None):
        self.autocorrect = autocorrect

    @property
    def DisplayAutoCorrectOptions(self):
        return self.autocorrect.DisplayAutoCorrectOptions

    @DisplayAutoCorrectOptions.setter
    def DisplayAutoCorrectOptions(self, value):
        self.autocorrect.DisplayAutoCorrectOptions = value

class BoundObjectFrame:

    def __init__(self, boundobjectframe=None):
        self.boundobjectframe = boundobjectframe

    @property
    def Action(self):
        return self.boundobjectframe.Action

    @Action.setter
    def Action(self, value):
        self.boundobjectframe.Action = value

    @property
    def AddColon(self):
        return self.boundobjectframe.AddColon

    @AddColon.setter
    def AddColon(self, value):
        self.boundobjectframe.AddColon = value

    @property
    def Application(self):
        return self.boundobjectframe.Application

    @property
    def AutoActivate(self):
        return self.boundobjectframe.AutoActivate

    @AutoActivate.setter
    def AutoActivate(self, value):
        self.boundobjectframe.AutoActivate = value

    @property
    def AutoLabel(self):
        return self.boundobjectframe.AutoLabel

    @AutoLabel.setter
    def AutoLabel(self, value):
        self.boundobjectframe.AutoLabel = value

    @property
    def BackColor(self):
        return self.boundobjectframe.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.boundobjectframe.BackColor = value

    @property
    def BackShade(self):
        return self.boundobjectframe.BackShade

    @BackShade.setter
    def BackShade(self, value):
        self.boundobjectframe.BackShade = value

    @property
    def BackStyle(self):
        return self.boundobjectframe.BackStyle

    @BackStyle.setter
    def BackStyle(self, value):
        self.boundobjectframe.BackStyle = value

    @property
    def BackThemeColorIndex(self):
        return self.boundobjectframe.BackThemeColorIndex

    @BackThemeColorIndex.setter
    def BackThemeColorIndex(self, value):
        self.boundobjectframe.BackThemeColorIndex = value

    @property
    def BackTint(self):
        return self.boundobjectframe.BackTint

    @BackTint.setter
    def BackTint(self, value):
        self.boundobjectframe.BackTint = value

    @property
    def BorderColor(self):
        return self.boundobjectframe.BorderColor

    @BorderColor.setter
    def BorderColor(self, value):
        self.boundobjectframe.BorderColor = value

    @property
    def BorderShade(self):
        return self.boundobjectframe.BorderShade

    @BorderShade.setter
    def BorderShade(self, value):
        self.boundobjectframe.BorderShade = value

    @property
    def BorderStyle(self):
        return self.boundobjectframe.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.boundobjectframe.BorderStyle = value

    @property
    def BorderThemeColorIndex(self):
        return self.boundobjectframe.BorderThemeColorIndex

    @BorderThemeColorIndex.setter
    def BorderThemeColorIndex(self, value):
        self.boundobjectframe.BorderThemeColorIndex = value

    @property
    def BorderTint(self):
        return self.boundobjectframe.BorderTint

    @BorderTint.setter
    def BorderTint(self, value):
        self.boundobjectframe.BorderTint = value

    @property
    def BorderWidth(self):
        return self.boundobjectframe.BorderWidth

    @BorderWidth.setter
    def BorderWidth(self, value):
        self.boundobjectframe.BorderWidth = value

    @property
    def BottomPadding(self):
        return self.boundobjectframe.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.boundobjectframe.BottomPadding = value

    @property
    def Class(self):
        return self.boundobjectframe.Class

    @Class.setter
    def Class(self, value):
        self.boundobjectframe.Class = value

    @property
    def ColumnHidden(self):
        return self.boundobjectframe.ColumnHidden

    @ColumnHidden.setter
    def ColumnHidden(self, value):
        self.boundobjectframe.ColumnHidden = value

    @property
    def ColumnOrder(self):
        return self.boundobjectframe.ColumnOrder

    @ColumnOrder.setter
    def ColumnOrder(self, value):
        self.boundobjectframe.ColumnOrder = value

    @property
    def ColumnWidth(self):
        return self.boundobjectframe.ColumnWidth

    @ColumnWidth.setter
    def ColumnWidth(self, value):
        self.boundobjectframe.ColumnWidth = value

    @property
    def Controls(self):
        return Controls(self.boundobjectframe.Controls)

    @property
    def ControlSource(self):
        return self.boundobjectframe.ControlSource

    @ControlSource.setter
    def ControlSource(self, value):
        self.boundobjectframe.ControlSource = value

    @property
    def ControlTipText(self):
        return self.boundobjectframe.ControlTipText

    @ControlTipText.setter
    def ControlTipText(self, value):
        self.boundobjectframe.ControlTipText = value

    @property
    def ControlType(self):
        return self.boundobjectframe.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.boundobjectframe.ControlType = value

    @property
    def DisplayType(self):
        return self.boundobjectframe.DisplayType

    @DisplayType.setter
    def DisplayType(self, value):
        self.boundobjectframe.DisplayType = value

    @property
    def DisplayWhen(self):
        return self.boundobjectframe.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.boundobjectframe.DisplayWhen = value

    @property
    def Enabled(self):
        return self.boundobjectframe.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.boundobjectframe.Enabled = value

    @property
    def EventProcPrefix(self):
        return self.boundobjectframe.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.boundobjectframe.EventProcPrefix = value

    @property
    def GridlineColor(self):
        return self.boundobjectframe.GridlineColor

    @GridlineColor.setter
    def GridlineColor(self, value):
        self.boundobjectframe.GridlineColor = value

    @property
    def GridlineShade(self):
        return self.boundobjectframe.GridlineShade

    @GridlineShade.setter
    def GridlineShade(self, value):
        self.boundobjectframe.GridlineShade = value

    @property
    def GridlineStyleBottom(self):
        return self.boundobjectframe.GridlineStyleBottom

    @GridlineStyleBottom.setter
    def GridlineStyleBottom(self, value):
        self.boundobjectframe.GridlineStyleBottom = value

    @property
    def GridlineStyleLeft(self):
        return self.boundobjectframe.GridlineStyleLeft

    @GridlineStyleLeft.setter
    def GridlineStyleLeft(self, value):
        self.boundobjectframe.GridlineStyleLeft = value

    @property
    def GridlineStyleRight(self):
        return self.boundobjectframe.GridlineStyleRight

    @GridlineStyleRight.setter
    def GridlineStyleRight(self, value):
        self.boundobjectframe.GridlineStyleRight = value

    @property
    def GridlineStyleTop(self):
        return self.boundobjectframe.GridlineStyleTop

    @GridlineStyleTop.setter
    def GridlineStyleTop(self, value):
        self.boundobjectframe.GridlineStyleTop = value

    @property
    def GridlineThemeColorIndex(self):
        return self.boundobjectframe.GridlineThemeColorIndex

    @GridlineThemeColorIndex.setter
    def GridlineThemeColorIndex(self, value):
        self.boundobjectframe.GridlineThemeColorIndex = value

    @property
    def GridlineTint(self):
        return self.boundobjectframe.GridlineTint

    @GridlineTint.setter
    def GridlineTint(self, value):
        self.boundobjectframe.GridlineTint = value

    @property
    def GridlineWidthBottom(self):
        return self.boundobjectframe.GridlineWidthBottom

    @GridlineWidthBottom.setter
    def GridlineWidthBottom(self, value):
        self.boundobjectframe.GridlineWidthBottom = value

    @property
    def GridlineWidthLeft(self):
        return self.boundobjectframe.GridlineWidthLeft

    @GridlineWidthLeft.setter
    def GridlineWidthLeft(self, value):
        self.boundobjectframe.GridlineWidthLeft = value

    @property
    def GridlineWidthRight(self):
        return self.boundobjectframe.GridlineWidthRight

    @GridlineWidthRight.setter
    def GridlineWidthRight(self, value):
        self.boundobjectframe.GridlineWidthRight = value

    @property
    def GridlineWidthTop(self):
        return self.boundobjectframe.GridlineWidthTop

    @GridlineWidthTop.setter
    def GridlineWidthTop(self, value):
        self.boundobjectframe.GridlineWidthTop = value

    @property
    def Height(self):
        return self.boundobjectframe.Height

    @Height.setter
    def Height(self, value):
        self.boundobjectframe.Height = value

    @property
    def HelpContextId(self):
        return self.boundobjectframe.HelpContextId

    @HelpContextId.setter
    def HelpContextId(self, value):
        self.boundobjectframe.HelpContextId = value

    @property
    def HorizontalAnchor(self):
        return self.boundobjectframe.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.boundobjectframe.HorizontalAnchor = value

    @property
    def InSelection(self):
        return self.boundobjectframe.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.boundobjectframe.InSelection = value

    @property
    def IsVisible(self):
        return self.boundobjectframe.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.boundobjectframe.IsVisible = value

    @property
    def LabelAlign(self):
        return self.boundobjectframe.LabelAlign

    @LabelAlign.setter
    def LabelAlign(self, value):
        self.boundobjectframe.LabelAlign = value

    @property
    def LabelX(self):
        return self.boundobjectframe.LabelX

    @LabelX.setter
    def LabelX(self, value):
        self.boundobjectframe.LabelX = value

    @property
    def LabelY(self):
        return self.boundobjectframe.LabelY

    @LabelY.setter
    def LabelY(self, value):
        self.boundobjectframe.LabelY = value

    @property
    def Layout(self):
        return AcLayoutType(self.boundobjectframe.Layout)

    @property
    def LayoutID(self):
        return self.boundobjectframe.LayoutID

    @property
    def Left(self):
        return self.boundobjectframe.Left

    @Left.setter
    def Left(self, value):
        self.boundobjectframe.Left = value

    @property
    def LeftPadding(self):
        return self.boundobjectframe.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.boundobjectframe.LeftPadding = value

    @property
    def Locked(self):
        return self.boundobjectframe.Locked

    @Locked.setter
    def Locked(self, value):
        self.boundobjectframe.Locked = value

    @property
    def Name(self):
        return self.boundobjectframe.Name

    @Name.setter
    def Name(self, value):
        self.boundobjectframe.Name = value

    @property
    def Object(self):
        return self.boundobjectframe.Object

    @property
    def ObjectPalette(self):
        return self.boundobjectframe.ObjectPalette

    @ObjectPalette.setter
    def ObjectPalette(self, value):
        self.boundobjectframe.ObjectPalette = value

    def ObjectVerbs(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.boundobjectframe.ObjectVerbs):
            return self.boundobjectframe.ObjectVerbs(*args, **arguments)
        else:
            return self.boundobjectframe.GetObjectVerbs(*args, **arguments)

    @property
    def ObjectVerbsCount(self):
        return self.boundobjectframe.ObjectVerbsCount

    @property
    def OldBorderStyle(self):
        return self.boundobjectframe.OldBorderStyle

    @OldBorderStyle.setter
    def OldBorderStyle(self, value):
        self.boundobjectframe.OldBorderStyle = value

    @property
    def OldValue(self):
        return self.boundobjectframe.OldValue

    @property
    def OLEType(self):
        return self.boundobjectframe.OLEType

    @OLEType.setter
    def OLEType(self, value):
        self.boundobjectframe.OLEType = value

    @property
    def OLETypeAllowed(self):
        return self.boundobjectframe.OLETypeAllowed

    @OLETypeAllowed.setter
    def OLETypeAllowed(self, value):
        self.boundobjectframe.OLETypeAllowed = value

    @property
    def OnClick(self):
        return self.boundobjectframe.OnClick

    @OnClick.setter
    def OnClick(self, value):
        self.boundobjectframe.OnClick = value

    @property
    def OnDblClick(self):
        return self.boundobjectframe.OnDblClick

    @OnDblClick.setter
    def OnDblClick(self, value):
        self.boundobjectframe.OnDblClick = value

    @property
    def OnEnter(self):
        return self.boundobjectframe.OnEnter

    @OnEnter.setter
    def OnEnter(self, value):
        self.boundobjectframe.OnEnter = value

    @property
    def OnExit(self):
        return self.boundobjectframe.OnExit

    @OnExit.setter
    def OnExit(self, value):
        self.boundobjectframe.OnExit = value

    @property
    def OnGotFocus(self):
        return self.boundobjectframe.OnGotFocus

    @OnGotFocus.setter
    def OnGotFocus(self, value):
        self.boundobjectframe.OnGotFocus = value

    @property
    def OnKeyDown(self):
        return self.boundobjectframe.OnKeyDown

    @OnKeyDown.setter
    def OnKeyDown(self, value):
        self.boundobjectframe.OnKeyDown = value

    @property
    def OnKeyPress(self):
        return self.boundobjectframe.OnKeyPress

    @OnKeyPress.setter
    def OnKeyPress(self, value):
        self.boundobjectframe.OnKeyPress = value

    @property
    def OnKeyUp(self):
        return self.boundobjectframe.OnKeyUp

    @OnKeyUp.setter
    def OnKeyUp(self, value):
        self.boundobjectframe.OnKeyUp = value

    @property
    def OnLostFocus(self):
        return self.boundobjectframe.OnLostFocus

    @OnLostFocus.setter
    def OnLostFocus(self, value):
        self.boundobjectframe.OnLostFocus = value

    @property
    def OnMouseDown(self):
        return self.boundobjectframe.OnMouseDown

    @OnMouseDown.setter
    def OnMouseDown(self, value):
        self.boundobjectframe.OnMouseDown = value

    @property
    def OnMouseMove(self):
        return self.boundobjectframe.OnMouseMove

    @OnMouseMove.setter
    def OnMouseMove(self, value):
        self.boundobjectframe.OnMouseMove = value

    @property
    def OnMouseUp(self):
        return self.boundobjectframe.OnMouseUp

    @OnMouseUp.setter
    def OnMouseUp(self, value):
        self.boundobjectframe.OnMouseUp = value

    @property
    def OnUpdated(self):
        return self.boundobjectframe.OnUpdated

    @OnUpdated.setter
    def OnUpdated(self, value):
        self.boundobjectframe.OnUpdated = value

    @property
    def Parent(self):
        return self.boundobjectframe.Parent

    @property
    def Properties(self):
        return Properties(self.boundobjectframe.Properties)

    @property
    def RightPadding(self):
        return self.boundobjectframe.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.boundobjectframe.RightPadding = value

    @property
    def Scaling(self):
        return self.boundobjectframe.Scaling

    @Scaling.setter
    def Scaling(self, value):
        self.boundobjectframe.Scaling = value

    @property
    def Section(self):
        return self.boundobjectframe.Section

    @Section.setter
    def Section(self, value):
        self.boundobjectframe.Section = value

    @property
    def ShortcutMenuBar(self):
        return self.boundobjectframe.ShortcutMenuBar

    @ShortcutMenuBar.setter
    def ShortcutMenuBar(self, value):
        self.boundobjectframe.ShortcutMenuBar = value

    @property
    def SizeMode(self):
        return self.boundobjectframe.SizeMode

    @property
    def SourceDoc(self):
        return self.boundobjectframe.SourceDoc

    @SourceDoc.setter
    def SourceDoc(self, value):
        self.boundobjectframe.SourceDoc = value

    @property
    def SourceItem(self):
        return self.boundobjectframe.SourceItem

    @SourceItem.setter
    def SourceItem(self, value):
        self.boundobjectframe.SourceItem = value

    @property
    def SpecialEffect(self):
        return self.boundobjectframe.SpecialEffect

    @SpecialEffect.setter
    def SpecialEffect(self, value):
        self.boundobjectframe.SpecialEffect = value

    @property
    def StatusBarText(self):
        return self.boundobjectframe.StatusBarText

    @StatusBarText.setter
    def StatusBarText(self, value):
        self.boundobjectframe.StatusBarText = value

    @property
    def TabIndex(self):
        return self.boundobjectframe.TabIndex

    @TabIndex.setter
    def TabIndex(self, value):
        self.boundobjectframe.TabIndex = value

    @property
    def TabStop(self):
        return self.boundobjectframe.TabStop

    @TabStop.setter
    def TabStop(self, value):
        self.boundobjectframe.TabStop = value

    @property
    def Tag(self):
        return self.boundobjectframe.Tag

    @Tag.setter
    def Tag(self, value):
        self.boundobjectframe.Tag = value

    @property
    def Top(self):
        return self.boundobjectframe.Top

    @Top.setter
    def Top(self, value):
        self.boundobjectframe.Top = value

    @property
    def TopPadding(self):
        return self.boundobjectframe.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.boundobjectframe.TopPadding = value

    @property
    def UpdateOptions(self):
        return self.boundobjectframe.UpdateOptions

    @UpdateOptions.setter
    def UpdateOptions(self, value):
        self.boundobjectframe.UpdateOptions = value

    @property
    def Value(self):
        return self.boundobjectframe.Value

    @Value.setter
    def Value(self, value):
        self.boundobjectframe.Value = value

    @property
    def VarOleObject(self):
        return self.boundobjectframe.VarOleObject

    @property
    def Verb(self):
        return self.boundobjectframe.Verb

    @Verb.setter
    def Verb(self, value):
        self.boundobjectframe.Verb = value

    @property
    def VerticalAnchor(self):
        return self.boundobjectframe.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.boundobjectframe.VerticalAnchor = value

    @property
    def Visible(self):
        return self.boundobjectframe.Visible

    @Visible.setter
    def Visible(self, value):
        self.boundobjectframe.Visible = value

    @property
    def Width(self):
        return self.boundobjectframe.Width

    @Width.setter
    def Width(self, value):
        self.boundobjectframe.Width = value

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.boundobjectframe.Move(*args, **arguments)

    def Requery(self):
        self.boundobjectframe.Requery()

    def SetFocus(self):
        return self.boundobjectframe.SetFocus()

    def SizeToFit(self):
        self.boundobjectframe.SizeToFit()

class Chart:

    def __init__(self, chart=None):
        self.chart = chart

    @property
    def CategoryAxisFontColor(self):
        return self.chart.CategoryAxisFontColor

    @CategoryAxisFontColor.setter
    def CategoryAxisFontColor(self, value):
        self.chart.CategoryAxisFontColor = value

    @property
    def CategoryAxisFontSize(self):
        return self.chart.CategoryAxisFontSize

    @CategoryAxisFontSize.setter
    def CategoryAxisFontSize(self, value):
        self.chart.CategoryAxisFontSize = value

    @property
    def CategoryAxisTitle(self):
        return self.chart.CategoryAxisTitle

    @CategoryAxisTitle.setter
    def CategoryAxisTitle(self, value):
        self.chart.CategoryAxisTitle = value

    @property
    def ChartAxis(self):
        return self.chart.ChartAxis

    @ChartAxis.setter
    def ChartAxis(self, value):
        self.chart.ChartAxis = value

    @property
    def ChartAxisCollection(self):
        return self.chart.ChartAxisCollection

    @property
    def ChartLegend(self):
        return self.chart.ChartLegend

    @ChartLegend.setter
    def ChartLegend(self, value):
        self.chart.ChartLegend = value

    @property
    def ChartSeriesCollection(self):
        return self.chart.ChartSeriesCollection

    @property
    def ChartSubtitle(self):
        return self.chart.ChartSubtitle

    @ChartSubtitle.setter
    def ChartSubtitle(self, value):
        self.chart.ChartSubtitle = value

    @property
    def ChartSubtitleFontColor(self):
        return self.chart.ChartSubtitleFontColor

    @ChartSubtitleFontColor.setter
    def ChartSubtitleFontColor(self, value):
        self.chart.ChartSubtitleFontColor = value

    @property
    def ChartSubtitleFontSize(self):
        return self.chart.ChartSubtitleFontSize

    @ChartSubtitleFontSize.setter
    def ChartSubtitleFontSize(self, value):
        self.chart.ChartSubtitleFontSize = value

    @property
    def Title(self):
        return self.chart.Title

    @Title.setter
    def Title(self, value):
        self.chart.Title = value

    @property
    def ChartTitleFontColor(self):
        return self.chart.ChartTitleFontColor

    @ChartTitleFontColor.setter
    def ChartTitleFontColor(self, value):
        self.chart.ChartTitleFontColor = value

    @property
    def ChartTitleFontName(self):
        return self.chart.ChartTitleFontName

    @ChartTitleFontName.setter
    def ChartTitleFontName(self, value):
        self.chart.ChartTitleFontName = value

    @property
    def ChartTitleFontSize(self):
        return self.chart.ChartTitleFontSize

    @ChartTitleFontSize.setter
    def ChartTitleFontSize(self, value):
        self.chart.ChartTitleFontSize = value

    @property
    def ChartType(self):
        return AcChartType(self.chart.ChartType)

    @ChartType.setter
    def ChartType(self, value):
        self.chart.ChartType = value

    @property
    def ChartValues(self):
        return self.chart.ChartValues

    @ChartValues.setter
    def ChartValues(self, value):
        self.chart.ChartValues = value

    @property
    def ChartValuesCollection(self):
        return self.chart.ChartValuesCollection

    @property
    def HasAxisTitles(self):
        return self.chart.HasAxisTitles

    @HasAxisTitles.setter
    def HasAxisTitles(self, value):
        self.chart.HasAxisTitles = value

    @property
    def HasLegend(self):
        return self.chart.HasLegend

    @HasLegend.setter
    def HasLegend(self, value):
        self.chart.HasLegend = value

    @property
    def HasSubtitle(self):
        return self.chart.HasSubtitle

    @HasSubtitle.setter
    def HasSubtitle(self, value):
        self.chart.HasSubtitle = value

    @property
    def HasTitle(self):
        return self.chart.HasTitle

    @HasTitle.setter
    def HasTitle(self, value):
        self.chart.HasTitle = value

    @property
    def LegendPosition(self):
        return AcLegendPosition(self.chart.LegendPosition)

    @LegendPosition.setter
    def LegendPosition(self, value):
        self.chart.LegendPosition = value

    @property
    def LegendTextFontColor(self):
        return self.chart.LegendTextFontColor

    @LegendTextFontColor.setter
    def LegendTextFontColor(self, value):
        self.chart.LegendTextFontColor = value

    @property
    def LegendTextFontSize(self):
        return self.chart.LegendTextFontSize

    @LegendTextFontSize.setter
    def LegendTextFontSize(self, value):
        self.chart.LegendTextFontSize = value

    @property
    def PrimaryValuesAxisDisplayUnits(self):
        return AcAxisUnits(self.chart.PrimaryValuesAxisDisplayUnits)

    @PrimaryValuesAxisDisplayUnits.setter
    def PrimaryValuesAxisDisplayUnits(self, value):
        self.chart.PrimaryValuesAxisDisplayUnits = value

    @property
    def PrimaryValuesAxisFontColor(self):
        return self.chart.PrimaryValuesAxisFontColor

    @PrimaryValuesAxisFontColor.setter
    def PrimaryValuesAxisFontColor(self, value):
        self.chart.PrimaryValuesAxisFontColor = value

    @property
    def PrimaryValuesAxisFontSize(self):
        return self.chart.PrimaryValuesAxisFontSize

    @PrimaryValuesAxisFontSize.setter
    def PrimaryValuesAxisFontSize(self, value):
        self.chart.PrimaryValuesAxisFontSize = value

    @property
    def PrimaryValuesAxisFormat(self):
        return self.chart.PrimaryValuesAxisFormat

    @PrimaryValuesAxisFormat.setter
    def PrimaryValuesAxisFormat(self, value):
        self.chart.PrimaryValuesAxisFormat = value

    @property
    def PrimaryValuesAxisMaximum(self):
        return self.chart.PrimaryValuesAxisMaximum

    @PrimaryValuesAxisMaximum.setter
    def PrimaryValuesAxisMaximum(self, value):
        self.chart.PrimaryValuesAxisMaximum = value

    @property
    def PrimaryValuesAxisMinimum(self):
        return self.chart.PrimaryValuesAxisMinimum

    @PrimaryValuesAxisMinimum.setter
    def PrimaryValuesAxisMinimum(self, value):
        self.chart.PrimaryValuesAxisMinimum = value

    @property
    def PrimaryValuesAxisRange(self):
        return AcAxisRange(self.chart.PrimaryValuesAxisRange)

    @PrimaryValuesAxisRange.setter
    def PrimaryValuesAxisRange(self, value):
        self.chart.PrimaryValuesAxisRange = value

    @property
    def CategoryAxisTitle(self):
        return self.chart.CategoryAxisTitle

    @CategoryAxisTitle.setter
    def CategoryAxisTitle(self, value):
        self.chart.CategoryAxisTitle = value

    @property
    def RowSource(self):
        return self.chart.RowSource

    @RowSource.setter
    def RowSource(self, value):
        self.chart.RowSource = value

    @property
    def SecondaryValuesAxisDisplayUnits(self):
        return AcAxisUnits(self.chart.SecondaryValuesAxisDisplayUnits)

    @SecondaryValuesAxisDisplayUnits.setter
    def SecondaryValuesAxisDisplayUnits(self, value):
        self.chart.SecondaryValuesAxisDisplayUnits = value

    @property
    def SecondaryValuesAxisFontColor(self):
        return self.chart.SecondaryValuesAxisFontColor

    @SecondaryValuesAxisFontColor.setter
    def SecondaryValuesAxisFontColor(self, value):
        self.chart.SecondaryValuesAxisFontColor = value

    @property
    def SecondaryValuesAxisFontSize(self):
        return self.chart.SecondaryValuesAxisFontSize

    @SecondaryValuesAxisFontSize.setter
    def SecondaryValuesAxisFontSize(self, value):
        self.chart.SecondaryValuesAxisFontSize = value

    @property
    def SecondaryValuesAxisFormat(self):
        return self.chart.SecondaryValuesAxisFormat

    @SecondaryValuesAxisFormat.setter
    def SecondaryValuesAxisFormat(self, value):
        self.chart.SecondaryValuesAxisFormat = value

    @property
    def SecondaryValuesAxisMaximum(self):
        return self.chart.SecondaryValuesAxisMaximum

    @SecondaryValuesAxisMaximum.setter
    def SecondaryValuesAxisMaximum(self, value):
        self.chart.SecondaryValuesAxisMaximum = value

    @property
    def SecondaryValuesAxisMinimum(self):
        return self.chart.SecondaryValuesAxisMinimum

    @SecondaryValuesAxisMinimum.setter
    def SecondaryValuesAxisMinimum(self, value):
        self.chart.SecondaryValuesAxisMinimum = value

    @property
    def SecondaryValuesAxisRange(self):
        return AcAxisRange(self.chart.SecondaryValuesAxisRange)

    @SecondaryValuesAxisRange.setter
    def SecondaryValuesAxisRange(self, value):
        self.chart.SecondaryValuesAxisRange = value

    @property
    def SecondaryValuesAxisTitle(self):
        return self.chart.SecondaryValuesAxisTitle

    @SecondaryValuesAxisTitle.setter
    def SecondaryValuesAxisTitle(self, value):
        self.chart.SecondaryValuesAxisTitle = value

    @property
    def TransformedRowSource(self):
        return self.chart.TransformedRowSource

class ChartAxis:

    def __init__(self, chartaxis=None):
        self.chartaxis = chartaxis

    @property
    def GroupType(self):
        return AcDateGroupType(self.chartaxis.GroupType)

    @GroupType.setter
    def GroupType(self, value):
        self.chartaxis.GroupType = value

    @property
    def Name(self):
        return ChartAxis(self.chartaxis.Name)

class ChartAxisCollection:

    def __init__(self, chartaxiscollection=None):
        self.chartaxiscollection = chartaxiscollection

    def __call__(self, item):
        return ChartAxisCollectio(self.chartaxiscollection(item))

class ChartSeries:

    def __init__(self, chartseries=None):
        self.chartseries = chartseries

    @property
    def BorderColor(self):
        return self.chartseries.BorderColor

    @BorderColor.setter
    def BorderColor(self, value):
        self.chartseries.BorderColor = value

    @property
    def ComboChartType(self):
        return AcChartType(self.chartseries.ComboChartType)

    @ComboChartType.setter
    def ComboChartType(self, value):
        self.chartseries.ComboChartType = value

    @property
    def DashType(self):
        return self.chartseries.DashType

    @DashType.setter
    def DashType(self, value):
        self.chartseries.DashType = value

    @property
    def DataLabelDisplayFormat(self):
        return AcDataLabelDisplayFormat(self.chartseries.DataLabelDisplayFormat)

    @DataLabelDisplayFormat.setter
    def DataLabelDisplayFormat(self, value):
        self.chartseries.DataLabelDisplayFormat = value

    @property
    def DataLabelPosition(self):
        return AcDataLabelPosition(self.chartseries.DataLabelPosition)

    @DataLabelPosition.setter
    def DataLabelPosition(self, value):
        self.chartseries.DataLabelPosition = value

    @property
    def DisplayBoxWhiskerDataPoints(self):
        return self.chartseries.DisplayBoxWhiskerDataPoints

    @DisplayBoxWhiskerDataPoints.setter
    def DisplayBoxWhiskerDataPoints(self, value):
        self.chartseries.DisplayBoxWhiskerDataPoints = value

    @property
    def DisplayBoxWhiskerMeanMarker(self):
        return self.chartseries.DisplayBoxWhiskerMeanMarker

    @DisplayBoxWhiskerMeanMarker.setter
    def DisplayBoxWhiskerMeanMarker(self, value):
        self.chartseries.DisplayBoxWhiskerMeanMarker = value

    @property
    def DisplayDataLabel(self):
        return self.chartseries.DisplayDataLabel

    @DisplayDataLabel.setter
    def DisplayDataLabel(self, value):
        self.chartseries.DisplayDataLabel = value

    @property
    def DisplayName(self):
        return self.chartseries.DisplayName

    @DisplayName.setter
    def DisplayName(self, value):
        self.chartseries.DisplayName = value

    @property
    def FillColor(self):
        return self.chartseries.FillColor

    @FillColor.setter
    def FillColor(self, value):
        self.chartseries.FillColor = value

    @property
    def GridlinesColor(self):
        return self.chartseries.GridlinesColor

    @GridlinesColor.setter
    def GridlinesColor(self, value):
        self.chartseries.GridlinesColor = value

    @property
    def GridlinesType(self):
        return AcGridlineType(self.chartseries.GridlinesType)

    @GridlinesType.setter
    def GridlinesType(self, value):
        self.chartseries.GridlinesType = value

    @property
    def LineWeight(self):
        return self.chartseries.LineWeight

    @LineWeight.setter
    def LineWeight(self, value):
        self.chartseries.LineWeight = value

    @property
    def MarkerType(self):
        return AcMarkerType(self.chartseries.MarkerType)

    @MarkerType.setter
    def MarkerType(self, value):
        self.chartseries.MarkerType = value

    @property
    def MissingDataPolicy(self):
        return AcMissingDataPolicy(self.chartseries.MissingDataPolicy)

    @MissingDataPolicy.setter
    def MissingDataPolicy(self, value):
        self.chartseries.MissingDataPolicy = value

    @property
    def Name(self):
        return ChartSeries(self.chartseries.Name)

    @property
    def ParetoLineColor(self):
        return self.chartseries.ParetoLineColor

    @ParetoLineColor.setter
    def ParetoLineColor(self, value):
        self.chartseries.ParetoLineColor = value

    @property
    def PercentageDataLabelDecimalPlaces(self):
        return AcPercentageDataLabelDecimalPlaces(self.chartseries.PercentageDataLabelDecimalPlaces)

    @PercentageDataLabelDecimalPlaces.setter
    def PercentageDataLabelDecimalPlaces(self, value):
        self.chartseries.PercentageDataLabelDecimalPlaces = value

    @property
    def PlotSeriesOn(self):
        return AcValueAxis(self.chartseries.PlotSeriesOn)

    @PlotSeriesOn.setter
    def PlotSeriesOn(self, value):
        self.chartseries.PlotSeriesOn = value

    @property
    def ShowFunnelPercentages(self):
        return self.chartseries.ShowFunnelPercentages

    @ShowFunnelPercentages.setter
    def ShowFunnelPercentages(self, value):
        self.chartseries.ShowFunnelPercentages = value

    @property
    def ShowWaterfallConnectorLines(self):
        return self.chartseries.ShowWaterfallConnectorLines

    @ShowWaterfallConnectorLines.setter
    def ShowWaterfallConnectorLines(self, value):
        self.chartseries.ShowWaterfallConnectorLines = value

    @property
    def ShowWaterfallTotal(self):
        return self.chartseries.ShowWaterfallTotal

    @ShowWaterfallTotal.setter
    def ShowWaterfallTotal(self, value):
        self.chartseries.ShowWaterfallTotal = value

    @property
    def SortOrderType(self):
        return AcSortOrderType(self.chartseries.SortOrderType)

    @SortOrderType.setter
    def SortOrderType(self, value):
        self.chartseries.SortOrderType = value

    @property
    def TrendlineName(self):
        return self.chartseries.TrendlineName

    @TrendlineName.setter
    def TrendlineName(self, value):
        self.chartseries.TrendlineName = value

    @property
    def TrendlineOptions(self):
        return AcTrendlineOptions(self.chartseries.TrendlineOptions)

    @TrendlineOptions.setter
    def TrendlineOptions(self, value):
        self.chartseries.TrendlineOptions = value

    @property
    def WordCloudShape(self):
        return AcWordCloudShape(self.chartseries.WordCloudShape)

    @WordCloudShape.setter
    def WordCloudShape(self, value):
        self.chartseries.WordCloudShape = value

    @property
    def WordCloudWordOrientation(self):
        return AcWordCloudWordOrientation(self.chartseries.WordCloudWordOrientation)

    @WordCloudWordOrientation.setter
    def WordCloudWordOrientation(self, value):
        self.chartseries.WordCloudWordOrientation = value

class ChartSeriesCollection:

    def __init__(self, chartseriescollection=None):
        self.chartseriescollection = chartseriescollection

    def __call__(self, item):
        return ChartSeriesCollectio(self.chartseriescollection(item))

class ChartValues:

    def __init__(self, chartvalues=None):
        self.chartvalues = chartvalues

    @property
    def AggregateType(self):
        return AcAggregateType(self.chartvalues.AggregateType)

    @AggregateType.setter
    def AggregateType(self, value):
        self.chartvalues.AggregateType = value

    @property
    def Name(self):
        return ChartValues(self.chartvalues.Name)

class ChartValuesCollection:

    def __init__(self, chartvaluescollection=None):
        self.chartvaluescollection = chartvaluescollection

    def __call__(self, item):
        return ChartValuesCollectio(self.chartvaluescollection(item))

class CheckBox:

    def __init__(self, checkbox=None):
        self.checkbox = checkbox

    @property
    def AddColon(self):
        return self.checkbox.AddColon

    @AddColon.setter
    def AddColon(self, value):
        self.checkbox.AddColon = value

    @property
    def Application(self):
        return self.checkbox.Application

    @property
    def AutoLabel(self):
        return self.checkbox.AutoLabel

    @AutoLabel.setter
    def AutoLabel(self, value):
        self.checkbox.AutoLabel = value

    @property
    def BorderColor(self):
        return self.checkbox.BorderColor

    @BorderColor.setter
    def BorderColor(self, value):
        self.checkbox.BorderColor = value

    @property
    def BorderShade(self):
        return self.checkbox.BorderShade

    @BorderShade.setter
    def BorderShade(self, value):
        self.checkbox.BorderShade = value

    @property
    def BorderStyle(self):
        return self.checkbox.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.checkbox.BorderStyle = value

    @property
    def BorderThemeColorIndex(self):
        return self.checkbox.BorderThemeColorIndex

    @BorderThemeColorIndex.setter
    def BorderThemeColorIndex(self, value):
        self.checkbox.BorderThemeColorIndex = value

    @property
    def BorderTint(self):
        return self.checkbox.BorderTint

    @BorderTint.setter
    def BorderTint(self, value):
        self.checkbox.BorderTint = value

    @property
    def BorderWidth(self):
        return self.checkbox.BorderWidth

    @BorderWidth.setter
    def BorderWidth(self, value):
        self.checkbox.BorderWidth = value

    @property
    def BottomPadding(self):
        return self.checkbox.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.checkbox.BottomPadding = value

    @property
    def ColumnHidden(self):
        return self.checkbox.ColumnHidden

    @ColumnHidden.setter
    def ColumnHidden(self, value):
        self.checkbox.ColumnHidden = value

    @property
    def ColumnOrder(self):
        return self.checkbox.ColumnOrder

    @ColumnOrder.setter
    def ColumnOrder(self, value):
        self.checkbox.ColumnOrder = value

    @property
    def ColumnWidth(self):
        return self.checkbox.ColumnWidth

    @ColumnWidth.setter
    def ColumnWidth(self, value):
        self.checkbox.ColumnWidth = value

    @property
    def Controls(self):
        return Controls(self.checkbox.Controls)

    @property
    def ControlSource(self):
        return self.checkbox.ControlSource

    @ControlSource.setter
    def ControlSource(self, value):
        self.checkbox.ControlSource = value

    @property
    def ControlTipText(self):
        return self.checkbox.ControlTipText

    @ControlTipText.setter
    def ControlTipText(self, value):
        self.checkbox.ControlTipText = value

    @property
    def ControlType(self):
        return self.checkbox.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.checkbox.ControlType = value

    @property
    def DefaultValue(self):
        return self.checkbox.DefaultValue

    @DefaultValue.setter
    def DefaultValue(self, value):
        self.checkbox.DefaultValue = value

    @property
    def DisplayWhen(self):
        return self.checkbox.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.checkbox.DisplayWhen = value

    @property
    def Enabled(self):
        return self.checkbox.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.checkbox.Enabled = value

    @property
    def EventProcPrefix(self):
        return self.checkbox.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.checkbox.EventProcPrefix = value

    @property
    def GridlineColor(self):
        return self.checkbox.GridlineColor

    @GridlineColor.setter
    def GridlineColor(self, value):
        self.checkbox.GridlineColor = value

    @property
    def GridlineShade(self):
        return self.checkbox.GridlineShade

    @GridlineShade.setter
    def GridlineShade(self, value):
        self.checkbox.GridlineShade = value

    @property
    def GridlineStyleBottom(self):
        return self.checkbox.GridlineStyleBottom

    @property
    def GridlineStyleLeft(self):
        return self.checkbox.GridlineStyleLeft

    @GridlineStyleLeft.setter
    def GridlineStyleLeft(self, value):
        self.checkbox.GridlineStyleLeft = value

    @property
    def GridlineStyleRight(self):
        return self.checkbox.GridlineStyleRight

    @GridlineStyleRight.setter
    def GridlineStyleRight(self, value):
        self.checkbox.GridlineStyleRight = value

    @property
    def GridlineStyleTop(self):
        return self.checkbox.GridlineStyleTop

    @GridlineStyleTop.setter
    def GridlineStyleTop(self, value):
        self.checkbox.GridlineStyleTop = value

    @property
    def GridlineThemeColorIndex(self):
        return self.checkbox.GridlineThemeColorIndex

    @GridlineThemeColorIndex.setter
    def GridlineThemeColorIndex(self, value):
        self.checkbox.GridlineThemeColorIndex = value

    @property
    def GridlineTint(self):
        return self.checkbox.GridlineTint

    @GridlineTint.setter
    def GridlineTint(self, value):
        self.checkbox.GridlineTint = value

    @property
    def GridlineWidthBottom(self):
        return self.checkbox.GridlineWidthBottom

    @GridlineWidthBottom.setter
    def GridlineWidthBottom(self, value):
        self.checkbox.GridlineWidthBottom = value

    @property
    def GridlineWidthLeft(self):
        return self.checkbox.GridlineWidthLeft

    @GridlineWidthLeft.setter
    def GridlineWidthLeft(self, value):
        self.checkbox.GridlineWidthLeft = value

    @property
    def GridlineWidthRight(self):
        return self.checkbox.GridlineWidthRight

    @GridlineWidthRight.setter
    def GridlineWidthRight(self, value):
        self.checkbox.GridlineWidthRight = value

    @property
    def GridlineWidthTop(self):
        return self.checkbox.GridlineWidthTop

    @GridlineWidthTop.setter
    def GridlineWidthTop(self, value):
        self.checkbox.GridlineWidthTop = value

    @property
    def Height(self):
        return self.checkbox.Height

    @Height.setter
    def Height(self, value):
        self.checkbox.Height = value

    @property
    def HelpContextId(self):
        return self.checkbox.HelpContextId

    @HelpContextId.setter
    def HelpContextId(self, value):
        self.checkbox.HelpContextId = value

    @property
    def HideDuplicates(self):
        return self.checkbox.HideDuplicates

    @HideDuplicates.setter
    def HideDuplicates(self, value):
        self.checkbox.HideDuplicates = value

    @property
    def HorizontalAnchor(self):
        return self.checkbox.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.checkbox.HorizontalAnchor = value

    @property
    def InSelection(self):
        return self.checkbox.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.checkbox.InSelection = value

    @property
    def IsVisible(self):
        return self.checkbox.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.checkbox.IsVisible = value

    @property
    def LabelAlign(self):
        return self.checkbox.LabelAlign

    @LabelAlign.setter
    def LabelAlign(self, value):
        self.checkbox.LabelAlign = value

    @property
    def LabelX(self):
        return self.checkbox.LabelX

    @LabelX.setter
    def LabelX(self, value):
        self.checkbox.LabelX = value

    @property
    def LabelY(self):
        return self.checkbox.LabelY

    @LabelY.setter
    def LabelY(self, value):
        self.checkbox.LabelY = value

    @property
    def Layout(self):
        return AcLayoutType(self.checkbox.Layout)

    @property
    def LayoutID(self):
        return self.checkbox.LayoutID

    @property
    def Left(self):
        return self.checkbox.Left

    @Left.setter
    def Left(self, value):
        self.checkbox.Left = value

    @property
    def LeftPadding(self):
        return self.checkbox.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.checkbox.LeftPadding = value

    @property
    def Locked(self):
        return self.checkbox.Locked

    @Locked.setter
    def Locked(self, value):
        self.checkbox.Locked = value

    @property
    def Name(self):
        return self.checkbox.Name

    @Name.setter
    def Name(self, value):
        self.checkbox.Name = value

    @property
    def OldBorderStyle(self):
        return self.checkbox.OldBorderStyle

    @OldBorderStyle.setter
    def OldBorderStyle(self, value):
        self.checkbox.OldBorderStyle = value

    @property
    def OldValue(self):
        return self.checkbox.OldValue

    @property
    def OnClick(self):
        return self.checkbox.OnClick

    @OnClick.setter
    def OnClick(self, value):
        self.checkbox.OnClick = value

    @property
    def OnDblClick(self):
        return self.checkbox.OnDblClick

    @OnDblClick.setter
    def OnDblClick(self, value):
        self.checkbox.OnDblClick = value

    @property
    def OnEnter(self):
        return self.checkbox.OnEnter

    @OnEnter.setter
    def OnEnter(self, value):
        self.checkbox.OnEnter = value

    @property
    def OnExit(self):
        return self.checkbox.OnExit

    @OnExit.setter
    def OnExit(self, value):
        self.checkbox.OnExit = value

    @property
    def OnGotFocus(self):
        return self.checkbox.OnGotFocus

    @OnGotFocus.setter
    def OnGotFocus(self, value):
        self.checkbox.OnGotFocus = value

    @property
    def OnKeyDown(self):
        return self.checkbox.OnKeyDown

    @OnKeyDown.setter
    def OnKeyDown(self, value):
        self.checkbox.OnKeyDown = value

    @property
    def OnKeyPress(self):
        return self.checkbox.OnKeyPress

    @OnKeyPress.setter
    def OnKeyPress(self, value):
        self.checkbox.OnKeyPress = value

    @property
    def OnKeyUp(self):
        return self.checkbox.OnKeyUp

    @OnKeyUp.setter
    def OnKeyUp(self, value):
        self.checkbox.OnKeyUp = value

    @property
    def OnLostFocus(self):
        return self.checkbox.OnLostFocus

    @OnLostFocus.setter
    def OnLostFocus(self, value):
        self.checkbox.OnLostFocus = value

    @property
    def OnMouseDown(self):
        return self.checkbox.OnMouseDown

    @OnMouseDown.setter
    def OnMouseDown(self, value):
        self.checkbox.OnMouseDown = value

    @property
    def OnMouseMove(self):
        return self.checkbox.OnMouseMove

    @OnMouseMove.setter
    def OnMouseMove(self, value):
        self.checkbox.OnMouseMove = value

    @property
    def OnMouseUp(self):
        return self.checkbox.OnMouseUp

    @OnMouseUp.setter
    def OnMouseUp(self, value):
        self.checkbox.OnMouseUp = value

    @property
    def OptionValue(self):
        return self.checkbox.OptionValue

    @OptionValue.setter
    def OptionValue(self, value):
        self.checkbox.OptionValue = value

    @property
    def Parent(self):
        return self.checkbox.Parent

    @property
    def Properties(self):
        return Properties(self.checkbox.Properties)

    @property
    def ReadingOrder(self):
        return self.checkbox.ReadingOrder

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.checkbox.ReadingOrder = value

    @property
    def RightPadding(self):
        return self.checkbox.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.checkbox.RightPadding = value

    @property
    def Section(self):
        return self.checkbox.Section

    @Section.setter
    def Section(self, value):
        self.checkbox.Section = value

    @property
    def ShortcutMenuBar(self):
        return self.checkbox.ShortcutMenuBar

    @ShortcutMenuBar.setter
    def ShortcutMenuBar(self, value):
        self.checkbox.ShortcutMenuBar = value

    @property
    def SpecialEffect(self):
        return self.checkbox.SpecialEffect

    @SpecialEffect.setter
    def SpecialEffect(self, value):
        self.checkbox.SpecialEffect = value

    @property
    def StatusBarText(self):
        return self.checkbox.StatusBarText

    @StatusBarText.setter
    def StatusBarText(self, value):
        self.checkbox.StatusBarText = value

    @property
    def TabIndex(self):
        return self.checkbox.TabIndex

    @TabIndex.setter
    def TabIndex(self, value):
        self.checkbox.TabIndex = value

    @property
    def TabStop(self):
        return self.checkbox.TabStop

    @TabStop.setter
    def TabStop(self, value):
        self.checkbox.TabStop = value

    @property
    def Tag(self):
        return self.checkbox.Tag

    @Tag.setter
    def Tag(self, value):
        self.checkbox.Tag = value

    @property
    def Top(self):
        return self.checkbox.Top

    @Top.setter
    def Top(self, value):
        self.checkbox.Top = value

    @property
    def TopPadding(self):
        return self.checkbox.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.checkbox.TopPadding = value

    @property
    def TripleState(self):
        return self.checkbox.TripleState

    @TripleState.setter
    def TripleState(self, value):
        self.checkbox.TripleState = value

    @property
    def ValidationRule(self):
        return self.checkbox.ValidationRule

    @ValidationRule.setter
    def ValidationRule(self, value):
        self.checkbox.ValidationRule = value

    @property
    def ValidationText(self):
        return self.checkbox.ValidationText

    @ValidationText.setter
    def ValidationText(self, value):
        self.checkbox.ValidationText = value

    @property
    def Value(self):
        return self.checkbox.Value

    @Value.setter
    def Value(self, value):
        self.checkbox.Value = value

    @property
    def VerticalAnchor(self):
        return self.checkbox.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.checkbox.VerticalAnchor = value

    @property
    def Visible(self):
        return self.checkbox.Visible

    @Visible.setter
    def Visible(self, value):
        self.checkbox.Visible = value

    @property
    def Width(self):
        return self.checkbox.Width

    @Width.setter
    def Width(self, value):
        self.checkbox.Width = value

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.checkbox.Move(*args, **arguments)

    def Requery(self):
        self.checkbox.Requery()

    def SetFocus(self):
        return self.checkbox.SetFocus()

    def SizeToFit(self):
        self.checkbox.SizeToFit()

    def Undo(self):
        self.checkbox.Undo()

class CodeData:

    def __init__(self, codedata=None):
        self.codedata = codedata

    @property
    def AllDatabaseDiagrams(self):
        return self.codedata.AllDatabaseDiagrams

    @property
    def AllFunctions(self):
        return self.codedata.AllFunctions

    @property
    def AllQueries(self):
        return self.codedata.AllQueries

    @property
    def AllStoredProcedures(self):
        return self.codedata.AllStoredProcedures

    @property
    def AllTables(self):
        return self.codedata.AllTables

    @property
    def AllViews(self):
        return self.codedata.AllViews

class CodeProject:

    def __init__(self, codeproject=None):
        self.codeproject = codeproject

    @property
    def AccessConnection(self):
        return self.codeproject.AccessConnection

    @property
    def AllForms(self):
        return self.codeproject.AllForms

    @property
    def AllMacros(self):
        return self.codeproject.AllMacros

    @property
    def AllModules(self):
        return self.codeproject.AllModules

    @property
    def AllReports(self):
        return self.codeproject.AllReports

    @property
    def Application(self):
        return self.codeproject.Application

    @property
    def BaseConnectionString(self):
        return self.codeproject.BaseConnectionString

    @property
    def Connection(self):
        return self.codeproject.Connection

    @property
    def FileFormat(self):
        return AcFileFormat(self.codeproject.FileFormat)

    @property
    def FullName(self):
        return self.codeproject.FullName

    @property
    def ImportExportSpecifications(self):
        return ImportExportSpecifications(self.codeproject.ImportExportSpecifications)

    @property
    def IsConnected(self):
        return self.codeproject.IsConnected

    @property
    def IsTrusted(self):
        return self.codeproject.IsTrusted

    @property
    def IsWeb(self):
        return self.codeproject.IsWeb

    @property
    def Name(self):
        return self.codeproject.Name

    @property
    def Parent(self):
        return self.codeproject.Parent

    @property
    def Path(self):
        return self.codeproject.Path

    @property
    def ProjectType(self):
        return self.codeproject.ProjectType

    @property
    def Properties(self):
        return CodeProject(self.codeproject.Properties)

    @property
    def RemovePersonalInformation(self):
        return self.codeproject.RemovePersonalInformation

    @RemovePersonalInformation.setter
    def RemovePersonalInformation(self, value):
        self.codeproject.RemovePersonalInformation = value

    @property
    def Resources(self):
        return self.codeproject.Resources

    @property
    def WebSite(self):
        return self.codeproject.WebSite

    def AddSharedImage(self, *args, SharedImageName=None, FileName=None):
        arguments = {"SharedImageName": SharedImageName, "FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.codeproject.AddSharedImage(*args, **arguments)

    def CloseConnection(self):
        return self.codeproject.CloseConnection()

    def OpenConnection(self, *args, BaseConnectionString=None, UserID=None, Password=None):
        arguments = {"BaseConnectionString": BaseConnectionString, "UserID": UserID, "Password": Password}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.codeproject.OpenConnection(*args, **arguments)

    def UpdateDependencyInfo(self):
        return self.codeproject.UpdateDependencyInfo()

class ComboBox:

    def __init__(self, combobox=None):
        self.combobox = combobox

class ComboBox:

    def __init__(self, combobox=None):
        self.combobox = combobox

    @property
    def AddColon(self):
        return self.combobox.AddColon

    @AddColon.setter
    def AddColon(self, value):
        self.combobox.AddColon = value

    @property
    def AllowAutoCorrect(self):
        return self.combobox.AllowAutoCorrect

    @AllowAutoCorrect.setter
    def AllowAutoCorrect(self, value):
        self.combobox.AllowAutoCorrect = value

    @property
    def AllowValueListEdits(self):
        return self.combobox.AllowValueListEdits

    @AllowValueListEdits.setter
    def AllowValueListEdits(self, value):
        self.combobox.AllowValueListEdits = value

    @property
    def Application(self):
        return self.combobox.Application

    @property
    def AutoExpand(self):
        return self.combobox.AutoExpand

    @AutoExpand.setter
    def AutoExpand(self, value):
        self.combobox.AutoExpand = value

    @property
    def AutoLabel(self):
        return self.combobox.AutoLabel

    @AutoLabel.setter
    def AutoLabel(self, value):
        self.combobox.AutoLabel = value

    @property
    def BackColor(self):
        return self.combobox.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.combobox.BackColor = value

    @property
    def BackShade(self):
        return self.combobox.BackShade

    @BackShade.setter
    def BackShade(self, value):
        self.combobox.BackShade = value

    @property
    def BackStyle(self):
        return self.combobox.BackStyle

    @BackStyle.setter
    def BackStyle(self, value):
        self.combobox.BackStyle = value

    @property
    def BackThemeColorIndex(self):
        return self.combobox.BackThemeColorIndex

    @BackThemeColorIndex.setter
    def BackThemeColorIndex(self, value):
        self.combobox.BackThemeColorIndex = value

    @property
    def BackTint(self):
        return self.combobox.BackTint

    @BackTint.setter
    def BackTint(self, value):
        self.combobox.BackTint = value

    @property
    def BorderColor(self):
        return self.combobox.BorderColor

    @BorderColor.setter
    def BorderColor(self, value):
        self.combobox.BorderColor = value

    @property
    def BorderShade(self):
        return self.combobox.BorderShade

    @BorderShade.setter
    def BorderShade(self, value):
        self.combobox.BorderShade = value

    @property
    def BorderStyle(self):
        return self.combobox.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.combobox.BorderStyle = value

    @property
    def BorderThemeColorIndex(self):
        return self.combobox.BorderThemeColorIndex

    @BorderThemeColorIndex.setter
    def BorderThemeColorIndex(self, value):
        self.combobox.BorderThemeColorIndex = value

    @property
    def BorderTint(self):
        return self.combobox.BorderTint

    @BorderTint.setter
    def BorderTint(self, value):
        self.combobox.BorderTint = value

    @property
    def BorderWidth(self):
        return self.combobox.BorderWidth

    @BorderWidth.setter
    def BorderWidth(self, value):
        self.combobox.BorderWidth = value

    @property
    def BottomMargin(self):
        return self.combobox.BottomMargin

    @BottomMargin.setter
    def BottomMargin(self, value):
        self.combobox.BottomMargin = value

    @property
    def BottomPadding(self):
        return self.combobox.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.combobox.BottomPadding = value

    @property
    def BoundColumn(self):
        return self.combobox.BoundColumn

    @BoundColumn.setter
    def BoundColumn(self, value):
        self.combobox.BoundColumn = value

    @property
    def CanGrow(self):
        return self.combobox.CanGrow

    @CanGrow.setter
    def CanGrow(self, value):
        self.combobox.CanGrow = value

    @property
    def CanShrink(self):
        return self.combobox.CanShrink

    @CanShrink.setter
    def CanShrink(self, value):
        self.combobox.CanShrink = value

    def Column(self, *args, Index=None, Row=None):
        arguments = {"Index": Index, "Row": Row}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.combobox.Column):
            return self.combobox.Column(*args, **arguments)
        else:
            return self.combobox.GetColumn(*args, **arguments)

    @property
    def ColumnCount(self):
        return self.combobox.ColumnCount

    @ColumnCount.setter
    def ColumnCount(self, value):
        self.combobox.ColumnCount = value

    @property
    def ColumnHeads(self):
        return self.combobox.ColumnHeads

    @ColumnHeads.setter
    def ColumnHeads(self, value):
        self.combobox.ColumnHeads = value

    @property
    def ColumnHidden(self):
        return self.combobox.ColumnHidden

    @ColumnHidden.setter
    def ColumnHidden(self, value):
        self.combobox.ColumnHidden = value

    @property
    def ColumnOrder(self):
        return self.combobox.ColumnOrder

    @ColumnOrder.setter
    def ColumnOrder(self, value):
        self.combobox.ColumnOrder = value

    @property
    def ColumnWidth(self):
        return self.combobox.ColumnWidth

    @ColumnWidth.setter
    def ColumnWidth(self, value):
        self.combobox.ColumnWidth = value

    @property
    def ColumnWidths(self):
        return self.combobox.ColumnWidths

    @ColumnWidths.setter
    def ColumnWidths(self, value):
        self.combobox.ColumnWidths = value

    @property
    def Controls(self):
        return Controls(self.combobox.Controls)

    @property
    def ControlSource(self):
        return self.combobox.ControlSource

    @ControlSource.setter
    def ControlSource(self, value):
        self.combobox.ControlSource = value

    @property
    def ControlTipText(self):
        return self.combobox.ControlTipText

    @ControlTipText.setter
    def ControlTipText(self, value):
        self.combobox.ControlTipText = value

    @property
    def ControlType(self):
        return self.combobox.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.combobox.ControlType = value

    @property
    def DecimalPlaces(self):
        return self.combobox.DecimalPlaces

    @DecimalPlaces.setter
    def DecimalPlaces(self, value):
        self.combobox.DecimalPlaces = value

    @property
    def DefaultValue(self):
        return self.combobox.DefaultValue

    @DefaultValue.setter
    def DefaultValue(self, value):
        self.combobox.DefaultValue = value

    @property
    def DisplayAsHyperlink(self):
        return self.combobox.DisplayAsHyperlink

    @DisplayAsHyperlink.setter
    def DisplayAsHyperlink(self, value):
        self.combobox.DisplayAsHyperlink = value

    @property
    def DisplayWhen(self):
        return self.combobox.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.combobox.DisplayWhen = value

    @property
    def Enabled(self):
        return self.combobox.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.combobox.Enabled = value

    @property
    def EventProcPrefix(self):
        return self.combobox.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.combobox.EventProcPrefix = value

    @property
    def FontBold(self):
        return self.combobox.FontBold

    @property
    def FontItalic(self):
        return self.combobox.FontItalic

    @property
    def FontName(self):
        return self.combobox.FontName

    @property
    def FontSize(self):
        return self.combobox.FontSize

    @property
    def FontUnderline(self):
        return self.combobox.FontUnderline

    @property
    def FontWeight(self):
        return self.combobox.FontWeight

    @FontWeight.setter
    def FontWeight(self, value):
        self.combobox.FontWeight = value

    @property
    def ForeColor(self):
        return self.combobox.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.combobox.ForeColor = value

    @property
    def ForeShade(self):
        return self.combobox.ForeShade

    @ForeShade.setter
    def ForeShade(self, value):
        self.combobox.ForeShade = value

    @property
    def ForeThemeColorIndex(self):
        return self.combobox.ForeThemeColorIndex

    @ForeThemeColorIndex.setter
    def ForeThemeColorIndex(self, value):
        self.combobox.ForeThemeColorIndex = value

    @property
    def ForeTint(self):
        return self.combobox.ForeTint

    @ForeTint.setter
    def ForeTint(self, value):
        self.combobox.ForeTint = value

    @property
    def Format(self):
        return self.combobox.Format

    @Format.setter
    def Format(self, value):
        self.combobox.Format = value

    @property
    def FormatConditions(self):
        return self.combobox.FormatConditions

    @property
    def GridlineColor(self):
        return self.combobox.GridlineColor

    @GridlineColor.setter
    def GridlineColor(self, value):
        self.combobox.GridlineColor = value

    @property
    def GridlineShade(self):
        return self.combobox.GridlineShade

    @GridlineShade.setter
    def GridlineShade(self, value):
        self.combobox.GridlineShade = value

    @property
    def GridlineStyleBottom(self):
        return self.combobox.GridlineStyleBottom

    @GridlineStyleBottom.setter
    def GridlineStyleBottom(self, value):
        self.combobox.GridlineStyleBottom = value

    @property
    def GridlineStyleLeft(self):
        return self.combobox.GridlineStyleLeft

    @property
    def GridlineStyleRight(self):
        return self.combobox.GridlineStyleRight

    @GridlineStyleRight.setter
    def GridlineStyleRight(self, value):
        self.combobox.GridlineStyleRight = value

    @property
    def GridlineStyleTop(self):
        return self.combobox.GridlineStyleTop

    @GridlineStyleTop.setter
    def GridlineStyleTop(self, value):
        self.combobox.GridlineStyleTop = value

    @property
    def GridlineThemeColorIndex(self):
        return self.combobox.GridlineThemeColorIndex

    @GridlineThemeColorIndex.setter
    def GridlineThemeColorIndex(self, value):
        self.combobox.GridlineThemeColorIndex = value

    @property
    def GridlineTint(self):
        return self.combobox.GridlineTint

    @GridlineTint.setter
    def GridlineTint(self, value):
        self.combobox.GridlineTint = value

    @property
    def GridlineWidthBottom(self):
        return self.combobox.GridlineWidthBottom

    @GridlineWidthBottom.setter
    def GridlineWidthBottom(self, value):
        self.combobox.GridlineWidthBottom = value

    @property
    def GridlineWidthLeft(self):
        return self.combobox.GridlineWidthLeft

    @GridlineWidthLeft.setter
    def GridlineWidthLeft(self, value):
        self.combobox.GridlineWidthLeft = value

    @property
    def GridlineWidthRight(self):
        return self.combobox.GridlineWidthRight

    @GridlineWidthRight.setter
    def GridlineWidthRight(self, value):
        self.combobox.GridlineWidthRight = value

    @property
    def GridlineWidthTop(self):
        return self.combobox.GridlineWidthTop

    @GridlineWidthTop.setter
    def GridlineWidthTop(self, value):
        self.combobox.GridlineWidthTop = value

    @property
    def Height(self):
        return self.combobox.Height

    @Height.setter
    def Height(self, value):
        self.combobox.Height = value

    @property
    def HelpContextId(self):
        return self.combobox.HelpContextId

    @HelpContextId.setter
    def HelpContextId(self, value):
        self.combobox.HelpContextId = value

    @property
    def HideDuplicates(self):
        return self.combobox.HideDuplicates

    @HideDuplicates.setter
    def HideDuplicates(self, value):
        self.combobox.HideDuplicates = value

    @property
    def HorizontalAnchor(self):
        return self.combobox.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.combobox.HorizontalAnchor = value

    @property
    def IMEHold(self):
        return self.combobox.IMEHold

    @IMEHold.setter
    def IMEHold(self, value):
        self.combobox.IMEHold = value

    @property
    def InheritValueList(self):
        return self.combobox.InheritValueList

    @InheritValueList.setter
    def InheritValueList(self, value):
        self.combobox.InheritValueList = value

    @property
    def InputMask(self):
        return self.combobox.InputMask

    @InputMask.setter
    def InputMask(self, value):
        self.combobox.InputMask = value

    @property
    def InSelection(self):
        return self.combobox.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.combobox.InSelection = value

    @property
    def IsHyperlink(self):
        return self.combobox.IsHyperlink

    @IsHyperlink.setter
    def IsHyperlink(self, value):
        self.combobox.IsHyperlink = value

    @property
    def IsVisible(self):
        return self.combobox.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.combobox.IsVisible = value

    def ItemData(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.combobox.ItemData):
            return self.combobox.ItemData(*args, **arguments)
        else:
            return self.combobox.GetItemData(*args, **arguments)

    @property
    def ItemsSelected(self):
        return self.combobox.ItemsSelected

    @property
    def LabelAlign(self):
        return self.combobox.LabelAlign

    @LabelAlign.setter
    def LabelAlign(self, value):
        self.combobox.LabelAlign = value

    @property
    def LabelX(self):
        return self.combobox.LabelX

    @LabelX.setter
    def LabelX(self, value):
        self.combobox.LabelX = value

    @property
    def LabelY(self):
        return self.combobox.LabelY

    @LabelY.setter
    def LabelY(self, value):
        self.combobox.LabelY = value

    @property
    def Layout(self):
        return AcLayoutType(self.combobox.Layout)

    @property
    def LayoutID(self):
        return self.combobox.LayoutID

    @property
    def Left(self):
        return self.combobox.Left

    @Left.setter
    def Left(self, value):
        self.combobox.Left = value

    @property
    def LeftMargin(self):
        return self.combobox.LeftMargin

    @LeftMargin.setter
    def LeftMargin(self, value):
        self.combobox.LeftMargin = value

    @property
    def LeftPadding(self):
        return self.combobox.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.combobox.LeftPadding = value

    @property
    def LimitToList(self):
        return self.combobox.LimitToList

    @LimitToList.setter
    def LimitToList(self, value):
        self.combobox.LimitToList = value

    @property
    def ListCount(self):
        return self.combobox.ListCount

    @ListCount.setter
    def ListCount(self, value):
        self.combobox.ListCount = value

    @property
    def ListIndex(self):
        return self.combobox.ListIndex

    @property
    def ListItemsEditForm(self):
        return self.combobox.ListItemsEditForm

    @ListItemsEditForm.setter
    def ListItemsEditForm(self, value):
        self.combobox.ListItemsEditForm = value

    @property
    def ListRows(self):
        return self.combobox.ListRows

    @ListRows.setter
    def ListRows(self, value):
        self.combobox.ListRows = value

    @property
    def ListWidth(self):
        return self.combobox.ListWidth

    @ListWidth.setter
    def ListWidth(self, value):
        self.combobox.ListWidth = value

    @property
    def Locked(self):
        return self.combobox.Locked

    @Locked.setter
    def Locked(self, value):
        self.combobox.Locked = value

    @property
    def Name(self):
        return self.combobox.Name

    @Name.setter
    def Name(self, value):
        self.combobox.Name = value

    @property
    def OldBorderStyle(self):
        return self.combobox.OldBorderStyle

    @OldBorderStyle.setter
    def OldBorderStyle(self, value):
        self.combobox.OldBorderStyle = value

    @property
    def OldValue(self):
        return self.combobox.OldValue

    @property
    def OnChange(self):
        return self.combobox.OnChange

    @OnChange.setter
    def OnChange(self, value):
        self.combobox.OnChange = value

    @property
    def OnClick(self):
        return self.combobox.OnClick

    @OnClick.setter
    def OnClick(self, value):
        self.combobox.OnClick = value

    @property
    def OnDblClick(self):
        return self.combobox.OnDblClick

    @OnDblClick.setter
    def OnDblClick(self, value):
        self.combobox.OnDblClick = value

    @property
    def OnDirty(self):
        return self.combobox.OnDirty

    @OnDirty.setter
    def OnDirty(self, value):
        self.combobox.OnDirty = value

    @property
    def OnEnter(self):
        return self.combobox.OnEnter

    @OnEnter.setter
    def OnEnter(self, value):
        self.combobox.OnEnter = value

    @property
    def OnExit(self):
        return self.combobox.OnExit

    @OnExit.setter
    def OnExit(self, value):
        self.combobox.OnExit = value

    @property
    def OnGotFocus(self):
        return self.combobox.OnGotFocus

    @OnGotFocus.setter
    def OnGotFocus(self, value):
        self.combobox.OnGotFocus = value

    @property
    def OnKeyDown(self):
        return self.combobox.OnKeyDown

    @OnKeyDown.setter
    def OnKeyDown(self, value):
        self.combobox.OnKeyDown = value

    @property
    def OnKeyPress(self):
        return self.combobox.OnKeyPress

    @OnKeyPress.setter
    def OnKeyPress(self, value):
        self.combobox.OnKeyPress = value

    @property
    def OnKeyUp(self):
        return self.combobox.OnKeyUp

    @OnKeyUp.setter
    def OnKeyUp(self, value):
        self.combobox.OnKeyUp = value

    @property
    def OnLostFocus(self):
        return self.combobox.OnLostFocus

    @OnLostFocus.setter
    def OnLostFocus(self, value):
        self.combobox.OnLostFocus = value

    @property
    def OnMouseDown(self):
        return self.combobox.OnMouseDown

    @OnMouseDown.setter
    def OnMouseDown(self, value):
        self.combobox.OnMouseDown = value

    @property
    def OnMouseMove(self):
        return self.combobox.OnMouseMove

    @OnMouseMove.setter
    def OnMouseMove(self, value):
        self.combobox.OnMouseMove = value

    @property
    def OnMouseUp(self):
        return self.combobox.OnMouseUp

    @OnMouseUp.setter
    def OnMouseUp(self, value):
        self.combobox.OnMouseUp = value

    @property
    def OnNotInList(self):
        return self.combobox.OnNotInList

    @OnNotInList.setter
    def OnNotInList(self, value):
        self.combobox.OnNotInList = value

    @property
    def OnUndo(self):
        return self.combobox.OnUndo

    @OnUndo.setter
    def OnUndo(self, value):
        self.combobox.OnUndo = value

    @property
    def Parent(self):
        return self.combobox.Parent

    @property
    def Properties(self):
        return Properties(self.combobox.Properties)

    @property
    def ReadingOrder(self):
        return self.combobox.ReadingOrder

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.combobox.ReadingOrder = value

    @property
    def Recordset(self):
        return self.combobox.Recordset

    @Recordset.setter
    def Recordset(self, value):
        self.combobox.Recordset = value

    @property
    def RightMargin(self):
        return self.combobox.RightMargin

    @RightMargin.setter
    def RightMargin(self, value):
        self.combobox.RightMargin = value

    @property
    def RightPadding(self):
        return self.combobox.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.combobox.RightPadding = value

    @property
    def RowSource(self):
        return self.combobox.RowSource

    @RowSource.setter
    def RowSource(self, value):
        self.combobox.RowSource = value

    @property
    def RowSourceType(self):
        return self.combobox.RowSourceType

    @RowSourceType.setter
    def RowSourceType(self, value):
        self.combobox.RowSourceType = value

    @property
    def ScrollBarAlign(self):
        return self.combobox.ScrollBarAlign

    @ScrollBarAlign.setter
    def ScrollBarAlign(self, value):
        self.combobox.ScrollBarAlign = value

    @property
    def Section(self):
        return self.combobox.Section

    @Section.setter
    def Section(self, value):
        self.combobox.Section = value

    @property
    def Selected(self):
        return self.combobox.Selected

    @Selected.setter
    def Selected(self, value):
        self.combobox.Selected = value

    @property
    def SelLength(self):
        return self.combobox.SelLength

    @SelLength.setter
    def SelLength(self, value):
        self.combobox.SelLength = value

    @property
    def SelStart(self):
        return self.combobox.SelStart

    @SelStart.setter
    def SelStart(self, value):
        self.combobox.SelStart = value

    @property
    def SelText(self):
        return self.combobox.SelText

    @SelText.setter
    def SelText(self, value):
        self.combobox.SelText = value

    @property
    def SeparatorCharacters(self):
        return self.combobox.SeparatorCharacters

    @SeparatorCharacters.setter
    def SeparatorCharacters(self, value):
        self.combobox.SeparatorCharacters = value

    @property
    def ShortcutMenuBar(self):
        return self.combobox.ShortcutMenuBar

    @ShortcutMenuBar.setter
    def ShortcutMenuBar(self, value):
        self.combobox.ShortcutMenuBar = value

    @property
    def ShowOnlyRowSourceValues(self):
        return self.combobox.ShowOnlyRowSourceValues

    @ShowOnlyRowSourceValues.setter
    def ShowOnlyRowSourceValues(self, value):
        self.combobox.ShowOnlyRowSourceValues = value

    @property
    def SmartTags(self):
        return SmartTags(self.combobox.SmartTags)

    @property
    def SpecialEffect(self):
        return self.combobox.SpecialEffect

    @SpecialEffect.setter
    def SpecialEffect(self, value):
        self.combobox.SpecialEffect = value

    @property
    def StatusBarText(self):
        return self.combobox.StatusBarText

    @StatusBarText.setter
    def StatusBarText(self, value):
        self.combobox.StatusBarText = value

    @property
    def TabIndex(self):
        return self.combobox.TabIndex

    @TabIndex.setter
    def TabIndex(self, value):
        self.combobox.TabIndex = value

    @property
    def TabStop(self):
        return self.combobox.TabStop

    @TabStop.setter
    def TabStop(self, value):
        self.combobox.TabStop = value

    @property
    def Tag(self):
        return self.combobox.Tag

    @Tag.setter
    def Tag(self, value):
        self.combobox.Tag = value

    @property
    def Text(self):
        return self.combobox.Text

    @Text.setter
    def Text(self, value):
        self.combobox.Text = value

    @property
    def TextAlign(self):
        return self.combobox.TextAlign

    @TextAlign.setter
    def TextAlign(self, value):
        self.combobox.TextAlign = value

    @property
    def ThemeFontIndex(self):
        return self.combobox.ThemeFontIndex

    @ThemeFontIndex.setter
    def ThemeFontIndex(self, value):
        self.combobox.ThemeFontIndex = value

    @property
    def Top(self):
        return self.combobox.Top

    @Top.setter
    def Top(self, value):
        self.combobox.Top = value

    @property
    def TopMargin(self):
        return self.combobox.TopMargin

    @TopMargin.setter
    def TopMargin(self, value):
        self.combobox.TopMargin = value

    @property
    def TopPadding(self):
        return self.combobox.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.combobox.TopPadding = value

    @property
    def ValidationRule(self):
        return self.combobox.ValidationRule

    @ValidationRule.setter
    def ValidationRule(self, value):
        self.combobox.ValidationRule = value

    @property
    def ValidationText(self):
        return self.combobox.ValidationText

    @ValidationText.setter
    def ValidationText(self, value):
        self.combobox.ValidationText = value

    @property
    def Value(self):
        return self.combobox.Value

    @Value.setter
    def Value(self, value):
        self.combobox.Value = value

    @property
    def VerticalAnchor(self):
        return self.combobox.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.combobox.VerticalAnchor = value

    @property
    def Visible(self):
        return self.combobox.Visible

    @Visible.setter
    def Visible(self, value):
        self.combobox.Visible = value

    @property
    def Width(self):
        return self.combobox.Width

    @Width.setter
    def Width(self, value):
        self.combobox.Width = value

    def AddItem(self, *args, Item=None, Index=None):
        arguments = {"Item": Item, "Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.combobox.AddItem(*args, **arguments)

    def Dropdown(self):
        return self.combobox.Dropdown()

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.combobox.Move(*args, **arguments)

    def RemoveItem(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.combobox.RemoveItem(*args, **arguments)

    def Requery(self):
        self.combobox.Requery()

    def SetFocus(self):
        return self.combobox.SetFocus()

    def SizeToFit(self):
        self.combobox.SizeToFit()

    def Undo(self):
        self.combobox.Undo()

class CommandButton:

    def __init__(self, commandbutton=None):
        self.commandbutton = commandbutton

    @property
    def AddColon(self):
        return self.commandbutton.AddColon

    @AddColon.setter
    def AddColon(self, value):
        self.commandbutton.AddColon = value

    @property
    def Alignment(self):
        return self.commandbutton.Alignment

    @Alignment.setter
    def Alignment(self, value):
        self.commandbutton.Alignment = value

    @property
    def Application(self):
        return self.commandbutton.Application

    @property
    def AutoLabel(self):
        return self.commandbutton.AutoLabel

    @AutoLabel.setter
    def AutoLabel(self, value):
        self.commandbutton.AutoLabel = value

    @property
    def AutoRepeat(self):
        return self.commandbutton.AutoRepeat

    @AutoRepeat.setter
    def AutoRepeat(self, value):
        self.commandbutton.AutoRepeat = value

    @property
    def BackColor(self):
        return self.commandbutton.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.commandbutton.BackColor = value

    @property
    def BackShade(self):
        return self.commandbutton.BackShade

    @BackShade.setter
    def BackShade(self, value):
        self.commandbutton.BackShade = value

    @property
    def BackStyle(self):
        return self.commandbutton.BackStyle

    @BackStyle.setter
    def BackStyle(self, value):
        self.commandbutton.BackStyle = value

    @property
    def BackThemeColorIndex(self):
        return self.commandbutton.BackThemeColorIndex

    @BackThemeColorIndex.setter
    def BackThemeColorIndex(self, value):
        self.commandbutton.BackThemeColorIndex = value

    @property
    def BackTint(self):
        return self.commandbutton.BackTint

    @BackTint.setter
    def BackTint(self, value):
        self.commandbutton.BackTint = value

    @property
    def Bevel(self):
        return self.commandbutton.Bevel

    @Bevel.setter
    def Bevel(self, value):
        self.commandbutton.Bevel = value

    @property
    def BorderColor(self):
        return self.commandbutton.BorderColor

    @BorderColor.setter
    def BorderColor(self, value):
        self.commandbutton.BorderColor = value

    @property
    def BorderShade(self):
        return self.commandbutton.BorderShade

    @BorderShade.setter
    def BorderShade(self, value):
        self.commandbutton.BorderShade = value

    @property
    def BorderStyle(self):
        return self.commandbutton.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.commandbutton.BorderStyle = value

    @property
    def BorderThemeColorIndex(self):
        return self.commandbutton.BorderThemeColorIndex

    @BorderThemeColorIndex.setter
    def BorderThemeColorIndex(self, value):
        self.commandbutton.BorderThemeColorIndex = value

    @property
    def BorderTint(self):
        return self.commandbutton.BorderTint

    @BorderTint.setter
    def BorderTint(self, value):
        self.commandbutton.BorderTint = value

    @property
    def BorderWidth(self):
        return self.commandbutton.BorderWidth

    @BorderWidth.setter
    def BorderWidth(self, value):
        self.commandbutton.BorderWidth = value

    @property
    def BottomPadding(self):
        return self.commandbutton.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.commandbutton.BottomPadding = value

    @property
    def Cancel(self):
        return self.commandbutton.Cancel

    @Cancel.setter
    def Cancel(self, value):
        self.commandbutton.Cancel = value

    @property
    def Caption(self):
        return self.commandbutton.Caption

    @Caption.setter
    def Caption(self, value):
        self.commandbutton.Caption = value

    @property
    def Controls(self):
        return Controls(self.commandbutton.Controls)

    @property
    def ControlTipText(self):
        return self.commandbutton.ControlTipText

    @ControlTipText.setter
    def ControlTipText(self, value):
        self.commandbutton.ControlTipText = value

    @property
    def ControlType(self):
        return self.commandbutton.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.commandbutton.ControlType = value

    @property
    def CursorOnHover(self):
        return self.commandbutton.CursorOnHover

    @CursorOnHover.setter
    def CursorOnHover(self, value):
        self.commandbutton.CursorOnHover = value

    @property
    def Default(self):
        return self.commandbutton.Default

    @Default.setter
    def Default(self, value):
        self.commandbutton.Default = value

    @property
    def DisplayWhen(self):
        return self.commandbutton.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.commandbutton.DisplayWhen = value

    @property
    def Enabled(self):
        return self.commandbutton.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.commandbutton.Enabled = value

    @property
    def EventProcPrefix(self):
        return self.commandbutton.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.commandbutton.EventProcPrefix = value

    @property
    def FontBold(self):
        return self.commandbutton.FontBold

    @property
    def FontItalic(self):
        return self.commandbutton.FontItalic

    @property
    def FontName(self):
        return self.commandbutton.FontName

    @property
    def FontSize(self):
        return self.commandbutton.FontSize

    @property
    def FontUnderline(self):
        return self.commandbutton.FontUnderline

    @property
    def FontWeight(self):
        return self.commandbutton.FontWeight

    @FontWeight.setter
    def FontWeight(self, value):
        self.commandbutton.FontWeight = value

    @property
    def ForeColor(self):
        return self.commandbutton.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.commandbutton.ForeColor = value

    @property
    def ForeShade(self):
        return self.commandbutton.ForeShade

    @ForeShade.setter
    def ForeShade(self, value):
        self.commandbutton.ForeShade = value

    @property
    def ForeThemeColorIndex(self):
        return self.commandbutton.ForeThemeColorIndex

    @ForeThemeColorIndex.setter
    def ForeThemeColorIndex(self, value):
        self.commandbutton.ForeThemeColorIndex = value

    @property
    def ForeTint(self):
        return self.commandbutton.ForeTint

    @ForeTint.setter
    def ForeTint(self, value):
        self.commandbutton.ForeTint = value

    @property
    def Glow(self):
        return self.commandbutton.Glow

    @Glow.setter
    def Glow(self, value):
        self.commandbutton.Glow = value

    @property
    def Gradient(self):
        return self.commandbutton.Gradient

    @Gradient.setter
    def Gradient(self, value):
        self.commandbutton.Gradient = value

    @property
    def GridlineColor(self):
        return self.commandbutton.GridlineColor

    @GridlineColor.setter
    def GridlineColor(self, value):
        self.commandbutton.GridlineColor = value

    @property
    def GridlineShade(self):
        return self.commandbutton.GridlineShade

    @GridlineShade.setter
    def GridlineShade(self, value):
        self.commandbutton.GridlineShade = value

    @property
    def GridlineStyleBottom(self):
        return self.commandbutton.GridlineStyleBottom

    @GridlineStyleBottom.setter
    def GridlineStyleBottom(self, value):
        self.commandbutton.GridlineStyleBottom = value

    @property
    def GridlineStyleLeft(self):
        return self.commandbutton.GridlineStyleLeft

    @GridlineStyleLeft.setter
    def GridlineStyleLeft(self, value):
        self.commandbutton.GridlineStyleLeft = value

    @property
    def GridlineStyleRight(self):
        return self.commandbutton.GridlineStyleRight

    @GridlineStyleRight.setter
    def GridlineStyleRight(self, value):
        self.commandbutton.GridlineStyleRight = value

    @property
    def GridlineStyleTop(self):
        return self.commandbutton.GridlineStyleTop

    @GridlineStyleTop.setter
    def GridlineStyleTop(self, value):
        self.commandbutton.GridlineStyleTop = value

    @property
    def GridlineThemeColorIndex(self):
        return self.commandbutton.GridlineThemeColorIndex

    @GridlineThemeColorIndex.setter
    def GridlineThemeColorIndex(self, value):
        self.commandbutton.GridlineThemeColorIndex = value

    @property
    def GridlineTint(self):
        return self.commandbutton.GridlineTint

    @GridlineTint.setter
    def GridlineTint(self, value):
        self.commandbutton.GridlineTint = value

    @property
    def GridlineWidthBottom(self):
        return self.commandbutton.GridlineWidthBottom

    @GridlineWidthBottom.setter
    def GridlineWidthBottom(self, value):
        self.commandbutton.GridlineWidthBottom = value

    @property
    def GridlineWidthLeft(self):
        return self.commandbutton.GridlineWidthLeft

    @GridlineWidthLeft.setter
    def GridlineWidthLeft(self, value):
        self.commandbutton.GridlineWidthLeft = value

    @property
    def GridlineWidthRight(self):
        return self.commandbutton.GridlineWidthRight

    @GridlineWidthRight.setter
    def GridlineWidthRight(self, value):
        self.commandbutton.GridlineWidthRight = value

    @property
    def GridlineWidthTop(self):
        return self.commandbutton.GridlineWidthTop

    @GridlineWidthTop.setter
    def GridlineWidthTop(self, value):
        self.commandbutton.GridlineWidthTop = value

    @property
    def Height(self):
        return self.commandbutton.Height

    @Height.setter
    def Height(self, value):
        self.commandbutton.Height = value

    @property
    def HelpContextId(self):
        return self.commandbutton.HelpContextId

    @HelpContextId.setter
    def HelpContextId(self, value):
        self.commandbutton.HelpContextId = value

    @property
    def HorizontalAnchor(self):
        return self.commandbutton.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.commandbutton.HorizontalAnchor = value

    @property
    def HoverColor(self):
        return self.commandbutton.HoverColor

    @HoverColor.setter
    def HoverColor(self, value):
        self.commandbutton.HoverColor = value

    @property
    def HoverForeColor(self):
        return self.commandbutton.HoverForeColor

    @HoverForeColor.setter
    def HoverForeColor(self, value):
        self.commandbutton.HoverForeColor = value

    @property
    def HoverForeShade(self):
        return self.commandbutton.HoverForeShade

    @HoverForeShade.setter
    def HoverForeShade(self, value):
        self.commandbutton.HoverForeShade = value

    @property
    def HoverForeThemeColorIndex(self):
        return self.commandbutton.HoverForeThemeColorIndex

    @HoverForeThemeColorIndex.setter
    def HoverForeThemeColorIndex(self, value):
        self.commandbutton.HoverForeThemeColorIndex = value

    @property
    def HoverForeTint(self):
        return self.commandbutton.HoverForeTint

    @HoverForeTint.setter
    def HoverForeTint(self, value):
        self.commandbutton.HoverForeTint = value

    @property
    def HoverShade(self):
        return self.commandbutton.HoverShade

    @HoverShade.setter
    def HoverShade(self, value):
        self.commandbutton.HoverShade = value

    @property
    def HoverThemeColorIndex(self):
        return self.commandbutton.HoverThemeColorIndex

    @HoverThemeColorIndex.setter
    def HoverThemeColorIndex(self, value):
        self.commandbutton.HoverThemeColorIndex = value

    @property
    def HoverTint(self):
        return self.commandbutton.HoverTint

    @HoverTint.setter
    def HoverTint(self, value):
        self.commandbutton.HoverTint = value

    @property
    def Hyperlink(self):
        return self.commandbutton.Hyperlink

    @property
    def HyperlinkAddress(self):
        return self.commandbutton.HyperlinkAddress

    @HyperlinkAddress.setter
    def HyperlinkAddress(self, value):
        self.commandbutton.HyperlinkAddress = value

    @property
    def HyperlinkSubAddress(self):
        return self.commandbutton.HyperlinkSubAddress

    @HyperlinkSubAddress.setter
    def HyperlinkSubAddress(self, value):
        self.commandbutton.HyperlinkSubAddress = value

    @property
    def InSelection(self):
        return self.commandbutton.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.commandbutton.InSelection = value

    @property
    def IsVisible(self):
        return self.commandbutton.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.commandbutton.IsVisible = value

    @property
    def LabelAlign(self):
        return self.commandbutton.LabelAlign

    @LabelAlign.setter
    def LabelAlign(self, value):
        self.commandbutton.LabelAlign = value

    @property
    def LabelX(self):
        return self.commandbutton.LabelX

    @LabelX.setter
    def LabelX(self, value):
        self.commandbutton.LabelX = value

    @property
    def LabelY(self):
        return self.commandbutton.LabelY

    @LabelY.setter
    def LabelY(self, value):
        self.commandbutton.LabelY = value

    @property
    def Layout(self):
        return AcLayoutType(self.commandbutton.Layout)

    @property
    def LayoutID(self):
        return self.commandbutton.LayoutID

    @property
    def Left(self):
        return self.commandbutton.Left

    @Left.setter
    def Left(self, value):
        self.commandbutton.Left = value

    @property
    def LeftPadding(self):
        return self.commandbutton.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.commandbutton.LeftPadding = value

    @property
    def Name(self):
        return self.commandbutton.Name

    @Name.setter
    def Name(self, value):
        self.commandbutton.Name = value

    @property
    def ObjectPalette(self):
        return self.commandbutton.ObjectPalette

    @ObjectPalette.setter
    def ObjectPalette(self, value):
        self.commandbutton.ObjectPalette = value

    @property
    def OldValue(self):
        return self.commandbutton.OldValue

    @property
    def OnClick(self):
        return self.commandbutton.OnClick

    @OnClick.setter
    def OnClick(self, value):
        self.commandbutton.OnClick = value

    @property
    def OnDblClick(self):
        return self.commandbutton.OnDblClick

    @OnDblClick.setter
    def OnDblClick(self, value):
        self.commandbutton.OnDblClick = value

    @property
    def OnEnter(self):
        return self.commandbutton.OnEnter

    @OnEnter.setter
    def OnEnter(self, value):
        self.commandbutton.OnEnter = value

    @property
    def OnExit(self):
        return self.commandbutton.OnExit

    @OnExit.setter
    def OnExit(self, value):
        self.commandbutton.OnExit = value

    @property
    def OnGotFocus(self):
        return self.commandbutton.OnGotFocus

    @OnGotFocus.setter
    def OnGotFocus(self, value):
        self.commandbutton.OnGotFocus = value

    @property
    def OnKeyDown(self):
        return self.commandbutton.OnKeyDown

    @OnKeyDown.setter
    def OnKeyDown(self, value):
        self.commandbutton.OnKeyDown = value

    @property
    def OnKeyPress(self):
        return self.commandbutton.OnKeyPress

    @OnKeyPress.setter
    def OnKeyPress(self, value):
        self.commandbutton.OnKeyPress = value

    @property
    def OnKeyUp(self):
        return self.commandbutton.OnKeyUp

    @OnKeyUp.setter
    def OnKeyUp(self, value):
        self.commandbutton.OnKeyUp = value

    @property
    def OnLostFocus(self):
        return self.commandbutton.OnLostFocus

    @OnLostFocus.setter
    def OnLostFocus(self, value):
        self.commandbutton.OnLostFocus = value

    @property
    def OnMouseDown(self):
        return self.commandbutton.OnMouseDown

    @OnMouseDown.setter
    def OnMouseDown(self, value):
        self.commandbutton.OnMouseDown = value

    @property
    def OnMouseMove(self):
        return self.commandbutton.OnMouseMove

    @OnMouseMove.setter
    def OnMouseMove(self, value):
        self.commandbutton.OnMouseMove = value

    @property
    def OnMouseUp(self):
        return self.commandbutton.OnMouseUp

    @OnMouseUp.setter
    def OnMouseUp(self, value):
        self.commandbutton.OnMouseUp = value

    @property
    def OnPush(self):
        return self.commandbutton.OnPush

    @OnPush.setter
    def OnPush(self, value):
        self.commandbutton.OnPush = value

    @property
    def Parent(self):
        return self.commandbutton.Parent

    @property
    def Picture(self):
        return self.commandbutton.Picture

    @Picture.setter
    def Picture(self, value):
        self.commandbutton.Picture = value

    @property
    def PictureCaptionArrangement(self):
        return self.commandbutton.PictureCaptionArrangement

    @PictureCaptionArrangement.setter
    def PictureCaptionArrangement(self, value):
        self.commandbutton.PictureCaptionArrangement = value

    @property
    def PictureData(self):
        return self.commandbutton.PictureData

    @PictureData.setter
    def PictureData(self, value):
        self.commandbutton.PictureData = value

    @property
    def PictureType(self):
        return self.commandbutton.PictureType

    @PictureType.setter
    def PictureType(self, value):
        self.commandbutton.PictureType = value

    @property
    def PressedColor(self):
        return self.commandbutton.PressedColor

    @PressedColor.setter
    def PressedColor(self, value):
        self.commandbutton.PressedColor = value

    @property
    def PressedForeColor(self):
        return self.commandbutton.PressedForeColor

    @PressedForeColor.setter
    def PressedForeColor(self, value):
        self.commandbutton.PressedForeColor = value

    @property
    def PressedForeShade(self):
        return self.commandbutton.PressedForeShade

    @PressedForeShade.setter
    def PressedForeShade(self, value):
        self.commandbutton.PressedForeShade = value

    @property
    def PressedForeThemeColorIndex(self):
        return self.commandbutton.PressedForeThemeColorIndex

    @PressedForeThemeColorIndex.setter
    def PressedForeThemeColorIndex(self, value):
        self.commandbutton.PressedForeThemeColorIndex = value

    @property
    def PressedForeTint(self):
        return self.commandbutton.PressedForeTint

    @PressedForeTint.setter
    def PressedForeTint(self, value):
        self.commandbutton.PressedForeTint = value

    @property
    def PressedShade(self):
        return self.commandbutton.PressedShade

    @PressedShade.setter
    def PressedShade(self, value):
        self.commandbutton.PressedShade = value

    @property
    def PressedThemeColorIndex(self):
        return self.commandbutton.PressedThemeColorIndex

    @PressedThemeColorIndex.setter
    def PressedThemeColorIndex(self, value):
        self.commandbutton.PressedThemeColorIndex = value

    @property
    def PressedTint(self):
        return self.commandbutton.PressedTint

    @PressedTint.setter
    def PressedTint(self, value):
        self.commandbutton.PressedTint = value

    @property
    def Properties(self):
        return Properties(self.commandbutton.Properties)

    @property
    def QuickStyle(self):
        return self.commandbutton.QuickStyle

    @QuickStyle.setter
    def QuickStyle(self, value):
        self.commandbutton.QuickStyle = value

    @property
    def ReadingOrder(self):
        return self.commandbutton.ReadingOrder

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.commandbutton.ReadingOrder = value

    @property
    def RightPadding(self):
        return self.commandbutton.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.commandbutton.RightPadding = value

    @property
    def Section(self):
        return self.commandbutton.Section

    @Section.setter
    def Section(self, value):
        self.commandbutton.Section = value

    @property
    def Shadow(self):
        return self.commandbutton.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.commandbutton.Shadow = value

    @property
    def Shape(self):
        return self.commandbutton.Shape

    @Shape.setter
    def Shape(self, value):
        self.commandbutton.Shape = value

    @property
    def ShortcutMenuBar(self):
        return self.commandbutton.ShortcutMenuBar

    @ShortcutMenuBar.setter
    def ShortcutMenuBar(self, value):
        self.commandbutton.ShortcutMenuBar = value

    @property
    def SoftEdges(self):
        return self.commandbutton.SoftEdges

    @SoftEdges.setter
    def SoftEdges(self, value):
        self.commandbutton.SoftEdges = value

    @property
    def StatusBarText(self):
        return self.commandbutton.StatusBarText

    @StatusBarText.setter
    def StatusBarText(self, value):
        self.commandbutton.StatusBarText = value

    @property
    def TabIndex(self):
        return self.commandbutton.TabIndex

    @TabIndex.setter
    def TabIndex(self, value):
        self.commandbutton.TabIndex = value

    @property
    def TabStop(self):
        return self.commandbutton.TabStop

    @TabStop.setter
    def TabStop(self, value):
        self.commandbutton.TabStop = value

    @property
    def Tag(self):
        return self.commandbutton.Tag

    @Tag.setter
    def Tag(self, value):
        self.commandbutton.Tag = value

    @property
    def ThemeFontIndex(self):
        return self.commandbutton.ThemeFontIndex

    @ThemeFontIndex.setter
    def ThemeFontIndex(self, value):
        self.commandbutton.ThemeFontIndex = value

    @property
    def Top(self):
        return self.commandbutton.Top

    @Top.setter
    def Top(self, value):
        self.commandbutton.Top = value

    @property
    def TopPadding(self):
        return self.commandbutton.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.commandbutton.TopPadding = value

    @property
    def Transparent(self):
        return self.commandbutton.Transparent

    @Transparent.setter
    def Transparent(self, value):
        self.commandbutton.Transparent = value

    @property
    def UseTheme(self):
        return self.commandbutton.UseTheme

    @UseTheme.setter
    def UseTheme(self, value):
        self.commandbutton.UseTheme = value

    @property
    def VerticalAnchor(self):
        return self.commandbutton.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.commandbutton.VerticalAnchor = value

    @property
    def Visible(self):
        return self.commandbutton.Visible

    @Visible.setter
    def Visible(self, value):
        self.commandbutton.Visible = value

    @property
    def Width(self):
        return self.commandbutton.Width

    @Width.setter
    def Width(self, value):
        self.commandbutton.Width = value

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.commandbutton.Move(*args, **arguments)

    def Requery(self):
        self.commandbutton.Requery()

    def SetFocus(self):
        return self.commandbutton.SetFocus()

    def SizeToFit(self):
        self.commandbutton.SizeToFit()

class Control:

    def __init__(self, control=None):
        self.control = control

    @property
    def Application(self):
        return self.control.Application

    @property
    def BottomPadding(self):
        return self.control.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.control.BottomPadding = value

    def Column(self, *args, Index=None, Row=None):
        arguments = {"Index": Index, "Row": Row}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.control.Column):
            return self.control.Column(*args, **arguments)
        else:
            return self.control.GetColumn(*args, **arguments)

    @property
    def Controls(self):
        return Controls(self.control.Controls)

    @property
    def Form(self):
        return self.control.Form

    @property
    def GridlineColor(self):
        return self.control.GridlineColor

    @GridlineColor.setter
    def GridlineColor(self, value):
        self.control.GridlineColor = value

    @property
    def GridlineStyleBottom(self):
        return self.control.GridlineStyleBottom

    @GridlineStyleBottom.setter
    def GridlineStyleBottom(self, value):
        self.control.GridlineStyleBottom = value

    @property
    def GridlineStyleLeft(self):
        return self.control.GridlineStyleLeft

    @GridlineStyleLeft.setter
    def GridlineStyleLeft(self, value):
        self.control.GridlineStyleLeft = value

    @property
    def GridlineStyleRight(self):
        return self.control.GridlineStyleRight

    @GridlineStyleRight.setter
    def GridlineStyleRight(self, value):
        self.control.GridlineStyleRight = value

    @property
    def GridlineStyleTop(self):
        return self.control.GridlineStyleTop

    @GridlineStyleTop.setter
    def GridlineStyleTop(self, value):
        self.control.GridlineStyleTop = value

    @property
    def GridlineWidthBottom(self):
        return self.control.GridlineWidthBottom

    @GridlineWidthBottom.setter
    def GridlineWidthBottom(self, value):
        self.control.GridlineWidthBottom = value

    @property
    def GridlineWidthLeft(self):
        return self.control.GridlineWidthLeft

    @GridlineWidthLeft.setter
    def GridlineWidthLeft(self, value):
        self.control.GridlineWidthLeft = value

    @property
    def GridlineWidthRight(self):
        return self.control.GridlineWidthRight

    @GridlineWidthRight.setter
    def GridlineWidthRight(self, value):
        self.control.GridlineWidthRight = value

    @property
    def GridlineWidthTop(self):
        return self.control.GridlineWidthTop

    @GridlineWidthTop.setter
    def GridlineWidthTop(self, value):
        self.control.GridlineWidthTop = value

    @property
    def HorizontalAnchor(self):
        return self.control.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.control.HorizontalAnchor = value

    @property
    def Hyperlink(self):
        return self.control.Hyperlink

    def ItemData(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.control.ItemData):
            return self.control.ItemData(*args, **arguments)
        else:
            return self.control.GetItemData(*args, **arguments)

    @property
    def ItemsSelected(self):
        return self.control.ItemsSelected

    @property
    def Layout(self):
        return AcLayoutType(self.control.Layout)

    @property
    def LayoutID(self):
        return self.control.LayoutID

    @property
    def LeftPadding(self):
        return self.control.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.control.LeftPadding = value

    @property
    def Name(self):
        return self.control.Name

    @property
    def Object(self):
        return self.control.Object

    def ObjectVerbs(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.control.ObjectVerbs):
            return self.control.ObjectVerbs(*args, **arguments)
        else:
            return self.control.GetObjectVerbs(*args, **arguments)

    @property
    def OldValue(self):
        return self.control.OldValue

    @property
    def Pages(self):
        return Pages(self.control.Pages)

    @property
    def Parent(self):
        return self.control.Parent

    @property
    def Properties(self):
        return Properties(self.control.Properties)

    @property
    def Report(self):
        return self.control.Report

    @property
    def RightPadding(self):
        return self.control.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.control.RightPadding = value

    @property
    def Selected(self):
        return self.control.Selected

    @Selected.setter
    def Selected(self, value):
        self.control.Selected = value

    @property
    def SmartTags(self):
        return SmartTags(self.control.SmartTags)

    @property
    def TopPadding(self):
        return self.control.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.control.TopPadding = value

    @property
    def VerticalAnchor(self):
        return self.control.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.control.VerticalAnchor = value

    def Dropdown(self):
        return self.control.Dropdown()

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.control.Move(*args, **arguments)

    def Requery(self):
        self.control.Requery()

    def SetFocus(self):
        return self.control.SetFocus()

    def SizeToFit(self):
        self.control.SizeToFit()

    def Undo(self):
        self.control.Undo()

class Controls:

    def __init__(self, controls=None):
        self.controls = controls

    @property
    def Application(self):
        return self.controls.Application

    @property
    def Count(self):
        return self.controls.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.controls.Item):
            return self.controls.Item(*args, **arguments)
        else:
            return self.controls.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.controls.Parent

class CurrentData:

    def __init__(self, currentdata=None):
        self.currentdata = currentdata

    @property
    def AllDatabaseDiagrams(self):
        return self.currentdata.AllDatabaseDiagrams

    @property
    def AllFunctions(self):
        return self.currentdata.AllFunctions

    @property
    def AllQueries(self):
        return self.currentdata.AllQueries

    @property
    def AllStoredProcedures(self):
        return self.currentdata.AllStoredProcedures

    @property
    def AllTables(self):
        return self.currentdata.AllTables

    @property
    def AllViews(self):
        return self.currentdata.AllViews

class CurrentProject:

    def __init__(self, currentproject=None):
        self.currentproject = currentproject

    @property
    def AccessConnection(self):
        return self.currentproject.AccessConnection

    @property
    def AllForms(self):
        return self.currentproject.AllForms

    @property
    def AllMacros(self):
        return self.currentproject.AllMacros

    @property
    def AllModules(self):
        return self.currentproject.AllModules

    @property
    def AllReports(self):
        return self.currentproject.AllReports

    @property
    def Application(self):
        return self.currentproject.Application

    @property
    def BaseConnectionString(self):
        return self.currentproject.BaseConnectionString

    @property
    def Connection(self):
        return self.currentproject.Connection

    @property
    def FileFormat(self):
        return AcFileFormat(self.currentproject.FileFormat)

    @property
    def FullName(self):
        return self.currentproject.FullName

    @property
    def ImportExportSpecifications(self):
        return ImportExportSpecifications(self.currentproject.ImportExportSpecifications)

    @property
    def IsConnected(self):
        return self.currentproject.IsConnected

    @property
    def IsTrusted(self):
        return self.currentproject.IsTrusted

    @property
    def IsWeb(self):
        return self.currentproject.IsWeb

    @property
    def Name(self):
        return self.currentproject.Name

    @property
    def Parent(self):
        return self.currentproject.Parent

    @property
    def Path(self):
        return self.currentproject.Path

    @property
    def ProjectType(self):
        return self.currentproject.ProjectType

    @property
    def Properties(self):
        return CurrentProject(self.currentproject.Properties)

    @property
    def RemovePersonalInformation(self):
        return self.currentproject.RemovePersonalInformation

    @RemovePersonalInformation.setter
    def RemovePersonalInformation(self, value):
        self.currentproject.RemovePersonalInformation = value

    @property
    def Resources(self):
        return self.currentproject.Resources

    @property
    def WebSite(self):
        return self.currentproject.WebSite

    def AddSharedImage(self, *args, SharedImageName=None, FileName=None):
        arguments = {"SharedImageName": SharedImageName, "FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.currentproject.AddSharedImage(*args, **arguments)

    def CloseConnection(self):
        return self.currentproject.CloseConnection()

    def OpenConnection(self, *args, BaseConnectionString=None, UserID=None, Password=None):
        arguments = {"BaseConnectionString": BaseConnectionString, "UserID": UserID, "Password": Password}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.currentproject.OpenConnection(*args, **arguments)

    def UpdateDependencyInfo(self):
        self.currentproject.UpdateDependencyInfo()

class CustomControl:

    def __init__(self, customcontrol=None):
        self.customcontrol = customcontrol

    @property
    def About(self):
        return self.customcontrol.About

    @About.setter
    def About(self, value):
        self.customcontrol.About = value

    @property
    def Application(self):
        return self.customcontrol.Application

    @property
    def BorderColor(self):
        return self.customcontrol.BorderColor

    @BorderColor.setter
    def BorderColor(self, value):
        self.customcontrol.BorderColor = value

    @property
    def BorderShade(self):
        return self.customcontrol.BorderShade

    @BorderShade.setter
    def BorderShade(self, value):
        self.customcontrol.BorderShade = value

    @property
    def BorderStyle(self):
        return self.customcontrol.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.customcontrol.BorderStyle = value

    @property
    def BorderThemeColorIndex(self):
        return self.customcontrol.BorderThemeColorIndex

    @BorderThemeColorIndex.setter
    def BorderThemeColorIndex(self, value):
        self.customcontrol.BorderThemeColorIndex = value

    @property
    def BorderTint(self):
        return self.customcontrol.BorderTint

    @BorderTint.setter
    def BorderTint(self, value):
        self.customcontrol.BorderTint = value

    @property
    def BorderWidth(self):
        return self.customcontrol.BorderWidth

    @BorderWidth.setter
    def BorderWidth(self, value):
        self.customcontrol.BorderWidth = value

    @property
    def BottomPadding(self):
        return self.customcontrol.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.customcontrol.BottomPadding = value

    @property
    def Cancel(self):
        return self.customcontrol.Cancel

    @Cancel.setter
    def Cancel(self, value):
        self.customcontrol.Cancel = value

    @property
    def Class(self):
        return self.customcontrol.Class

    @Class.setter
    def Class(self, value):
        self.customcontrol.Class = value

    @property
    def Controls(self):
        return Controls(self.customcontrol.Controls)

    @property
    def ControlSource(self):
        return self.customcontrol.ControlSource

    @ControlSource.setter
    def ControlSource(self, value):
        self.customcontrol.ControlSource = value

    @property
    def ControlTipText(self):
        return self.customcontrol.ControlTipText

    @ControlTipText.setter
    def ControlTipText(self, value):
        self.customcontrol.ControlTipText = value

    @property
    def ControlType(self):
        return self.customcontrol.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.customcontrol.ControlType = value

    @property
    def Custom(self):
        return self.customcontrol.Custom

    @Custom.setter
    def Custom(self, value):
        self.customcontrol.Custom = value

    @property
    def Default(self):
        return self.customcontrol.Default

    @Default.setter
    def Default(self, value):
        self.customcontrol.Default = value

    @property
    def DisplayWhen(self):
        return self.customcontrol.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.customcontrol.DisplayWhen = value

    @property
    def Enabled(self):
        return self.customcontrol.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.customcontrol.Enabled = value

    @property
    def EventProcPrefix(self):
        return self.customcontrol.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.customcontrol.EventProcPrefix = value

    @property
    def GridlineColor(self):
        return self.customcontrol.GridlineColor

    @GridlineColor.setter
    def GridlineColor(self, value):
        self.customcontrol.GridlineColor = value

    @property
    def GridlineStyleBottom(self):
        return self.customcontrol.GridlineStyleBottom

    @GridlineStyleBottom.setter
    def GridlineStyleBottom(self, value):
        self.customcontrol.GridlineStyleBottom = value

    @property
    def GridlineStyleLeft(self):
        return self.customcontrol.GridlineStyleLeft

    @GridlineStyleLeft.setter
    def GridlineStyleLeft(self, value):
        self.customcontrol.GridlineStyleLeft = value

    @property
    def GridlineStyleRight(self):
        return self.customcontrol.GridlineStyleRight

    @GridlineStyleRight.setter
    def GridlineStyleRight(self, value):
        self.customcontrol.GridlineStyleRight = value

    @property
    def GridlineStyleTop(self):
        return self.customcontrol.GridlineStyleTop

    @GridlineStyleTop.setter
    def GridlineStyleTop(self, value):
        self.customcontrol.GridlineStyleTop = value

    @property
    def GridlineWidthBottom(self):
        return self.customcontrol.GridlineWidthBottom

    @GridlineWidthBottom.setter
    def GridlineWidthBottom(self, value):
        self.customcontrol.GridlineWidthBottom = value

    @property
    def GridlineWidthLeft(self):
        return self.customcontrol.GridlineWidthLeft

    @GridlineWidthLeft.setter
    def GridlineWidthLeft(self, value):
        self.customcontrol.GridlineWidthLeft = value

    @property
    def GridlineWidthRight(self):
        return self.customcontrol.GridlineWidthRight

    @GridlineWidthRight.setter
    def GridlineWidthRight(self, value):
        self.customcontrol.GridlineWidthRight = value

    @property
    def GridlineWidthTop(self):
        return self.customcontrol.GridlineWidthTop

    @GridlineWidthTop.setter
    def GridlineWidthTop(self, value):
        self.customcontrol.GridlineWidthTop = value

    @property
    def Height(self):
        return self.customcontrol.Height

    @Height.setter
    def Height(self, value):
        self.customcontrol.Height = value

    @property
    def HelpContextId(self):
        return self.customcontrol.HelpContextId

    @HelpContextId.setter
    def HelpContextId(self, value):
        self.customcontrol.HelpContextId = value

    @property
    def HorizontalAnchor(self):
        return self.customcontrol.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.customcontrol.HorizontalAnchor = value

    @property
    def InSelection(self):
        return self.customcontrol.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.customcontrol.InSelection = value

    @property
    def IsVisible(self):
        return self.customcontrol.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.customcontrol.IsVisible = value

    @property
    def Layout(self):
        return AcLayoutType(self.customcontrol.Layout)

    @property
    def LayoutID(self):
        return self.customcontrol.LayoutID

    @property
    def Left(self):
        return self.customcontrol.Left

    @Left.setter
    def Left(self, value):
        self.customcontrol.Left = value

    @property
    def LeftPadding(self):
        return self.customcontrol.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.customcontrol.LeftPadding = value

    @property
    def Locked(self):
        return self.customcontrol.Locked

    @Locked.setter
    def Locked(self, value):
        self.customcontrol.Locked = value

    @property
    def Name(self):
        return self.customcontrol.Name

    @Name.setter
    def Name(self, value):
        self.customcontrol.Name = value

    @property
    def Object(self):
        return self.customcontrol.Object

    @property
    def ObjectPalette(self):
        return self.customcontrol.ObjectPalette

    @ObjectPalette.setter
    def ObjectPalette(self, value):
        self.customcontrol.ObjectPalette = value

    def ObjectVerbs(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.customcontrol.ObjectVerbs):
            return self.customcontrol.ObjectVerbs(*args, **arguments)
        else:
            return self.customcontrol.GetObjectVerbs(*args, **arguments)

    @property
    def ObjectVerbsCount(self):
        return self.customcontrol.ObjectVerbsCount

    @property
    def OldBorderStyle(self):
        return self.customcontrol.OldBorderStyle

    @OldBorderStyle.setter
    def OldBorderStyle(self, value):
        self.customcontrol.OldBorderStyle = value

    @property
    def OldValue(self):
        return self.customcontrol.OldValue

    @property
    def OLEClass(self):
        return self.customcontrol.OLEClass

    @property
    def OnEnter(self):
        return self.customcontrol.OnEnter

    @OnEnter.setter
    def OnEnter(self, value):
        self.customcontrol.OnEnter = value

    @property
    def OnExit(self):
        return self.customcontrol.OnExit

    @OnExit.setter
    def OnExit(self, value):
        self.customcontrol.OnExit = value

    @property
    def OnGotFocus(self):
        return self.customcontrol.OnGotFocus

    @OnGotFocus.setter
    def OnGotFocus(self, value):
        self.customcontrol.OnGotFocus = value

    @property
    def OnLostFocus(self):
        return self.customcontrol.OnLostFocus

    @OnLostFocus.setter
    def OnLostFocus(self, value):
        self.customcontrol.OnLostFocus = value

    @property
    def OnUpdated(self):
        return self.customcontrol.OnUpdated

    @OnUpdated.setter
    def OnUpdated(self, value):
        self.customcontrol.OnUpdated = value

    @property
    def Parent(self):
        return self.customcontrol.Parent

    @property
    def Properties(self):
        return Properties(self.customcontrol.Properties)

    @property
    def RightPadding(self):
        return self.customcontrol.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.customcontrol.RightPadding = value

    @property
    def Section(self):
        return self.customcontrol.Section

    @Section.setter
    def Section(self, value):
        self.customcontrol.Section = value

    @property
    def SpecialEffect(self):
        return self.customcontrol.SpecialEffect

    @SpecialEffect.setter
    def SpecialEffect(self, value):
        self.customcontrol.SpecialEffect = value

    @property
    def TabIndex(self):
        return self.customcontrol.TabIndex

    @TabIndex.setter
    def TabIndex(self, value):
        self.customcontrol.TabIndex = value

    @property
    def TabStop(self):
        return self.customcontrol.TabStop

    @TabStop.setter
    def TabStop(self, value):
        self.customcontrol.TabStop = value

    @property
    def Tag(self):
        return self.customcontrol.Tag

    @Tag.setter
    def Tag(self, value):
        self.customcontrol.Tag = value

    @property
    def Top(self):
        return self.customcontrol.Top

    @Top.setter
    def Top(self, value):
        self.customcontrol.Top = value

    @property
    def TopPadding(self):
        return self.customcontrol.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.customcontrol.TopPadding = value

    @property
    def Value(self):
        return self.customcontrol.Value

    @Value.setter
    def Value(self, value):
        self.customcontrol.Value = value

    @property
    def VarOleObject(self):
        return self.customcontrol.VarOleObject

    @property
    def Verb(self):
        return self.customcontrol.Verb

    @Verb.setter
    def Verb(self, value):
        self.customcontrol.Verb = value

    @property
    def VerticalAnchor(self):
        return self.customcontrol.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.customcontrol.VerticalAnchor = value

    @property
    def Visible(self):
        return self.customcontrol.Visible

    @Visible.setter
    def Visible(self, value):
        self.customcontrol.Visible = value

    @property
    def Width(self):
        return self.customcontrol.Width

    @Width.setter
    def Width(self, value):
        self.customcontrol.Width = value

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.customcontrol.Move(*args, **arguments)

    def Requery(self):
        self.customcontrol.Requery()

    def SetFocus(self):
        return self.customcontrol.SetFocus()

    def SizeToFit(self):
        self.customcontrol.SizeToFit()

class DependencyInfo:

    def __init__(self, dependencyinfo=None):
        self.dependencyinfo = dependencyinfo

    @property
    def Dependants(self):
        return DependencyObjects(self.dependencyinfo.Dependants)

    @property
    def Dependencies(self):
        return DependencyObjects(self.dependencyinfo.Dependencies)

    @property
    def InsufficientPermissions(self):
        return DependencyObjects(self.dependencyinfo.InsufficientPermissions)

    @property
    def OutOfDateObjects(self):
        return DependencyObjects(self.dependencyinfo.OutOfDateObjects)

    @property
    def Parent(self):
        return self.dependencyinfo.Parent

    @property
    def UnsupportedObjects(self):
        return DependencyObjects(self.dependencyinfo.UnsupportedObjects)

class DependencyObjects:

    def __init__(self, dependencyobjects=None):
        self.dependencyobjects = dependencyobjects

    @property
    def Application(self):
        return self.dependencyobjects.Application

    @property
    def Count(self):
        return self.dependencyobjects.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.dependencyobjects.Item):
            return self.dependencyobjects.Item(*args, **arguments)
        else:
            return self.dependencyobjects.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.dependencyobjects.Parent

class DoCmd:

    def __init__(self, docmd=None):
        self.docmd = docmd

    def AddMenu(self, *args, MenuName=None, MenuMacroName=None, StatusBarText=None):
        arguments = {"MenuName": MenuName, "MenuMacroName": MenuMacroName, "StatusBarText": StatusBarText}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.AddMenu(*args, **arguments)

    def ApplyFilter(self, *args, FilterName=None, WhereCondition=None, ControlName=None):
        arguments = {"FilterName": FilterName, "WhereCondition": WhereCondition, "ControlName": ControlName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.ApplyFilter(*args, **arguments)

    def Beep(self):
        self.docmd.Beep()

    def BrowseTo(self, *args, ObjectType=None, ObjectName=None, PathtoSubformControl=None, WhereCondition=None, Page=None, DataMode=None):
        arguments = {"ObjectType": ObjectType, "ObjectName": ObjectName, "PathtoSubformControl": PathtoSubformControl, "WhereCondition": WhereCondition, "Page": Page, "DataMode": DataMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.BrowseTo(*args, **arguments)

    def CancelEvent(self):
        self.docmd.CancelEvent()

    def ClearMacroError(self):
        self.docmd.ClearMacroError()

    def Close(self, *args, ObjectType=None, ObjectName=None, Save=None):
        arguments = {"ObjectType": ObjectType, "ObjectName": ObjectName, "Save": Save}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.Close(*args, **arguments)

    def CloseDatabase(self):
        self.docmd.CloseDatabase()

    def CopyDatabaseFile(self, *args, DatabaseFileName=None, OverwriteExistingFile=None, DisconnectAllUsers=None):
        arguments = {"DatabaseFileName": DatabaseFileName, "OverwriteExistingFile": OverwriteExistingFile, "DisconnectAllUsers": DisconnectAllUsers}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.CopyDatabaseFile(*args, **arguments)

    def CopyObject(self, *args, DestinationDatabase=None, NewName=None, SourceObjectType=None, SourceObjectName=None):
        arguments = {"DestinationDatabase": DestinationDatabase, "NewName": NewName, "SourceObjectType": SourceObjectType, "SourceObjectName": SourceObjectName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.CopyObject(*args, **arguments)

    def DeleteObject(self, *args, ObjectType=None, ObjectName=None):
        arguments = {"ObjectType": ObjectType, "ObjectName": ObjectName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.DeleteObject(*args, **arguments)

    def DoMenuItem(self, *args, MenuBar=None, MenuName=None, Command=None, Subcommand=None, Version=None):
        arguments = {"MenuBar": MenuBar, "MenuName": MenuName, "Command": Command, "Subcommand": Subcommand, "Version": Version}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.DoMenuItem(*args, **arguments)

    def Echo(self, *args, EchoOn=None, StatusBarText=None):
        arguments = {"EchoOn": EchoOn, "StatusBarText": StatusBarText}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.Echo(*args, **arguments)

    def FindNext(self):
        self.docmd.FindNext()

    def FindRecord(self, *args, FindWhat=None, Match=None, MatchCase=None, Search=None, SearchAsFormatted=None, OnlyCurrentField=None, FindFirst=None):
        arguments = {"FindWhat": FindWhat, "Match": Match, "MatchCase": MatchCase, "Search": Search, "SearchAsFormatted": SearchAsFormatted, "OnlyCurrentField": OnlyCurrentField, "FindFirst": FindFirst}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.FindRecord(*args, **arguments)

    def GoToControl(self, *args, ControlName=None):
        arguments = {"ControlName": ControlName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.GoToControl(*args, **arguments)

    def GoToPage(self, *args, PageNumber=None, Right=None, Down=None):
        arguments = {"PageNumber": PageNumber, "Right": Right, "Down": Down}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.docmd.GoToPage(*args, **arguments)

    def GoToRecord(self, *args, ObjectType=None, ObjectName=None, Record=None, Offset=None):
        arguments = {"ObjectType": ObjectType, "ObjectName": ObjectName, "Record": Record, "Offset": Offset}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.GoToRecord(*args, **arguments)

    def Hourglass(self, *args, HourglassOn=None):
        arguments = {"HourglassOn": HourglassOn}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.Hourglass(*args, **arguments)

    def LockNavigationPane(self, *args, Lock=None):
        arguments = {"Lock": Lock}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.LockNavigationPane(*args, **arguments)

    def Maximize(self):
        self.docmd.Maximize()

    def Minimize(self):
        self.docmd.Minimize()

    def MoveSize(self, *args, Right=None, Down=None, Width=None, Height=None):
        arguments = {"Right": Right, "Down": Down, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.MoveSize(*args, **arguments)

    def NavigateTo(self, *args, Category=None, Group=None):
        arguments = {"Category": Category, "Group": Group}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.NavigateTo(*args, **arguments)

    def OpenDataAccessPage(self, *args, DataAccessPageName=None, View=None):
        arguments = {"DataAccessPageName": DataAccessPageName, "View": View}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.OpenDataAccessPage(*args, **arguments)

    def OpenDiagram(self, *args, DiagramName=None):
        arguments = {"DiagramName": DiagramName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.OpenDiagram(*args, **arguments)

    def OpenForm(self, *args, FormName=None, View=None, FilterName=None, WhereCondition=None, DataMode=None, WindowMode=None, OpenArgs=None):
        arguments = {"FormName": FormName, "View": View, "FilterName": FilterName, "WhereCondition": WhereCondition, "DataMode": DataMode, "WindowMode": WindowMode, "OpenArgs": OpenArgs}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.OpenForm(*args, **arguments)

    def OpenFunction(self, *args, FunctionName=None, View=None, DataMode=None):
        arguments = {"FunctionName": FunctionName, "View": View, "DataMode": DataMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.OpenFunction(*args, **arguments)

    def OpenModule(self, *args, ModuleName=None, ProcedureName=None):
        arguments = {"ModuleName": ModuleName, "ProcedureName": ProcedureName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.OpenModule(*args, **arguments)

    def OpenQuery(self, *args, QueryName=None, View=None, DataMode=None):
        arguments = {"QueryName": QueryName, "View": View, "DataMode": DataMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.OpenQuery(*args, **arguments)

    def OpenReport(self, *args, ReportName=None, View=None, FilterName=None, WhereCondition=None, WindowMode=None, OpenArgs=None):
        arguments = {"ReportName": ReportName, "View": View, "FilterName": FilterName, "WhereCondition": WhereCondition, "WindowMode": WindowMode, "OpenArgs": OpenArgs}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.OpenReport(*args, **arguments)

    def OpenStoredProcedure(self, *args, ProcedureName=None, View=None, DataMode=None):
        arguments = {"ProcedureName": ProcedureName, "View": View, "DataMode": DataMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.OpenStoredProcedure(*args, **arguments)

    def OpenTable(self, *args, TableName=None, View=None, DataMode=None):
        arguments = {"TableName": TableName, "View": View, "DataMode": DataMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.OpenTable(*args, **arguments)

    def OpenView(self, *args, ViewName=None, View=None, DataMode=None):
        arguments = {"ViewName": ViewName, "View": View, "DataMode": DataMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.OpenView(*args, **arguments)

    def OutputTo(self, *args, ObjectType=None, ObjectName=None, OutputFormat=None, OutputFile=None, AutoStart=None, TemplateFile=None, Encoding=None, OutputQuality=None):
        arguments = {"ObjectType": ObjectType, "ObjectName": ObjectName, "OutputFormat": OutputFormat, "OutputFile": OutputFile, "AutoStart": AutoStart, "TemplateFile": TemplateFile, "Encoding": Encoding, "OutputQuality": OutputQuality}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.OutputTo(*args, **arguments)

    def PrintOut(self, *args, PrintRange=None, PageFrom=None, PageTo=None, PrintQuality=None, Copies=None, CollateCopies=None):
        arguments = {"PrintRange": PrintRange, "PageFrom": PageFrom, "PageTo": PageTo, "PrintQuality": PrintQuality, "Copies": Copies, "CollateCopies": CollateCopies}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.PrintOut(*args, **arguments)

    def Quit(self, *args, Options=None):
        arguments = {"Options": Options}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.Quit(*args, **arguments)

    def RefreshRecord(self):
        self.docmd.RefreshRecord()

    def Rename(self, *args, NewName=None, ObjectType=None, OldName=None):
        arguments = {"NewName": NewName, "ObjectType": ObjectType, "OldName": OldName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.Rename(*args, **arguments)

    def RepaintObject(self, *args, ObjectType=None, ObjectName=None):
        arguments = {"ObjectType": ObjectType, "ObjectName": ObjectName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.RepaintObject(*args, **arguments)

    def Requery(self, *args, ControlName=None):
        arguments = {"ControlName": ControlName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.Requery(*args, **arguments)

    def Restore(self):
        self.docmd.Restore()

    def RunCommand(self, *args, Command=None):
        arguments = {"Command": Command}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.RunCommand(*args, **arguments)

    def RunDataMacro(self, *args, MacroName=None):
        arguments = {"MacroName": MacroName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.RunDataMacro(*args, **arguments)

    def RunMacro(self, *args, MacroName=None, RepeatCount=None, RepeatExpression=None):
        arguments = {"MacroName": MacroName, "RepeatCount": RepeatCount, "RepeatExpression": RepeatExpression}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.RunMacro(*args, **arguments)

    def RunSavedImportExport(self, *args, SavedImportExportName=None):
        arguments = {"SavedImportExportName": SavedImportExportName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.RunSavedImportExport(*args, **arguments)

    def RunSQL(self, *args, SQLStatement=None, UseTransaction=None):
        arguments = {"SQLStatement": SQLStatement, "UseTransaction": UseTransaction}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.RunSQL(*args, **arguments)

    def Save(self, *args, ObjectType=None, ObjectName=None):
        arguments = {"ObjectType": ObjectType, "ObjectName": ObjectName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.Save(*args, **arguments)

    def SearchForRecord(self, *args, ObjectType=None, ObjectName=None, Record=None, WhereCondition=None):
        arguments = {"ObjectType": ObjectType, "ObjectName": ObjectName, "Record": Record, "WhereCondition": WhereCondition}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.SearchForRecord(*args, **arguments)

    def SelectObject(self, *args, ObjectType=None, ObjectName=None, InNavigationPane=None):
        arguments = {"ObjectType": ObjectType, "ObjectName": ObjectName, "InNavigationPane": InNavigationPane}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.SelectObject(*args, **arguments)

    def SendObject(self, *args, ObjectType=None, ObjectName=None, OutputFormat=None, To=None, Cc=None, Bcc=None, Subject=None, MessageText=None, EditMessage=None, TemplateFile=None):
        arguments = {"ObjectType": ObjectType, "ObjectName": ObjectName, "OutputFormat": OutputFormat, "To": To, "Cc": Cc, "Bcc": Bcc, "Subject": Subject, "MessageText": MessageText, "EditMessage": EditMessage, "TemplateFile": TemplateFile}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.SendObject(*args, **arguments)

    def SetDisplayedCategories(self, *args, Show=None, Category=None):
        arguments = {"Show": Show, "Category": Category}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.SetDisplayedCategories(*args, **arguments)

    def SetFilter(self, *args, FilterName=None, WhereCondition=None, ControlName=None):
        arguments = {"FilterName": FilterName, "WhereCondition": WhereCondition, "ControlName": ControlName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.SetFilter(*args, **arguments)

    def SetMenuItem(self, *args, MenuIndex=None, CommandIndex=None, SubcommandIndex=None, Flag=None):
        arguments = {"MenuIndex": MenuIndex, "CommandIndex": CommandIndex, "SubcommandIndex": SubcommandIndex, "Flag": Flag}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.SetMenuItem(*args, **arguments)

    def SetOrderBy(self, *args, OrderBy=None, ControlName=None):
        arguments = {"OrderBy": OrderBy, "ControlName": ControlName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.SetOrderBy(*args, **arguments)

    def SetParameter(self, *args, Name=None, Expression=None):
        arguments = {"Name": Name, "Expression": Expression}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.SetParameter(*args, **arguments)

    def SetProperty(self, *args, ControlName=None, Property=None, Value=None):
        arguments = {"ControlName": ControlName, "Property": Property, "Value": Value}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.SetProperty(*args, **arguments)

    def SetWarnings(self, *args, WarningsOn=None):
        arguments = {"WarningsOn": WarningsOn}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.SetWarnings(*args, **arguments)

    def ShowAllRecords(self):
        self.docmd.ShowAllRecords()

    def ShowToolbar(self, *args, ToolbarName=None, Show=None):
        arguments = {"ToolbarName": ToolbarName, "Show": Show}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.ShowToolbar(*args, **arguments)

    def SingleStep(self):
        self.docmd.SingleStep()

    def TransferDatabase(self, *args, TransferType=None, DatabaseType=None, DatabaseName=None, ObjectType=None, Source=None, Destination=None, StructureOnly=None, StoreLogin=None):
        arguments = {"TransferType": TransferType, "DatabaseType": DatabaseType, "DatabaseName": DatabaseName, "ObjectType": ObjectType, "Source": Source, "Destination": Destination, "StructureOnly": StructureOnly, "StoreLogin": StoreLogin}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.TransferDatabase(*args, **arguments)

    def TransferSharePointList(self, *args, TransferType=None, SiteAddress=None, ListID=None, ViewID=None, TableName=None, GetLookupDisplayValues=None):
        arguments = {"TransferType": TransferType, "SiteAddress": SiteAddress, "ListID": ListID, "ViewID": ViewID, "TableName": TableName, "GetLookupDisplayValues": GetLookupDisplayValues}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.TransferSharePointList(*args, **arguments)

    def TransferSpreadsheet(self, *args, TransferType=None, SpreadsheetType=None, TableName=None, FileName=None, HasFieldNames=None, Range=None, UseOA=None):
        arguments = {"TransferType": TransferType, "SpreadsheetType": SpreadsheetType, "TableName": TableName, "FileName": FileName, "HasFieldNames": HasFieldNames, "Range": Range, "UseOA": UseOA}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.TransferSpreadsheet(*args, **arguments)

    def TransferSQLDatabase(self, *args, Server=None, Database=None, UseTrustedConnection=None, Login=None, Password=None, TransferCopyData=None):
        arguments = {"Server": Server, "Database": Database, "UseTrustedConnection": UseTrustedConnection, "Login": Login, "Password": Password, "TransferCopyData": TransferCopyData}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.TransferSQLDatabase(*args, **arguments)

    def TransferText(self, *args, TransferType=None, SpecificationName=None, TableName=None, FileName=None, HasFieldNames=None, HTMLTableName=None, CodePage=None):
        arguments = {"TransferType": TransferType, "SpecificationName": SpecificationName, "TableName": TableName, "FileName": FileName, "HasFieldNames": HasFieldNames, "HTMLTableName": HTMLTableName, "CodePage": CodePage}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.docmd.TransferText(*args, **arguments)

class EdgeBrowserControl:

    def __init__(self, edgebrowsercontrol=None):
        self.edgebrowsercontrol = edgebrowsercontrol

    @property
    def TrustedDomains(self):
        return self.edgebrowsercontrol.TrustedDomains

    @TrustedDomains.setter
    def TrustedDomains(self, value):
        self.edgebrowsercontrol.TrustedDomains = value

    def ExecuteJavascript(self, *args, script=None):
        arguments = {"script": script}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.edgebrowsercontrol.ExecuteJavascript(*args, **arguments)

    def Navigate(self, *args, url=None):
        arguments = {"url": url}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.edgebrowsercontrol.Navigate(*args, **arguments)

    def Refresh(self):
        self.edgebrowsercontrol.Refresh()

    def RetrieveJavascriptValue(self, *args, expression=None):
        arguments = {"expression": expression}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.edgebrowsercontrol.RetrieveJavascriptValue(*args, **arguments)

class EmptyCell:

    def __init__(self, emptycell=None):
        self.emptycell = emptycell

    @property
    def Application(self):
        return self.emptycell.Application

    @property
    def BackColor(self):
        return self.emptycell.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.emptycell.BackColor = value

    @property
    def BackShade(self):
        return self.emptycell.BackShade

    @BackShade.setter
    def BackShade(self, value):
        self.emptycell.BackShade = value

    @property
    def BackStyle(self):
        return self.emptycell.BackStyle

    @BackStyle.setter
    def BackStyle(self, value):
        self.emptycell.BackStyle = value

    @property
    def BackThemeColorIndex(self):
        return self.emptycell.BackThemeColorIndex

    @BackThemeColorIndex.setter
    def BackThemeColorIndex(self, value):
        self.emptycell.BackThemeColorIndex = value

    @property
    def BackTint(self):
        return self.emptycell.BackTint

    @BackTint.setter
    def BackTint(self, value):
        self.emptycell.BackTint = value

    @property
    def BottomPadding(self):
        return self.emptycell.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.emptycell.BottomPadding = value

    @property
    def ControlType(self):
        return self.emptycell.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.emptycell.ControlType = value

    @property
    def DisplayWhen(self):
        return self.emptycell.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.emptycell.DisplayWhen = value

    @property
    def EventProcPrefix(self):
        return self.emptycell.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.emptycell.EventProcPrefix = value

    @property
    def GridlineColor(self):
        return self.emptycell.GridlineColor

    @GridlineColor.setter
    def GridlineColor(self, value):
        self.emptycell.GridlineColor = value

    @property
    def GridlineShade(self):
        return self.emptycell.GridlineShade

    @GridlineShade.setter
    def GridlineShade(self, value):
        self.emptycell.GridlineShade = value

    @property
    def GridlineStyleBottom(self):
        return self.emptycell.GridlineStyleBottom

    @GridlineStyleBottom.setter
    def GridlineStyleBottom(self, value):
        self.emptycell.GridlineStyleBottom = value

    @property
    def GridlineStyleLeft(self):
        return self.emptycell.GridlineStyleLeft

    @GridlineStyleLeft.setter
    def GridlineStyleLeft(self, value):
        self.emptycell.GridlineStyleLeft = value

    @property
    def GridlineStyleRight(self):
        return self.emptycell.GridlineStyleRight

    @GridlineStyleRight.setter
    def GridlineStyleRight(self, value):
        self.emptycell.GridlineStyleRight = value

    @property
    def GridlineStyleTop(self):
        return self.emptycell.GridlineStyleTop

    @GridlineStyleTop.setter
    def GridlineStyleTop(self, value):
        self.emptycell.GridlineStyleTop = value

    @property
    def GridlineThemeColorIndex(self):
        return self.emptycell.GridlineThemeColorIndex

    @GridlineThemeColorIndex.setter
    def GridlineThemeColorIndex(self, value):
        self.emptycell.GridlineThemeColorIndex = value

    @property
    def GridlineTint(self):
        return self.emptycell.GridlineTint

    @GridlineTint.setter
    def GridlineTint(self, value):
        self.emptycell.GridlineTint = value

    @property
    def GridlineWidthBottom(self):
        return self.emptycell.GridlineWidthBottom

    @GridlineWidthBottom.setter
    def GridlineWidthBottom(self, value):
        self.emptycell.GridlineWidthBottom = value

    @property
    def GridlineWidthLeft(self):
        return self.emptycell.GridlineWidthLeft

    @GridlineWidthLeft.setter
    def GridlineWidthLeft(self, value):
        self.emptycell.GridlineWidthLeft = value

    @property
    def GridlineWidthRight(self):
        return self.emptycell.GridlineWidthRight

    @GridlineWidthRight.setter
    def GridlineWidthRight(self, value):
        self.emptycell.GridlineWidthRight = value

    @property
    def GridlineWidthTop(self):
        return self.emptycell.GridlineWidthTop

    @GridlineWidthTop.setter
    def GridlineWidthTop(self, value):
        self.emptycell.GridlineWidthTop = value

    @property
    def Height(self):
        return self.emptycell.Height

    @Height.setter
    def Height(self, value):
        self.emptycell.Height = value

    @property
    def HelpContextId(self):
        return self.emptycell.HelpContextId

    @HelpContextId.setter
    def HelpContextId(self, value):
        self.emptycell.HelpContextId = value

    @property
    def HorizontalAnchor(self):
        return self.emptycell.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.emptycell.HorizontalAnchor = value

    @property
    def InSelection(self):
        return self.emptycell.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.emptycell.InSelection = value

    @property
    def IsVisible(self):
        return self.emptycell.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.emptycell.IsVisible = value

    @property
    def Layout(self):
        return AcLayoutType(self.emptycell.Layout)

    @property
    def LayoutID(self):
        return self.emptycell.LayoutID

    @property
    def Left(self):
        return self.emptycell.Left

    @Left.setter
    def Left(self, value):
        self.emptycell.Left = value

    @property
    def LeftPadding(self):
        return self.emptycell.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.emptycell.LeftPadding = value

    @property
    def Name(self):
        return self.emptycell.Name

    @Name.setter
    def Name(self, value):
        self.emptycell.Name = value

    @property
    def Parent(self):
        return self.emptycell.Parent

    @property
    def Properties(self):
        return Properties(self.emptycell.Properties)

    @property
    def RightPadding(self):
        return self.emptycell.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.emptycell.RightPadding = value

    @property
    def Section(self):
        return self.emptycell.Section

    @Section.setter
    def Section(self, value):
        self.emptycell.Section = value

    @property
    def ShortcutMenuBar(self):
        return self.emptycell.ShortcutMenuBar

    @ShortcutMenuBar.setter
    def ShortcutMenuBar(self, value):
        self.emptycell.ShortcutMenuBar = value

    @property
    def SpecialEffect(self):
        return self.emptycell.SpecialEffect

    @SpecialEffect.setter
    def SpecialEffect(self, value):
        self.emptycell.SpecialEffect = value

    @property
    def StatusBarText(self):
        return self.emptycell.StatusBarText

    @StatusBarText.setter
    def StatusBarText(self, value):
        self.emptycell.StatusBarText = value

    @property
    def Tag(self):
        return self.emptycell.Tag

    @Tag.setter
    def Tag(self, value):
        self.emptycell.Tag = value

    @property
    def Top(self):
        return self.emptycell.Top

    @Top.setter
    def Top(self, value):
        self.emptycell.Top = value

    @property
    def TopPadding(self):
        return self.emptycell.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.emptycell.TopPadding = value

    @property
    def VerticalAnchor(self):
        return self.emptycell.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.emptycell.VerticalAnchor = value

    @property
    def Visible(self):
        return self.emptycell.Visible

    @Visible.setter
    def Visible(self, value):
        self.emptycell.Visible = value

    @property
    def Width(self):
        return self.emptycell.Width

    @Width.setter
    def Width(self, value):
        self.emptycell.Width = value

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.emptycell.Move(*args, **arguments)

    def SizeToFit(self):
        self.emptycell.SizeToFit()

class Entities:

    def __init__(self, entities=None):
        self.entities = entities

    @property
    def Count(self):
        return self.entities.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.entities.Item):
            return self.entities.Item(*args, **arguments)
        else:
            return self.entities.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.entities.Parent

class Entity:

    def __init__(self, entity=None):
        self.entity = entity

    @property
    def Name(self):
        return self.entity.Name

    @Name.setter
    def Name(self, value):
        self.entity.Name = value

    @property
    def Operations(self):
        return self.entity.Operations

    @property
    def Parent(self):
        return self.entity.Parent

class Form:

    def __init__(self, form=None):
        self.form = form

    @property
    def ActiveControl(self):
        return self.form.ActiveControl

    @property
    def AfterDelConfirm(self):
        return self.form.AfterDelConfirm

    @AfterDelConfirm.setter
    def AfterDelConfirm(self, value):
        self.form.AfterDelConfirm = value

    @property
    def AfterFinalRender(self):
        return self.form.AfterFinalRender

    @AfterFinalRender.setter
    def AfterFinalRender(self, value):
        self.form.AfterFinalRender = value

    @property
    def AfterInsert(self):
        return self.form.AfterInsert

    @AfterInsert.setter
    def AfterInsert(self, value):
        self.form.AfterInsert = value

    @property
    def AfterLayout(self):
        return self.form.AfterLayout

    @AfterLayout.setter
    def AfterLayout(self, value):
        self.form.AfterLayout = value

    @property
    def AfterRender(self):
        return self.form.AfterRender

    @AfterRender.setter
    def AfterRender(self, value):
        self.form.AfterRender = value

    @property
    def AllowAdditions(self):
        return self.form.AllowAdditions

    @AllowAdditions.setter
    def AllowAdditions(self, value):
        self.form.AllowAdditions = value

    @property
    def AllowDatasheetView(self):
        return self.form.AllowDatasheetView

    @AllowDatasheetView.setter
    def AllowDatasheetView(self, value):
        self.form.AllowDatasheetView = value

    @property
    def AllowDeletions(self):
        return self.form.AllowDeletions

    @AllowDeletions.setter
    def AllowDeletions(self, value):
        self.form.AllowDeletions = value

    @property
    def AllowEdits(self):
        return self.form.AllowEdits

    @AllowEdits.setter
    def AllowEdits(self, value):
        self.form.AllowEdits = value

    @property
    def AllowFilters(self):
        return self.form.AllowFilters

    @AllowFilters.setter
    def AllowFilters(self, value):
        self.form.AllowFilters = value

    @property
    def AllowFormView(self):
        return self.form.AllowFormView

    @AllowFormView.setter
    def AllowFormView(self, value):
        self.form.AllowFormView = value

    @property
    def AllowLayoutView(self):
        return self.form.AllowLayoutView

    @AllowLayoutView.setter
    def AllowLayoutView(self, value):
        self.form.AllowLayoutView = value

    @property
    def AllowPivotChartView(self):
        return self.form.AllowPivotChartView

    @AllowPivotChartView.setter
    def AllowPivotChartView(self, value):
        self.form.AllowPivotChartView = value

    @property
    def AllowPivotTableView(self):
        return self.form.AllowPivotTableView

    @AllowPivotTableView.setter
    def AllowPivotTableView(self, value):
        self.form.AllowPivotTableView = value

    @property
    def Application(self):
        return self.form.Application

    @property
    def AutoCenter(self):
        return self.form.AutoCenter

    @AutoCenter.setter
    def AutoCenter(self, value):
        self.form.AutoCenter = value

    @property
    def AutoResize(self):
        return self.form.AutoResize

    @AutoResize.setter
    def AutoResize(self, value):
        self.form.AutoResize = value

    @property
    def BeforeDelConfirm(self):
        return self.form.BeforeDelConfirm

    @BeforeDelConfirm.setter
    def BeforeDelConfirm(self, value):
        self.form.BeforeDelConfirm = value

    @property
    def BeforeInsert(self):
        return self.form.BeforeInsert

    @BeforeInsert.setter
    def BeforeInsert(self, value):
        self.form.BeforeInsert = value

    @property
    def BeforeQuery(self):
        return self.form.BeforeQuery

    @BeforeQuery.setter
    def BeforeQuery(self, value):
        self.form.BeforeQuery = value

    @property
    def BeforeRender(self):
        return self.form.BeforeRender

    @BeforeRender.setter
    def BeforeRender(self, value):
        self.form.BeforeRender = value

    @property
    def BeforeScreenTip(self):
        return self.form.BeforeScreenTip

    @BeforeScreenTip.setter
    def BeforeScreenTip(self, value):
        self.form.BeforeScreenTip = value

    @property
    def Bookmark(self):
        return self.form.Bookmark

    @Bookmark.setter
    def Bookmark(self, value):
        self.form.Bookmark = value

    @property
    def BorderStyle(self):
        return self.form.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.form.BorderStyle = value

    @property
    def Caption(self):
        return self.form.Caption

    @Caption.setter
    def Caption(self, value):
        self.form.Caption = value

    @property
    def ChartSpace(self):
        return self.form.ChartSpace

    @property
    def CloseButton(self):
        return self.form.CloseButton

    @CloseButton.setter
    def CloseButton(self, value):
        self.form.CloseButton = value

    @property
    def CommandBeforeExecute(self):
        return self.form.CommandBeforeExecute

    @CommandBeforeExecute.setter
    def CommandBeforeExecute(self, value):
        self.form.CommandBeforeExecute = value

    @property
    def CommandChecked(self):
        return self.form.CommandChecked

    @CommandChecked.setter
    def CommandChecked(self, value):
        self.form.CommandChecked = value

    @property
    def CommandEnabled(self):
        return self.form.CommandEnabled

    @CommandEnabled.setter
    def CommandEnabled(self, value):
        self.form.CommandEnabled = value

    @property
    def CommandExecute(self):
        return self.form.CommandExecute

    @CommandExecute.setter
    def CommandExecute(self, value):
        self.form.CommandExecute = value

    @property
    def ControlBox(self):
        return self.form.ControlBox

    @ControlBox.setter
    def ControlBox(self, value):
        self.form.ControlBox = value

    @property
    def Controls(self):
        return Controls(self.form.Controls)

    @property
    def Count(self):
        return self.form.Count

    @property
    def CurrentRecord(self):
        return self.form.CurrentRecord

    @CurrentRecord.setter
    def CurrentRecord(self, value):
        self.form.CurrentRecord = value

    @property
    def CurrentSectionLeft(self):
        return self.form.CurrentSectionLeft

    @CurrentSectionLeft.setter
    def CurrentSectionLeft(self, value):
        self.form.CurrentSectionLeft = value

    @property
    def CurrentSectionTop(self):
        return self.form.CurrentSectionTop

    @CurrentSectionTop.setter
    def CurrentSectionTop(self, value):
        self.form.CurrentSectionTop = value

    @property
    def CurrentView(self):
        return self.form.CurrentView

    @CurrentView.setter
    def CurrentView(self, value):
        self.form.CurrentView = value

    @property
    def Cycle(self):
        return self.form.Cycle

    @Cycle.setter
    def Cycle(self, value):
        self.form.Cycle = value

    @property
    def DataChange(self):
        return self.form.DataChange

    @DataChange.setter
    def DataChange(self, value):
        self.form.DataChange = value

    @property
    def DataEntry(self):
        return self.form.DataEntry

    @DataEntry.setter
    def DataEntry(self, value):
        self.form.DataEntry = value

    @property
    def DataSetChange(self):
        return self.form.DataSetChange

    @DataSetChange.setter
    def DataSetChange(self, value):
        self.form.DataSetChange = value

    @property
    def DatasheetAlternateBackColor(self):
        return self.form.DatasheetAlternateBackColor

    @DatasheetAlternateBackColor.setter
    def DatasheetAlternateBackColor(self, value):
        self.form.DatasheetAlternateBackColor = value

    @property
    def DatasheetBackColor(self):
        return self.form.DatasheetBackColor

    @DatasheetBackColor.setter
    def DatasheetBackColor(self, value):
        self.form.DatasheetBackColor = value

    @property
    def DatasheetBorderLineStyle(self):
        return self.form.DatasheetBorderLineStyle

    @DatasheetBorderLineStyle.setter
    def DatasheetBorderLineStyle(self, value):
        self.form.DatasheetBorderLineStyle = value

    @property
    def DatasheetCellsEffect(self):
        return self.form.DatasheetCellsEffect

    @DatasheetCellsEffect.setter
    def DatasheetCellsEffect(self, value):
        self.form.DatasheetCellsEffect = value

    @property
    def DatasheetColumnHeaderUnderlineStyle(self):
        return self.form.DatasheetColumnHeaderUnderlineStyle

    @DatasheetColumnHeaderUnderlineStyle.setter
    def DatasheetColumnHeaderUnderlineStyle(self, value):
        self.form.DatasheetColumnHeaderUnderlineStyle = value

    @property
    def DatasheetFontHeight(self):
        return self.form.DatasheetFontHeight

    @DatasheetFontHeight.setter
    def DatasheetFontHeight(self, value):
        self.form.DatasheetFontHeight = value

    @property
    def DatasheetFontItalic(self):
        return self.form.DatasheetFontItalic

    @DatasheetFontItalic.setter
    def DatasheetFontItalic(self, value):
        self.form.DatasheetFontItalic = value

    @property
    def DatasheetFontName(self):
        return self.form.DatasheetFontName

    @DatasheetFontName.setter
    def DatasheetFontName(self, value):
        self.form.DatasheetFontName = value

    @property
    def DatasheetFontUnderline(self):
        return self.form.DatasheetFontUnderline

    @DatasheetFontUnderline.setter
    def DatasheetFontUnderline(self, value):
        self.form.DatasheetFontUnderline = value

    @property
    def DatasheetFontWeight(self):
        return self.form.DatasheetFontWeight

    @DatasheetFontWeight.setter
    def DatasheetFontWeight(self, value):
        self.form.DatasheetFontWeight = value

    @property
    def DatasheetForeColor(self):
        return self.form.DatasheetForeColor

    @DatasheetForeColor.setter
    def DatasheetForeColor(self, value):
        self.form.DatasheetForeColor = value

    @property
    def DatasheetGridlinesBehavior(self):
        return self.form.DatasheetGridlinesBehavior

    @DatasheetGridlinesBehavior.setter
    def DatasheetGridlinesBehavior(self, value):
        self.form.DatasheetGridlinesBehavior = value

    @property
    def DatasheetGridlinesColor(self):
        return self.form.DatasheetGridlinesColor

    @DatasheetGridlinesColor.setter
    def DatasheetGridlinesColor(self, value):
        self.form.DatasheetGridlinesColor = value

    def DefaultControl(self, *args, ControlType=None):
        arguments = {"ControlType": ControlType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.form.DefaultControl):
            return self.form.DefaultControl(*args, **arguments)
        else:
            return self.form.GetDefaultControl(*args, **arguments)

    @property
    def DefaultView(self):
        return self.form.DefaultView

    @DefaultView.setter
    def DefaultView(self, value):
        self.form.DefaultView = value

    @property
    def Dirty(self):
        return self.form.Dirty

    @Dirty.setter
    def Dirty(self, value):
        self.form.Dirty = value

    @property
    def DisplayOnSharePointSite(self):
        return self.form.DisplayOnSharePointSite

    @DisplayOnSharePointSite.setter
    def DisplayOnSharePointSite(self, value):
        self.form.DisplayOnSharePointSite = value

    @property
    def DividingLines(self):
        return self.form.DividingLines

    @DividingLines.setter
    def DividingLines(self, value):
        self.form.DividingLines = value

    @property
    def FastLaserPrinting(self):
        return self.form.FastLaserPrinting

    @FastLaserPrinting.setter
    def FastLaserPrinting(self, value):
        self.form.FastLaserPrinting = value

    @property
    def FetchDefaults(self):
        return self.form.FetchDefaults

    @FetchDefaults.setter
    def FetchDefaults(self, value):
        self.form.FetchDefaults = value

    @property
    def Filter(self):
        return self.form.Filter

    @Filter.setter
    def Filter(self, value):
        self.form.Filter = value

    @property
    def FilterOn(self):
        return self.form.FilterOn

    @FilterOn.setter
    def FilterOn(self, value):
        self.form.FilterOn = value

    @property
    def FilterOnLoad(self):
        return self.form.FilterOnLoad

    @FilterOnLoad.setter
    def FilterOnLoad(self, value):
        self.form.FilterOnLoad = value

    @property
    def FitToScreen(self):
        return self.form.FitToScreen

    @FitToScreen.setter
    def FitToScreen(self, value):
        self.form.FitToScreen = value

    @property
    def Form(self):
        return self.form.Form

    @property
    def FrozenColumns(self):
        return self.form.FrozenColumns

    @FrozenColumns.setter
    def FrozenColumns(self, value):
        self.form.FrozenColumns = value

    @property
    def GridX(self):
        return self.form.GridX

    @GridX.setter
    def GridX(self, value):
        self.form.GridX = value

    @property
    def GridY(self):
        return self.form.GridY

    @GridY.setter
    def GridY(self, value):
        self.form.GridY = value

    @property
    def HasModule(self):
        return self.form.HasModule

    @HasModule.setter
    def HasModule(self, value):
        self.form.HasModule = value

    @property
    def HelpContextId(self):
        return self.form.HelpContextId

    @HelpContextId.setter
    def HelpContextId(self, value):
        self.form.HelpContextId = value

    @property
    def HelpFile(self):
        return self.form.HelpFile

    @HelpFile.setter
    def HelpFile(self, value):
        self.form.HelpFile = value

    @property
    def HorizontalDatasheetGridlineStyle(self):
        return self.form.HorizontalDatasheetGridlineStyle

    @HorizontalDatasheetGridlineStyle.setter
    def HorizontalDatasheetGridlineStyle(self, value):
        self.form.HorizontalDatasheetGridlineStyle = value

    @property
    def Hwnd(self):
        return self.form.Hwnd

    @Hwnd.setter
    def Hwnd(self, value):
        self.form.Hwnd = value

    @property
    def InsideHeight(self):
        return self.form.InsideHeight

    @InsideHeight.setter
    def InsideHeight(self, value):
        self.form.InsideHeight = value

    @property
    def InsideWidth(self):
        return self.form.InsideWidth

    @InsideWidth.setter
    def InsideWidth(self, value):
        self.form.InsideWidth = value

    @property
    def KeyPreview(self):
        return self.form.KeyPreview

    @KeyPreview.setter
    def KeyPreview(self, value):
        self.form.KeyPreview = value

    @property
    def LayoutForPrint(self):
        return self.form.LayoutForPrint

    @LayoutForPrint.setter
    def LayoutForPrint(self, value):
        self.form.LayoutForPrint = value

    @property
    def MaxRecButton(self):
        return self.form.MaxRecButton

    @MaxRecButton.setter
    def MaxRecButton(self, value):
        self.form.MaxRecButton = value

    @property
    def MaxRecords(self):
        return self.form.MaxRecords

    @MaxRecords.setter
    def MaxRecords(self, value):
        self.form.MaxRecords = value

    @property
    def MenuBar(self):
        return self.form.MenuBar

    @MenuBar.setter
    def MenuBar(self, value):
        self.form.MenuBar = value

    @property
    def MinMaxButtons(self):
        return self.form.MinMaxButtons

    @MinMaxButtons.setter
    def MinMaxButtons(self, value):
        self.form.MinMaxButtons = value

    @property
    def Modal(self):
        return self.form.Modal

    @Modal.setter
    def Modal(self, value):
        self.form.Modal = value

    @property
    def Module(self):
        return self.form.Module

    @property
    def MouseWheel(self):
        return self.form.MouseWheel

    @MouseWheel.setter
    def MouseWheel(self, value):
        self.form.MouseWheel = value

    @property
    def Moveable(self):
        return self.form.Moveable

    @Moveable.setter
    def Moveable(self, value):
        self.form.Moveable = value

    @property
    def Name(self):
        return self.form.Name

    @Name.setter
    def Name(self, value):
        self.form.Name = value

    @property
    def NavigationButtons(self):
        return self.form.NavigationButtons

    @NavigationButtons.setter
    def NavigationButtons(self, value):
        self.form.NavigationButtons = value

    @property
    def NavigationCaption(self):
        return self.form.NavigationCaption

    @NavigationCaption.setter
    def NavigationCaption(self, value):
        self.form.NavigationCaption = value

    @property
    def NewRecord(self):
        return self.form.NewRecord

    @property
    def OnActivate(self):
        return self.form.OnActivate

    @OnActivate.setter
    def OnActivate(self, value):
        self.form.OnActivate = value

    @property
    def OnApplyFilter(self):
        return self.form.OnApplyFilter

    @OnApplyFilter.setter
    def OnApplyFilter(self, value):
        self.form.OnApplyFilter = value

    @property
    def OnClick(self):
        return self.form.OnClick

    @OnClick.setter
    def OnClick(self, value):
        self.form.OnClick = value

    @property
    def OnClose(self):
        return self.form.OnClose

    @OnClose.setter
    def OnClose(self, value):
        self.form.OnClose = value

    @property
    def OnConnect(self):
        return self.form.OnConnect

    @OnConnect.setter
    def OnConnect(self, value):
        self.form.OnConnect = value

    @property
    def OnCurrent(self):
        return self.form.OnCurrent

    @OnCurrent.setter
    def OnCurrent(self, value):
        self.form.OnCurrent = value

    @property
    def OnDblClick(self):
        return self.form.OnDblClick

    @OnDblClick.setter
    def OnDblClick(self, value):
        self.form.OnDblClick = value

    @property
    def OnDeactivate(self):
        return self.form.OnDeactivate

    @OnDeactivate.setter
    def OnDeactivate(self, value):
        self.form.OnDeactivate = value

    @property
    def OnDelete(self):
        return self.form.OnDelete

    @OnDelete.setter
    def OnDelete(self, value):
        self.form.OnDelete = value

    @property
    def OnDirty(self):
        return self.form.OnDirty

    @OnDirty.setter
    def OnDirty(self, value):
        self.form.OnDirty = value

    @property
    def OnDisconnect(self):
        return self.form.OnDisconnect

    @OnDisconnect.setter
    def OnDisconnect(self, value):
        self.form.OnDisconnect = value

    @property
    def OnError(self):
        return self.form.OnError

    @OnError.setter
    def OnError(self, value):
        self.form.OnError = value

    @property
    def OnFilter(self):
        return self.form.OnFilter

    @OnFilter.setter
    def OnFilter(self, value):
        self.form.OnFilter = value

    @property
    def OnGotFocus(self):
        return self.form.OnGotFocus

    @OnGotFocus.setter
    def OnGotFocus(self, value):
        self.form.OnGotFocus = value

    @property
    def OnInsert(self):
        return self.form.OnInsert

    @OnInsert.setter
    def OnInsert(self, value):
        self.form.OnInsert = value

    @property
    def OnKeyDown(self):
        return self.form.OnKeyDown

    @OnKeyDown.setter
    def OnKeyDown(self, value):
        self.form.OnKeyDown = value

    @property
    def OnKeyPress(self):
        return self.form.OnKeyPress

    @OnKeyPress.setter
    def OnKeyPress(self, value):
        self.form.OnKeyPress = value

    @property
    def OnKeyUp(self):
        return self.form.OnKeyUp

    @OnKeyUp.setter
    def OnKeyUp(self, value):
        self.form.OnKeyUp = value

    @property
    def OnLoad(self):
        return self.form.OnLoad

    @OnLoad.setter
    def OnLoad(self, value):
        self.form.OnLoad = value

    @property
    def OnLostFocus(self):
        return self.form.OnLostFocus

    @OnLostFocus.setter
    def OnLostFocus(self, value):
        self.form.OnLostFocus = value

    @property
    def OnMouseDown(self):
        return self.form.OnMouseDown

    @OnMouseDown.setter
    def OnMouseDown(self, value):
        self.form.OnMouseDown = value

    @property
    def OnMouseMove(self):
        return self.form.OnMouseMove

    @OnMouseMove.setter
    def OnMouseMove(self, value):
        self.form.OnMouseMove = value

    @property
    def OnMouseUp(self):
        return self.form.OnMouseUp

    @OnMouseUp.setter
    def OnMouseUp(self, value):
        self.form.OnMouseUp = value

    @property
    def OnOpen(self):
        return self.form.OnOpen

    @OnOpen.setter
    def OnOpen(self, value):
        self.form.OnOpen = value

    @property
    def OnResize(self):
        return self.form.OnResize

    @OnResize.setter
    def OnResize(self, value):
        self.form.OnResize = value

    @property
    def OnTimer(self):
        return self.form.OnTimer

    @OnTimer.setter
    def OnTimer(self, value):
        self.form.OnTimer = value

    @property
    def OnUndo(self):
        return self.form.OnUndo

    @OnUndo.setter
    def OnUndo(self, value):
        self.form.OnUndo = value

    @property
    def OnUnload(self):
        return self.form.OnUnload

    @OnUnload.setter
    def OnUnload(self, value):
        self.form.OnUnload = value

    @property
    def OpenArgs(self):
        return self.form.OpenArgs

    @OpenArgs.setter
    def OpenArgs(self, value):
        self.form.OpenArgs = value

    @property
    def OrderBy(self):
        return self.form.OrderBy

    @OrderBy.setter
    def OrderBy(self, value):
        self.form.OrderBy = value

    @property
    def OrderByOn(self):
        return self.form.OrderByOn

    @OrderByOn.setter
    def OrderByOn(self, value):
        self.form.OrderByOn = value

    @property
    def OrderByOnLoad(self):
        return self.form.OrderByOnLoad

    @OrderByOnLoad.setter
    def OrderByOnLoad(self, value):
        self.form.OrderByOnLoad = value

    @property
    def Orientation(self):
        return self.form.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.form.Orientation = value

    @property
    def Page(self):
        return self.form.Page

    @Page.setter
    def Page(self, value):
        self.form.Page = value

    @property
    def Pages(self):
        return self.form.Pages

    @Pages.setter
    def Pages(self, value):
        self.form.Pages = value

    @property
    def Painting(self):
        return self.form.Painting

    @Painting.setter
    def Painting(self, value):
        self.form.Painting = value

    @property
    def PaintPalette(self):
        return self.form.PaintPalette

    @PaintPalette.setter
    def PaintPalette(self, value):
        self.form.PaintPalette = value

    @property
    def PaletteSource(self):
        return self.form.PaletteSource

    @PaletteSource.setter
    def PaletteSource(self, value):
        self.form.PaletteSource = value

    @property
    def Parent(self):
        return self.form.Parent

    @property
    def Picture(self):
        return self.form.Picture

    @Picture.setter
    def Picture(self, value):
        self.form.Picture = value

    @property
    def PictureAlignment(self):
        return self.form.PictureAlignment

    @PictureAlignment.setter
    def PictureAlignment(self, value):
        self.form.PictureAlignment = value

    @property
    def PictureData(self):
        return self.form.PictureData

    @PictureData.setter
    def PictureData(self, value):
        self.form.PictureData = value

    @property
    def PicturePalette(self):
        return self.form.PicturePalette

    @PicturePalette.setter
    def PicturePalette(self, value):
        self.form.PicturePalette = value

    @property
    def PictureSizeMode(self):
        return self.form.PictureSizeMode

    @PictureSizeMode.setter
    def PictureSizeMode(self, value):
        self.form.PictureSizeMode = value

    @property
    def PictureTiling(self):
        return self.form.PictureTiling

    @PictureTiling.setter
    def PictureTiling(self, value):
        self.form.PictureTiling = value

    @property
    def PictureType(self):
        return self.form.PictureType

    @PictureType.setter
    def PictureType(self, value):
        self.form.PictureType = value

    @property
    def PivotTable(self):
        return self.form.PivotTable

    @property
    def PivotTableChange(self):
        return self.form.PivotTableChange

    @PivotTableChange.setter
    def PivotTableChange(self, value):
        self.form.PivotTableChange = value

    @property
    def PopUp(self):
        return self.form.PopUp

    @PopUp.setter
    def PopUp(self, value):
        self.form.PopUp = value

    @property
    def Printer(self):
        return Printer(self.form.Printer)

    @Printer.setter
    def Printer(self, value):
        self.form.Printer = value

    @property
    def Properties(self):
        return Properties(self.form.Properties)

    @property
    def PrtDevMode(self):
        return self.form.PrtDevMode

    @PrtDevMode.setter
    def PrtDevMode(self, value):
        self.form.PrtDevMode = value

    @property
    def PrtDevNames(self):
        return self.form.PrtDevNames

    @PrtDevNames.setter
    def PrtDevNames(self, value):
        self.form.PrtDevNames = value

    @property
    def PrtMip(self):
        return self.form.PrtMip

    @property
    def Query(self):
        return self.form.Query

    @Query.setter
    def Query(self, value):
        self.form.Query = value

    @property
    def RecordLocks(self):
        return self.form.RecordLocks

    @RecordLocks.setter
    def RecordLocks(self, value):
        self.form.RecordLocks = value

    @property
    def RecordSelectors(self):
        return self.form.RecordSelectors

    @RecordSelectors.setter
    def RecordSelectors(self, value):
        self.form.RecordSelectors = value

    @property
    def Recordset(self):
        return self.form.Recordset

    @Recordset.setter
    def Recordset(self, value):
        self.form.Recordset = value

    @property
    def RecordsetClone(self):
        return self.form.RecordsetClone

    @property
    def RecordSource(self):
        return self.form.RecordSource

    @RecordSource.setter
    def RecordSource(self, value):
        self.form.RecordSource = value

    @property
    def RecordSourceQualifier(self):
        return self.form.RecordSourceQualifier

    @RecordSourceQualifier.setter
    def RecordSourceQualifier(self, value):
        self.form.RecordSourceQualifier = value

    @property
    def ResyncCommand(self):
        return self.form.ResyncCommand

    @ResyncCommand.setter
    def ResyncCommand(self, value):
        self.form.ResyncCommand = value

    @property
    def RibbonName(self):
        return self.form.RibbonName

    @RibbonName.setter
    def RibbonName(self, value):
        self.form.RibbonName = value

    @property
    def RowHeight(self):
        return self.form.RowHeight

    @RowHeight.setter
    def RowHeight(self, value):
        self.form.RowHeight = value

    @property
    def ScrollBars(self):
        return self.form.ScrollBars

    @ScrollBars.setter
    def ScrollBars(self, value):
        self.form.ScrollBars = value

    def Section(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.form.Section):
            return self.form.Section(*args, **arguments)
        else:
            return self.form.GetSection(*args, **arguments)

    @property
    def SelectionChange(self):
        return self.form.SelectionChange

    @SelectionChange.setter
    def SelectionChange(self, value):
        self.form.SelectionChange = value

    @property
    def SelHeight(self):
        return self.form.SelHeight

    @SelHeight.setter
    def SelHeight(self, value):
        self.form.SelHeight = value

    @property
    def SelLeft(self):
        return self.form.SelLeft

    @SelLeft.setter
    def SelLeft(self, value):
        self.form.SelLeft = value

    @property
    def SelTop(self):
        return self.form.SelTop

    @SelTop.setter
    def SelTop(self, value):
        self.form.SelTop = value

    @property
    def SelWidth(self):
        return self.form.SelWidth

    @SelWidth.setter
    def SelWidth(self, value):
        self.form.SelWidth = value

    @property
    def ServerFilter(self):
        return self.form.ServerFilter

    @ServerFilter.setter
    def ServerFilter(self, value):
        self.form.ServerFilter = value

    @property
    def ServerFilterByForm(self):
        return self.form.ServerFilterByForm

    @ServerFilterByForm.setter
    def ServerFilterByForm(self, value):
        self.form.ServerFilterByForm = value

    @property
    def ShortcutMenu(self):
        return self.form.ShortcutMenu

    @ShortcutMenu.setter
    def ShortcutMenu(self, value):
        self.form.ShortcutMenu = value

    @property
    def ShortcutMenuBar(self):
        return self.form.ShortcutMenuBar

    @ShortcutMenuBar.setter
    def ShortcutMenuBar(self, value):
        self.form.ShortcutMenuBar = value

    @property
    def SplitFormDatasheet(self):
        return self.form.SplitFormDatasheet

    @SplitFormDatasheet.setter
    def SplitFormDatasheet(self, value):
        self.form.SplitFormDatasheet = value

    @property
    def SplitFormOrientation(self):
        return self.form.SplitFormOrientation

    @SplitFormOrientation.setter
    def SplitFormOrientation(self, value):
        self.form.SplitFormOrientation = value

    @property
    def SplitFormPrinting(self):
        return self.form.SplitFormPrinting

    @SplitFormPrinting.setter
    def SplitFormPrinting(self, value):
        self.form.SplitFormPrinting = value

    @property
    def SplitFormSize(self):
        return self.form.SplitFormSize

    @SplitFormSize.setter
    def SplitFormSize(self, value):
        self.form.SplitFormSize = value

    @property
    def SplitFormSplitterBar(self):
        return self.form.SplitFormSplitterBar

    @SplitFormSplitterBar.setter
    def SplitFormSplitterBar(self, value):
        self.form.SplitFormSplitterBar = value

    @property
    def SplitFormSplitterBarSave(self):
        return self.form.SplitFormSplitterBarSave

    @SplitFormSplitterBarSave.setter
    def SplitFormSplitterBarSave(self, value):
        self.form.SplitFormSplitterBarSave = value

    @property
    def SubdatasheetExpanded(self):
        return self.form.SubdatasheetExpanded

    @SubdatasheetExpanded.setter
    def SubdatasheetExpanded(self, value):
        self.form.SubdatasheetExpanded = value

    @property
    def SubdatasheetHeight(self):
        return self.form.SubdatasheetHeight

    @SubdatasheetHeight.setter
    def SubdatasheetHeight(self, value):
        self.form.SubdatasheetHeight = value

    @property
    def Tag(self):
        return self.form.Tag

    @Tag.setter
    def Tag(self, value):
        self.form.Tag = value

    @property
    def TimerInterval(self):
        return self.form.TimerInterval

    @TimerInterval.setter
    def TimerInterval(self, value):
        self.form.TimerInterval = value

    @property
    def Toolbar(self):
        return self.form.Toolbar

    @Toolbar.setter
    def Toolbar(self, value):
        self.form.Toolbar = value

    @property
    def UniqueTable(self):
        return self.form.UniqueTable

    @property
    def UseDefaultPrinter(self):
        return self.form.UseDefaultPrinter

    @UseDefaultPrinter.setter
    def UseDefaultPrinter(self, value):
        self.form.UseDefaultPrinter = value

    @property
    def VerticalDatasheetGridlineStyle(self):
        return self.form.VerticalDatasheetGridlineStyle

    @VerticalDatasheetGridlineStyle.setter
    def VerticalDatasheetGridlineStyle(self, value):
        self.form.VerticalDatasheetGridlineStyle = value

    @property
    def ViewChange(self):
        return self.form.ViewChange

    @ViewChange.setter
    def ViewChange(self, value):
        self.form.ViewChange = value

    @property
    def ViewsAllowed(self):
        return self.form.ViewsAllowed

    @ViewsAllowed.setter
    def ViewsAllowed(self, value):
        self.form.ViewsAllowed = value

    @property
    def Visible(self):
        return self.form.Visible

    @Visible.setter
    def Visible(self, value):
        self.form.Visible = value

    @property
    def Width(self):
        return self.form.Width

    @Width.setter
    def Width(self, value):
        self.form.Width = value

    @property
    def WindowHeight(self):
        return self.form.WindowHeight

    @property
    def WindowLeft(self):
        return self.form.WindowLeft

    @property
    def WindowTop(self):
        return self.form.WindowTop

    @property
    def WindowWidth(self):
        return self.form.WindowWidth

    def GoToPage(self, *args, PageNumber=None, Right=None, Down=None):
        arguments = {"PageNumber": PageNumber, "Right": Right, "Down": Down}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.form.GoToPage(*args, **arguments)

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.form.Move(*args, **arguments)

    def Recalc(self):
        return self.form.Recalc()

    def Refresh(self):
        return self.form.Refresh()

    def Repaint(self):
        return self.form.Repaint()

    def Requery(self):
        self.form.Requery()

    def SetFocus(self):
        self.form.SetFocus()

    def Undo(self):
        self.form.Undo()

class FormatCondition:

    def __init__(self, formatcondition=None):
        self.formatcondition = formatcondition

    @property
    def BackColor(self):
        return self.formatcondition.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.formatcondition.BackColor = value

    @property
    def Enabled(self):
        return self.formatcondition.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.formatcondition.Enabled = value

    @property
    def Expression1(self):
        return self.formatcondition.Expression1

    @property
    def Expression2(self):
        return self.formatcondition.Expression2

    @property
    def FontBold(self):
        return self.formatcondition.FontBold

    @property
    def FontItalic(self):
        return self.formatcondition.FontItalic

    @property
    def FontUnderline(self):
        return self.formatcondition.FontUnderline

    @property
    def ForeColor(self):
        return self.formatcondition.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.formatcondition.ForeColor = value

    @property
    def LongestBarLimit(self):
        return self.formatcondition.LongestBarLimit

    @LongestBarLimit.setter
    def LongestBarLimit(self, value):
        self.formatcondition.LongestBarLimit = value

    @property
    def LongestBarValue(self):
        return self.formatcondition.LongestBarValue

    @LongestBarValue.setter
    def LongestBarValue(self, value):
        self.formatcondition.LongestBarValue = value

    @property
    def Operator(self):
        return self.formatcondition.Operator

    @property
    def ShortestBarLimit(self):
        return self.formatcondition.ShortestBarLimit

    @ShortestBarLimit.setter
    def ShortestBarLimit(self, value):
        self.formatcondition.ShortestBarLimit = value

    @property
    def ShortestBarValue(self):
        return self.formatcondition.ShortestBarValue

    @ShortestBarValue.setter
    def ShortestBarValue(self, value):
        self.formatcondition.ShortestBarValue = value

    @property
    def ShowBarOnly(self):
        return self.formatcondition.ShowBarOnly

    @ShowBarOnly.setter
    def ShowBarOnly(self, value):
        self.formatcondition.ShowBarOnly = value

    @property
    def Type(self):
        return AcFormatConditionType(self.formatcondition.Type)

    def Delete(self):
        return self.formatcondition.Delete()

    def Modify(self, *args, Type=None, Operator=None, Expression1=None, Expression2=None):
        arguments = {"Type": Type, "Operator": Operator, "Expression1": Expression1, "Expression2": Expression2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.formatcondition.Modify(*args, **arguments)

class FormatConditions:

    def __init__(self, formatconditions=None):
        self.formatconditions = formatconditions

    @property
    def Application(self):
        return self.formatconditions.Application

    @property
    def Count(self):
        return self.formatconditions.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.formatconditions.Item):
            return self.formatconditions.Item(*args, **arguments)
        else:
            return self.formatconditions.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.formatconditions.Parent

    def Add(self, *args, Type=None, Operator=None, Expression1=None, Expression2=None):
        arguments = {"Type": Type, "Operator": Operator, "Expression1": Expression1, "Expression2": Expression2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.formatconditions.Add(*args, **arguments)

    def Delete(self):
        self.formatconditions.Delete()

class Forms:

    def __init__(self, forms=None):
        self.forms = forms

    @property
    def Application(self):
        return self.forms.Application

    @property
    def Count(self):
        return self.forms.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.forms.Item):
            return self.forms.Item(*args, **arguments)
        else:
            return self.forms.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.forms.Parent

class GroupLevel:

    def __init__(self, grouplevel=None):
        self.grouplevel = grouplevel

    @property
    def Application(self):
        return self.grouplevel.Application

    @property
    def ControlSource(self):
        return self.grouplevel.ControlSource

    @ControlSource.setter
    def ControlSource(self, value):
        self.grouplevel.ControlSource = value

    @property
    def GroupFooter(self):
        return self.grouplevel.GroupFooter

    @GroupFooter.setter
    def GroupFooter(self, value):
        self.grouplevel.GroupFooter = value

    @property
    def GroupHeader(self):
        return self.grouplevel.GroupHeader

    @GroupHeader.setter
    def GroupHeader(self, value):
        self.grouplevel.GroupHeader = value

    @property
    def GroupInterval(self):
        return self.grouplevel.GroupInterval

    @GroupInterval.setter
    def GroupInterval(self, value):
        self.grouplevel.GroupInterval = value

    @property
    def GroupOn(self):
        return self.grouplevel.GroupOn

    @GroupOn.setter
    def GroupOn(self, value):
        self.grouplevel.GroupOn = value

    @property
    def KeepTogether(self):
        return self.grouplevel.KeepTogether

    @KeepTogether.setter
    def KeepTogether(self, value):
        self.grouplevel.KeepTogether = value

    @property
    def Parent(self):
        return self.grouplevel.Parent

    @property
    def Properties(self):
        return Properties(self.grouplevel.Properties)

    @property
    def SortOrder(self):
        return self.grouplevel.SortOrder

    @SortOrder.setter
    def SortOrder(self, value):
        self.grouplevel.SortOrder = value

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
    def EmailSubject(self):
        return self.hyperlink.EmailSubject

    @EmailSubject.setter
    def EmailSubject(self, value):
        self.hyperlink.EmailSubject = value

    @property
    def ScreenTip(self):
        return self.hyperlink.ScreenTip

    @ScreenTip.setter
    def ScreenTip(self, value):
        self.hyperlink.ScreenTip = value

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

    def AddToFavorites(self):
        return self.hyperlink.AddToFavorites()

    def CreateNewDocument(self, *args, FileName=None, EditNow=None, Overwrite=None):
        arguments = {"FileName": FileName, "EditNow": EditNow, "Overwrite": Overwrite}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.hyperlink.CreateNewDocument(*args, **arguments)

    def Follow(self, *args, NewWindow=None, AddHistory=None, ExtraInfo=None, Method=None, HeaderInfo=None):
        arguments = {"NewWindow": NewWindow, "AddHistory": AddHistory, "ExtraInfo": ExtraInfo, "Method": Method, "HeaderInfo": HeaderInfo}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.hyperlink.Follow(*args, **arguments)

class Image:

    def __init__(self, image=None):
        self.image = image

    @property
    def Application(self):
        return self.image.Application

    @property
    def BackColor(self):
        return self.image.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.image.BackColor = value

    @property
    def BackShade(self):
        return self.image.BackShade

    @BackShade.setter
    def BackShade(self, value):
        self.image.BackShade = value

    @property
    def BackStyle(self):
        return self.image.BackStyle

    @BackStyle.setter
    def BackStyle(self, value):
        self.image.BackStyle = value

    @property
    def BackThemeColorIndex(self):
        return self.image.BackThemeColorIndex

    @BackThemeColorIndex.setter
    def BackThemeColorIndex(self, value):
        self.image.BackThemeColorIndex = value

    @property
    def BackTint(self):
        return self.image.BackTint

    @BackTint.setter
    def BackTint(self, value):
        self.image.BackTint = value

    @property
    def BorderColor(self):
        return self.image.BorderColor

    @BorderColor.setter
    def BorderColor(self, value):
        self.image.BorderColor = value

    @property
    def BorderShade(self):
        return self.image.BorderShade

    @BorderShade.setter
    def BorderShade(self, value):
        self.image.BorderShade = value

    @property
    def BorderStyle(self):
        return self.image.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.image.BorderStyle = value

    @property
    def BorderThemeColorIndex(self):
        return self.image.BorderThemeColorIndex

    @BorderThemeColorIndex.setter
    def BorderThemeColorIndex(self, value):
        self.image.BorderThemeColorIndex = value

    @property
    def BorderTint(self):
        return self.image.BorderTint

    @BorderTint.setter
    def BorderTint(self, value):
        self.image.BorderTint = value

    @property
    def BorderWidth(self):
        return self.image.BorderWidth

    @BorderWidth.setter
    def BorderWidth(self, value):
        self.image.BorderWidth = value

    @property
    def BottomPadding(self):
        return self.image.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.image.BottomPadding = value

    @property
    def Controls(self):
        return Controls(self.image.Controls)

    @property
    def ControlTipText(self):
        return self.image.ControlTipText

    @ControlTipText.setter
    def ControlTipText(self, value):
        self.image.ControlTipText = value

    @property
    def ControlType(self):
        return self.image.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.image.ControlType = value

    @property
    def DisplayWhen(self):
        return self.image.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.image.DisplayWhen = value

    @property
    def EventProcPrefix(self):
        return self.image.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.image.EventProcPrefix = value

    @property
    def GridlineColor(self):
        return self.image.GridlineColor

    @GridlineColor.setter
    def GridlineColor(self, value):
        self.image.GridlineColor = value

    @property
    def GridlineShade(self):
        return self.image.GridlineShade

    @GridlineShade.setter
    def GridlineShade(self, value):
        self.image.GridlineShade = value

    @property
    def GridlineStyleBottom(self):
        return self.image.GridlineStyleBottom

    @GridlineStyleBottom.setter
    def GridlineStyleBottom(self, value):
        self.image.GridlineStyleBottom = value

    @property
    def GridlineStyleLeft(self):
        return self.image.GridlineStyleLeft

    @GridlineStyleLeft.setter
    def GridlineStyleLeft(self, value):
        self.image.GridlineStyleLeft = value

    @property
    def GridlineStyleRight(self):
        return self.image.GridlineStyleRight

    @GridlineStyleRight.setter
    def GridlineStyleRight(self, value):
        self.image.GridlineStyleRight = value

    @property
    def GridlineStyleTop(self):
        return self.image.GridlineStyleTop

    @GridlineStyleTop.setter
    def GridlineStyleTop(self, value):
        self.image.GridlineStyleTop = value

    @property
    def GridlineThemeColorIndex(self):
        return self.image.GridlineThemeColorIndex

    @GridlineThemeColorIndex.setter
    def GridlineThemeColorIndex(self, value):
        self.image.GridlineThemeColorIndex = value

    @property
    def GridlineTint(self):
        return self.image.GridlineTint

    @GridlineTint.setter
    def GridlineTint(self, value):
        self.image.GridlineTint = value

    @property
    def GridlineWidthBottom(self):
        return self.image.GridlineWidthBottom

    @GridlineWidthBottom.setter
    def GridlineWidthBottom(self, value):
        self.image.GridlineWidthBottom = value

    @property
    def GridlineWidthLeft(self):
        return self.image.GridlineWidthLeft

    @GridlineWidthLeft.setter
    def GridlineWidthLeft(self, value):
        self.image.GridlineWidthLeft = value

    @property
    def GridlineWidthRight(self):
        return self.image.GridlineWidthRight

    @GridlineWidthRight.setter
    def GridlineWidthRight(self, value):
        self.image.GridlineWidthRight = value

    @property
    def GridlineWidthTop(self):
        return self.image.GridlineWidthTop

    @GridlineWidthTop.setter
    def GridlineWidthTop(self, value):
        self.image.GridlineWidthTop = value

    @property
    def Height(self):
        return self.image.Height

    @Height.setter
    def Height(self, value):
        self.image.Height = value

    @property
    def HelpContextId(self):
        return self.image.HelpContextId

    @HelpContextId.setter
    def HelpContextId(self, value):
        self.image.HelpContextId = value

    @property
    def HorizontalAnchor(self):
        return self.image.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.image.HorizontalAnchor = value

    @property
    def Hyperlink(self):
        return self.image.Hyperlink

    @property
    def HyperlinkAddress(self):
        return self.image.HyperlinkAddress

    @HyperlinkAddress.setter
    def HyperlinkAddress(self, value):
        self.image.HyperlinkAddress = value

    @property
    def HyperlinkSubAddress(self):
        return self.image.HyperlinkSubAddress

    @HyperlinkSubAddress.setter
    def HyperlinkSubAddress(self, value):
        self.image.HyperlinkSubAddress = value

    @property
    def ImageHeight(self):
        return self.image.ImageHeight

    @ImageHeight.setter
    def ImageHeight(self, value):
        self.image.ImageHeight = value

    @property
    def ImageWidth(self):
        return self.image.ImageWidth

    @ImageWidth.setter
    def ImageWidth(self, value):
        self.image.ImageWidth = value

    @property
    def InSelection(self):
        return self.image.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.image.InSelection = value

    @property
    def IsVisible(self):
        return self.image.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.image.IsVisible = value

    @property
    def Layout(self):
        return AcLayoutType(self.image.Layout)

    @property
    def LayoutID(self):
        return self.image.LayoutID

    @property
    def Left(self):
        return self.image.Left

    @Left.setter
    def Left(self, value):
        self.image.Left = value

    @property
    def LeftPadding(self):
        return self.image.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.image.LeftPadding = value

    @property
    def Name(self):
        return self.image.Name

    @Name.setter
    def Name(self, value):
        self.image.Name = value

    @property
    def ObjectPalette(self):
        return self.image.ObjectPalette

    @ObjectPalette.setter
    def ObjectPalette(self, value):
        self.image.ObjectPalette = value

    @property
    def OldBorderStyle(self):
        return self.image.OldBorderStyle

    @OldBorderStyle.setter
    def OldBorderStyle(self, value):
        self.image.OldBorderStyle = value

    @property
    def OldValue(self):
        return self.image.OldValue

    @property
    def OnClick(self):
        return self.image.OnClick

    @OnClick.setter
    def OnClick(self, value):
        self.image.OnClick = value

    @property
    def OnDblClick(self):
        return self.image.OnDblClick

    @OnDblClick.setter
    def OnDblClick(self, value):
        self.image.OnDblClick = value

    @property
    def OnMouseDown(self):
        return self.image.OnMouseDown

    @OnMouseDown.setter
    def OnMouseDown(self, value):
        self.image.OnMouseDown = value

    @property
    def OnMouseMove(self):
        return self.image.OnMouseMove

    @OnMouseMove.setter
    def OnMouseMove(self, value):
        self.image.OnMouseMove = value

    @property
    def OnMouseUp(self):
        return self.image.OnMouseUp

    @OnMouseUp.setter
    def OnMouseUp(self, value):
        self.image.OnMouseUp = value

    @property
    def Parent(self):
        return self.image.Parent

    @property
    def Picture(self):
        return self.image.Picture

    @Picture.setter
    def Picture(self, value):
        self.image.Picture = value

    @property
    def PictureAlignment(self):
        return self.image.PictureAlignment

    @PictureAlignment.setter
    def PictureAlignment(self, value):
        self.image.PictureAlignment = value

    @property
    def PictureData(self):
        return self.image.PictureData

    @PictureData.setter
    def PictureData(self, value):
        self.image.PictureData = value

    @property
    def PictureTiling(self):
        return self.image.PictureTiling

    @PictureTiling.setter
    def PictureTiling(self, value):
        self.image.PictureTiling = value

    @property
    def PictureType(self):
        return self.image.PictureType

    @PictureType.setter
    def PictureType(self, value):
        self.image.PictureType = value

    @property
    def Properties(self):
        return Properties(self.image.Properties)

    @property
    def RightPadding(self):
        return self.image.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.image.RightPadding = value

    @property
    def Section(self):
        return self.image.Section

    @Section.setter
    def Section(self, value):
        self.image.Section = value

    @property
    def ShortcutMenuBar(self):
        return self.image.ShortcutMenuBar

    @ShortcutMenuBar.setter
    def ShortcutMenuBar(self, value):
        self.image.ShortcutMenuBar = value

    @property
    def SizeMode(self):
        return self.image.SizeMode

    @property
    def SpecialEffect(self):
        return self.image.SpecialEffect

    @SpecialEffect.setter
    def SpecialEffect(self, value):
        self.image.SpecialEffect = value

    @property
    def Tag(self):
        return self.image.Tag

    @Tag.setter
    def Tag(self, value):
        self.image.Tag = value

    @property
    def Top(self):
        return self.image.Top

    @Top.setter
    def Top(self, value):
        self.image.Top = value

    @property
    def TopPadding(self):
        return self.image.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.image.TopPadding = value

    @property
    def VerticalAnchor(self):
        return self.image.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.image.VerticalAnchor = value

    @property
    def Visible(self):
        return self.image.Visible

    @Visible.setter
    def Visible(self, value):
        self.image.Visible = value

    @property
    def Width(self):
        return self.image.Width

    @Width.setter
    def Width(self, value):
        self.image.Width = value

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.image.Move(*args, **arguments)

    def Requery(self):
        self.image.Requery()

    def SetFocus(self):
        return self.image.SetFocus()

    def SizeToFit(self):
        self.image.SizeToFit()

class ImportExportSpecification:

    def __init__(self, importexportspecification=None):
        self.importexportspecification = importexportspecification

    @property
    def Application(self):
        return self.importexportspecification.Application

    @property
    def Description(self):
        return self.importexportspecification.Description

    @Description.setter
    def Description(self, value):
        self.importexportspecification.Description = value

    @property
    def Name(self):
        return self.importexportspecification.Name

    @Name.setter
    def Name(self, value):
        self.importexportspecification.Name = value

    @property
    def Parent(self):
        return self.importexportspecification.Parent

    @property
    def XML(self):
        return self.importexportspecification.XML

    @XML.setter
    def XML(self, value):
        self.importexportspecification.XML = value

    def Delete(self):
        self.importexportspecification.Delete()

    def Execute(self, *args, OverwritePrompt=None):
        arguments = {"OverwritePrompt": OverwritePrompt}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.importexportspecification.Execute(*args, **arguments)

class ImportExportSpecifications:

    def __init__(self, importexportspecifications=None):
        self.importexportspecifications = importexportspecifications

    @property
    def Application(self):
        return self.importexportspecifications.Application

    @property
    def Count(self):
        return self.importexportspecifications.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.importexportspecifications.Item):
            return self.importexportspecifications.Item(*args, **arguments)
        else:
            return self.importexportspecifications.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.importexportspecifications.Parent

    def Add(self, *args, Name=None, SpecificationDefinition=None):
        arguments = {"Name": Name, "SpecificationDefinition": SpecificationDefinition}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.importexportspecifications.Add(*args, **arguments)

class Label:

    def __init__(self, label=None):
        self.label = label

    @property
    def Application(self):
        return self.label.Application

    @property
    def BackColor(self):
        return self.label.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.label.BackColor = value

    @property
    def BackShade(self):
        return self.label.BackShade

    @BackShade.setter
    def BackShade(self, value):
        self.label.BackShade = value

    @property
    def BackStyle(self):
        return self.label.BackStyle

    @BackStyle.setter
    def BackStyle(self, value):
        self.label.BackStyle = value

    @property
    def BackThemeColorIndex(self):
        return self.label.BackThemeColorIndex

    @BackThemeColorIndex.setter
    def BackThemeColorIndex(self, value):
        self.label.BackThemeColorIndex = value

    @property
    def BackTint(self):
        return self.label.BackTint

    @BackTint.setter
    def BackTint(self, value):
        self.label.BackTint = value

    @property
    def BorderColor(self):
        return self.label.BorderColor

    @BorderColor.setter
    def BorderColor(self, value):
        self.label.BorderColor = value

    @property
    def BorderShade(self):
        return self.label.BorderShade

    @BorderShade.setter
    def BorderShade(self, value):
        self.label.BorderShade = value

    @property
    def BorderStyle(self):
        return self.label.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.label.BorderStyle = value

    @property
    def BorderThemeColorIndex(self):
        return self.label.BorderThemeColorIndex

    @BorderThemeColorIndex.setter
    def BorderThemeColorIndex(self, value):
        self.label.BorderThemeColorIndex = value

    @property
    def BorderTint(self):
        return self.label.BorderTint

    @BorderTint.setter
    def BorderTint(self, value):
        self.label.BorderTint = value

    @property
    def BorderWidth(self):
        return self.label.BorderWidth

    @BorderWidth.setter
    def BorderWidth(self, value):
        self.label.BorderWidth = value

    @property
    def BottomMargin(self):
        return self.label.BottomMargin

    @BottomMargin.setter
    def BottomMargin(self, value):
        self.label.BottomMargin = value

    @property
    def BottomPadding(self):
        return self.label.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.label.BottomPadding = value

    @property
    def Caption(self):
        return self.label.Caption

    @Caption.setter
    def Caption(self, value):
        self.label.Caption = value

    @property
    def ControlTipText(self):
        return self.label.ControlTipText

    @ControlTipText.setter
    def ControlTipText(self, value):
        self.label.ControlTipText = value

    @property
    def ControlType(self):
        return self.label.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.label.ControlType = value

    @property
    def DisplayWhen(self):
        return self.label.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.label.DisplayWhen = value

    @property
    def EventProcPrefix(self):
        return self.label.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.label.EventProcPrefix = value

    @property
    def FontBold(self):
        return self.label.FontBold

    @property
    def FontItalic(self):
        return self.label.FontItalic

    @property
    def FontName(self):
        return self.label.FontName

    @property
    def FontSize(self):
        return self.label.FontSize

    @property
    def FontUnderline(self):
        return self.label.FontUnderline

    @property
    def FontWeight(self):
        return self.label.FontWeight

    @FontWeight.setter
    def FontWeight(self, value):
        self.label.FontWeight = value

    @property
    def ForeColor(self):
        return self.label.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.label.ForeColor = value

    @property
    def ForeShade(self):
        return self.label.ForeShade

    @ForeShade.setter
    def ForeShade(self, value):
        self.label.ForeShade = value

    @property
    def ForeThemeColorIndex(self):
        return self.label.ForeThemeColorIndex

    @ForeThemeColorIndex.setter
    def ForeThemeColorIndex(self, value):
        self.label.ForeThemeColorIndex = value

    @property
    def ForeTint(self):
        return self.label.ForeTint

    @ForeTint.setter
    def ForeTint(self, value):
        self.label.ForeTint = value

    @property
    def GridlineColor(self):
        return self.label.GridlineColor

    @GridlineColor.setter
    def GridlineColor(self, value):
        self.label.GridlineColor = value

    @property
    def GridlineShade(self):
        return self.label.GridlineShade

    @GridlineShade.setter
    def GridlineShade(self, value):
        self.label.GridlineShade = value

    @property
    def GridlineStyleBottom(self):
        return self.label.GridlineStyleBottom

    @GridlineStyleBottom.setter
    def GridlineStyleBottom(self, value):
        self.label.GridlineStyleBottom = value

    @property
    def GridlineStyleLeft(self):
        return self.label.GridlineStyleLeft

    @GridlineStyleLeft.setter
    def GridlineStyleLeft(self, value):
        self.label.GridlineStyleLeft = value

    @property
    def GridlineStyleRight(self):
        return self.label.GridlineStyleRight

    @GridlineStyleRight.setter
    def GridlineStyleRight(self, value):
        self.label.GridlineStyleRight = value

    @property
    def GridlineStyleTop(self):
        return self.label.GridlineStyleTop

    @GridlineStyleTop.setter
    def GridlineStyleTop(self, value):
        self.label.GridlineStyleTop = value

    @property
    def GridlineThemeColorIndex(self):
        return self.label.GridlineThemeColorIndex

    @GridlineThemeColorIndex.setter
    def GridlineThemeColorIndex(self, value):
        self.label.GridlineThemeColorIndex = value

    @property
    def GridlineTint(self):
        return self.label.GridlineTint

    @GridlineTint.setter
    def GridlineTint(self, value):
        self.label.GridlineTint = value

    @property
    def GridlineWidthBottom(self):
        return self.label.GridlineWidthBottom

    @GridlineWidthBottom.setter
    def GridlineWidthBottom(self, value):
        self.label.GridlineWidthBottom = value

    @property
    def GridlineWidthLeft(self):
        return self.label.GridlineWidthLeft

    @GridlineWidthLeft.setter
    def GridlineWidthLeft(self, value):
        self.label.GridlineWidthLeft = value

    @property
    def GridlineWidthRight(self):
        return self.label.GridlineWidthRight

    @GridlineWidthRight.setter
    def GridlineWidthRight(self, value):
        self.label.GridlineWidthRight = value

    @property
    def GridlineWidthTop(self):
        return self.label.GridlineWidthTop

    @GridlineWidthTop.setter
    def GridlineWidthTop(self, value):
        self.label.GridlineWidthTop = value

    @property
    def Height(self):
        return self.label.Height

    @Height.setter
    def Height(self, value):
        self.label.Height = value

    @property
    def HelpContextId(self):
        return self.label.HelpContextId

    @HelpContextId.setter
    def HelpContextId(self, value):
        self.label.HelpContextId = value

    @property
    def HorizontalAnchor(self):
        return self.label.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.label.HorizontalAnchor = value

    @property
    def Hyperlink(self):
        return self.label.Hyperlink

    @property
    def HyperlinkAddress(self):
        return self.label.HyperlinkAddress

    @HyperlinkAddress.setter
    def HyperlinkAddress(self, value):
        self.label.HyperlinkAddress = value

    @property
    def HyperlinkSubAddress(self):
        return self.label.HyperlinkSubAddress

    @HyperlinkSubAddress.setter
    def HyperlinkSubAddress(self, value):
        self.label.HyperlinkSubAddress = value

    @property
    def InSelection(self):
        return self.label.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.label.InSelection = value

    @property
    def IsVisible(self):
        return self.label.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.label.IsVisible = value

    @property
    def Layout(self):
        return AcLayoutType(self.label.Layout)

    @property
    def LayoutID(self):
        return self.label.LayoutID

    @property
    def Left(self):
        return self.label.Left

    @Left.setter
    def Left(self, value):
        self.label.Left = value

    @property
    def LeftMargin(self):
        return self.label.LeftMargin

    @LeftMargin.setter
    def LeftMargin(self, value):
        self.label.LeftMargin = value

    @property
    def LeftPadding(self):
        return self.label.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.label.LeftPadding = value

    @property
    def LineSpacing(self):
        return self.label.LineSpacing

    @LineSpacing.setter
    def LineSpacing(self, value):
        self.label.LineSpacing = value

    @property
    def Name(self):
        return self.label.Name

    @Name.setter
    def Name(self, value):
        self.label.Name = value

    @property
    def OldBorderStyle(self):
        return self.label.OldBorderStyle

    @OldBorderStyle.setter
    def OldBorderStyle(self, value):
        self.label.OldBorderStyle = value

    @property
    def OnClick(self):
        return self.label.OnClick

    @OnClick.setter
    def OnClick(self, value):
        self.label.OnClick = value

    @property
    def OnDblClick(self):
        return self.label.OnDblClick

    @OnDblClick.setter
    def OnDblClick(self, value):
        self.label.OnDblClick = value

    @property
    def OnMouseDown(self):
        return self.label.OnMouseDown

    @OnMouseDown.setter
    def OnMouseDown(self, value):
        self.label.OnMouseDown = value

    @property
    def OnMouseMove(self):
        return self.label.OnMouseMove

    @OnMouseMove.setter
    def OnMouseMove(self, value):
        self.label.OnMouseMove = value

    @property
    def OnMouseUp(self):
        return self.label.OnMouseUp

    @OnMouseUp.setter
    def OnMouseUp(self, value):
        self.label.OnMouseUp = value

    @property
    def Parent(self):
        return self.label.Parent

    @property
    def Properties(self):
        return Properties(self.label.Properties)

    @property
    def ReadingOrder(self):
        return self.label.ReadingOrder

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.label.ReadingOrder = value

    @property
    def RightMargin(self):
        return self.label.RightMargin

    @RightMargin.setter
    def RightMargin(self, value):
        self.label.RightMargin = value

    @property
    def RightPadding(self):
        return self.label.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.label.RightPadding = value

    @property
    def Section(self):
        return self.label.Section

    @Section.setter
    def Section(self, value):
        self.label.Section = value

    @property
    def ShortcutMenuBar(self):
        return self.label.ShortcutMenuBar

    @ShortcutMenuBar.setter
    def ShortcutMenuBar(self, value):
        self.label.ShortcutMenuBar = value

    @property
    def SmartTags(self):
        return SmartTags(self.label.SmartTags)

    @property
    def SpecialEffect(self):
        return self.label.SpecialEffect

    @SpecialEffect.setter
    def SpecialEffect(self, value):
        self.label.SpecialEffect = value

    @property
    def Tag(self):
        return self.label.Tag

    @Tag.setter
    def Tag(self, value):
        self.label.Tag = value

    @property
    def TextAlign(self):
        return self.label.TextAlign

    @TextAlign.setter
    def TextAlign(self, value):
        self.label.TextAlign = value

    @property
    def ThemeFontIndex(self):
        return self.label.ThemeFontIndex

    @ThemeFontIndex.setter
    def ThemeFontIndex(self, value):
        self.label.ThemeFontIndex = value

    @property
    def Top(self):
        return self.label.Top

    @Top.setter
    def Top(self, value):
        self.label.Top = value

    @property
    def TopMargin(self):
        return self.label.TopMargin

    @TopMargin.setter
    def TopMargin(self, value):
        self.label.TopMargin = value

    @property
    def TopPadding(self):
        return self.label.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.label.TopPadding = value

    @property
    def Vertical(self):
        return self.label.Vertical

    @Vertical.setter
    def Vertical(self, value):
        self.label.Vertical = value

    @property
    def VerticalAnchor(self):
        return self.label.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.label.VerticalAnchor = value

    @property
    def Visible(self):
        return self.label.Visible

    @Visible.setter
    def Visible(self, value):
        self.label.Visible = value

    @property
    def Width(self):
        return self.label.Width

    @Width.setter
    def Width(self, value):
        self.label.Width = value

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.label.Move(*args, **arguments)

    def SizeToFit(self):
        self.label.SizeToFit()

class Line:

    def __init__(self, line=None):
        self.line = line

    @property
    def Application(self):
        return self.line.Application

    @property
    def BorderColor(self):
        return self.line.BorderColor

    @BorderColor.setter
    def BorderColor(self, value):
        self.line.BorderColor = value

    @property
    def BorderShade(self):
        return self.line.BorderShade

    @BorderShade.setter
    def BorderShade(self, value):
        self.line.BorderShade = value

    @property
    def BorderStyle(self):
        return self.line.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.line.BorderStyle = value

    @property
    def BorderThemeColorIndex(self):
        return self.line.BorderThemeColorIndex

    @BorderThemeColorIndex.setter
    def BorderThemeColorIndex(self, value):
        self.line.BorderThemeColorIndex = value

    @property
    def BorderTint(self):
        return self.line.BorderTint

    @BorderTint.setter
    def BorderTint(self, value):
        self.line.BorderTint = value

    @property
    def BorderWidth(self):
        return self.line.BorderWidth

    @BorderWidth.setter
    def BorderWidth(self, value):
        self.line.BorderWidth = value

    @property
    def ControlType(self):
        return self.line.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.line.ControlType = value

    @property
    def DisplayWhen(self):
        return self.line.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.line.DisplayWhen = value

    @property
    def EventProcPrefix(self):
        return self.line.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.line.EventProcPrefix = value

    @property
    def Height(self):
        return self.line.Height

    @Height.setter
    def Height(self, value):
        self.line.Height = value

    @property
    def HorizontalAnchor(self):
        return self.line.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.line.HorizontalAnchor = value

    @property
    def InSelection(self):
        return self.line.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.line.InSelection = value

    @property
    def IsVisible(self):
        return self.line.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.line.IsVisible = value

    @property
    def Left(self):
        return self.line.Left

    @Left.setter
    def Left(self, value):
        self.line.Left = value

    @property
    def LineSlant(self):
        return self.line.LineSlant

    @LineSlant.setter
    def LineSlant(self, value):
        self.line.LineSlant = value

    @property
    def Name(self):
        return self.line.Name

    @Name.setter
    def Name(self, value):
        self.line.Name = value

    @property
    def OldBorderStyle(self):
        return self.line.OldBorderStyle

    @OldBorderStyle.setter
    def OldBorderStyle(self, value):
        self.line.OldBorderStyle = value

    @property
    def Parent(self):
        return self.line.Parent

    @property
    def Properties(self):
        return Properties(self.line.Properties)

    @property
    def Section(self):
        return self.line.Section

    @Section.setter
    def Section(self, value):
        self.line.Section = value

    @property
    def SpecialEffect(self):
        return self.line.SpecialEffect

    @SpecialEffect.setter
    def SpecialEffect(self, value):
        self.line.SpecialEffect = value

    @property
    def Tag(self):
        return self.line.Tag

    @Tag.setter
    def Tag(self, value):
        self.line.Tag = value

    @property
    def Top(self):
        return self.line.Top

    @Top.setter
    def Top(self, value):
        self.line.Top = value

    @property
    def VerticalAnchor(self):
        return self.line.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.line.VerticalAnchor = value

    @property
    def Visible(self):
        return self.line.Visible

    @Visible.setter
    def Visible(self, value):
        self.line.Visible = value

    @property
    def Width(self):
        return self.line.Width

    @Width.setter
    def Width(self, value):
        self.line.Width = value

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.line.Move(*args, **arguments)

    def SizeToFit(self):
        self.line.SizeToFit()

class ListBox:

    def __init__(self, listbox=None):
        self.listbox = listbox

    @property
    def AddColon(self):
        return self.listbox.AddColon

    @AddColon.setter
    def AddColon(self, value):
        self.listbox.AddColon = value

    @property
    def AllowValueListEdits(self):
        return self.listbox.AllowValueListEdits

    @AllowValueListEdits.setter
    def AllowValueListEdits(self, value):
        self.listbox.AllowValueListEdits = value

    @property
    def Application(self):
        return self.listbox.Application

    @property
    def AutoLabel(self):
        return self.listbox.AutoLabel

    @AutoLabel.setter
    def AutoLabel(self, value):
        self.listbox.AutoLabel = value

    @property
    def BackColor(self):
        return self.listbox.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.listbox.BackColor = value

    @property
    def BackShade(self):
        return self.listbox.BackShade

    @BackShade.setter
    def BackShade(self, value):
        self.listbox.BackShade = value

    @property
    def BackThemeColorIndex(self):
        return self.listbox.BackThemeColorIndex

    @BackThemeColorIndex.setter
    def BackThemeColorIndex(self, value):
        self.listbox.BackThemeColorIndex = value

    @property
    def BackTint(self):
        return self.listbox.BackTint

    @BackTint.setter
    def BackTint(self, value):
        self.listbox.BackTint = value

    @property
    def BorderColor(self):
        return self.listbox.BorderColor

    @BorderColor.setter
    def BorderColor(self, value):
        self.listbox.BorderColor = value

    @property
    def BorderShade(self):
        return self.listbox.BorderShade

    @BorderShade.setter
    def BorderShade(self, value):
        self.listbox.BorderShade = value

    @property
    def BorderStyle(self):
        return self.listbox.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.listbox.BorderStyle = value

    @property
    def BorderThemeColorIndex(self):
        return self.listbox.BorderThemeColorIndex

    @BorderThemeColorIndex.setter
    def BorderThemeColorIndex(self, value):
        self.listbox.BorderThemeColorIndex = value

    @property
    def BorderTint(self):
        return self.listbox.BorderTint

    @BorderTint.setter
    def BorderTint(self, value):
        self.listbox.BorderTint = value

    @property
    def BorderWidth(self):
        return self.listbox.BorderWidth

    @BorderWidth.setter
    def BorderWidth(self, value):
        self.listbox.BorderWidth = value

    @property
    def BottomPadding(self):
        return self.listbox.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.listbox.BottomPadding = value

    @property
    def BoundColumn(self):
        return self.listbox.BoundColumn

    @BoundColumn.setter
    def BoundColumn(self, value):
        self.listbox.BoundColumn = value

    def Column(self, *args, Index=None, Row=None):
        arguments = {"Index": Index, "Row": Row}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.listbox.Column):
            return self.listbox.Column(*args, **arguments)
        else:
            return self.listbox.GetColumn(*args, **arguments)

    @property
    def ColumnCount(self):
        return self.listbox.ColumnCount

    @ColumnCount.setter
    def ColumnCount(self, value):
        self.listbox.ColumnCount = value

    @property
    def ColumnHeads(self):
        return self.listbox.ColumnHeads

    @ColumnHeads.setter
    def ColumnHeads(self, value):
        self.listbox.ColumnHeads = value

    @property
    def ColumnHidden(self):
        return self.listbox.ColumnHidden

    @ColumnHidden.setter
    def ColumnHidden(self, value):
        self.listbox.ColumnHidden = value

    @property
    def ColumnOrder(self):
        return self.listbox.ColumnOrder

    @ColumnOrder.setter
    def ColumnOrder(self, value):
        self.listbox.ColumnOrder = value

    @property
    def ColumnWidth(self):
        return self.listbox.ColumnWidth

    @ColumnWidth.setter
    def ColumnWidth(self, value):
        self.listbox.ColumnWidth = value

    @property
    def ColumnWidths(self):
        return self.listbox.ColumnWidths

    @ColumnWidths.setter
    def ColumnWidths(self, value):
        self.listbox.ColumnWidths = value

    @property
    def Controls(self):
        return Controls(self.listbox.Controls)

    @property
    def ControlSource(self):
        return self.listbox.ControlSource

    @ControlSource.setter
    def ControlSource(self, value):
        self.listbox.ControlSource = value

    @property
    def ControlTipText(self):
        return self.listbox.ControlTipText

    @ControlTipText.setter
    def ControlTipText(self, value):
        self.listbox.ControlTipText = value

    @property
    def ControlType(self):
        return self.listbox.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.listbox.ControlType = value

    @property
    def DefaultValue(self):
        return self.listbox.DefaultValue

    @DefaultValue.setter
    def DefaultValue(self, value):
        self.listbox.DefaultValue = value

    @property
    def DisplayWhen(self):
        return self.listbox.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.listbox.DisplayWhen = value

    @property
    def Enabled(self):
        return self.listbox.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.listbox.Enabled = value

    @property
    def EventProcPrefix(self):
        return self.listbox.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.listbox.EventProcPrefix = value

    @property
    def FontBold(self):
        return self.listbox.FontBold

    @property
    def FontItalic(self):
        return self.listbox.FontItalic

    @property
    def FontName(self):
        return self.listbox.FontName

    @property
    def FontSize(self):
        return self.listbox.FontSize

    @property
    def FontUnderline(self):
        return self.listbox.FontUnderline

    @property
    def FontWeight(self):
        return self.listbox.FontWeight

    @FontWeight.setter
    def FontWeight(self, value):
        self.listbox.FontWeight = value

    @property
    def ForeColor(self):
        return self.listbox.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.listbox.ForeColor = value

    @property
    def ForeShade(self):
        return self.listbox.ForeShade

    @ForeShade.setter
    def ForeShade(self, value):
        self.listbox.ForeShade = value

    @property
    def ForeThemeColorIndex(self):
        return self.listbox.ForeThemeColorIndex

    @ForeThemeColorIndex.setter
    def ForeThemeColorIndex(self, value):
        self.listbox.ForeThemeColorIndex = value

    @property
    def ForeTint(self):
        return self.listbox.ForeTint

    @ForeTint.setter
    def ForeTint(self, value):
        self.listbox.ForeTint = value

    @property
    def GridlineColor(self):
        return self.listbox.GridlineColor

    @GridlineColor.setter
    def GridlineColor(self, value):
        self.listbox.GridlineColor = value

    @property
    def GridlineShade(self):
        return self.listbox.GridlineShade

    @GridlineShade.setter
    def GridlineShade(self, value):
        self.listbox.GridlineShade = value

    @property
    def GridlineStyleBottom(self):
        return self.listbox.GridlineStyleBottom

    @GridlineStyleBottom.setter
    def GridlineStyleBottom(self, value):
        self.listbox.GridlineStyleBottom = value

    @property
    def GridlineStyleLeft(self):
        return self.listbox.GridlineStyleLeft

    @GridlineStyleLeft.setter
    def GridlineStyleLeft(self, value):
        self.listbox.GridlineStyleLeft = value

    @property
    def GridlineStyleRight(self):
        return self.listbox.GridlineStyleRight

    @GridlineStyleRight.setter
    def GridlineStyleRight(self, value):
        self.listbox.GridlineStyleRight = value

    @property
    def GridlineStyleTop(self):
        return self.listbox.GridlineStyleTop

    @GridlineStyleTop.setter
    def GridlineStyleTop(self, value):
        self.listbox.GridlineStyleTop = value

    @property
    def GridlineThemeColorIndex(self):
        return self.listbox.GridlineThemeColorIndex

    @GridlineThemeColorIndex.setter
    def GridlineThemeColorIndex(self, value):
        self.listbox.GridlineThemeColorIndex = value

    @property
    def GridlineTint(self):
        return self.listbox.GridlineTint

    @GridlineTint.setter
    def GridlineTint(self, value):
        self.listbox.GridlineTint = value

    @property
    def GridlineWidthBottom(self):
        return self.listbox.GridlineWidthBottom

    @GridlineWidthBottom.setter
    def GridlineWidthBottom(self, value):
        self.listbox.GridlineWidthBottom = value

    @property
    def GridlineWidthLeft(self):
        return self.listbox.GridlineWidthLeft

    @GridlineWidthLeft.setter
    def GridlineWidthLeft(self, value):
        self.listbox.GridlineWidthLeft = value

    @property
    def GridlineWidthRight(self):
        return self.listbox.GridlineWidthRight

    @GridlineWidthRight.setter
    def GridlineWidthRight(self, value):
        self.listbox.GridlineWidthRight = value

    @property
    def GridlineWidthTop(self):
        return self.listbox.GridlineWidthTop

    @GridlineWidthTop.setter
    def GridlineWidthTop(self, value):
        self.listbox.GridlineWidthTop = value

    @property
    def Height(self):
        return self.listbox.Height

    @Height.setter
    def Height(self, value):
        self.listbox.Height = value

    @property
    def HelpContextId(self):
        return self.listbox.HelpContextId

    @HelpContextId.setter
    def HelpContextId(self, value):
        self.listbox.HelpContextId = value

    @property
    def HideDuplicates(self):
        return self.listbox.HideDuplicates

    @HideDuplicates.setter
    def HideDuplicates(self, value):
        self.listbox.HideDuplicates = value

    @property
    def HorizontalAnchor(self):
        return self.listbox.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.listbox.HorizontalAnchor = value

    @property
    def Hyperlink(self):
        return self.listbox.Hyperlink

    @property
    def IMEHold(self):
        return self.listbox.IMEHold

    @IMEHold.setter
    def IMEHold(self, value):
        self.listbox.IMEHold = value

    @property
    def InheritValueList(self):
        return self.listbox.InheritValueList

    @InheritValueList.setter
    def InheritValueList(self, value):
        self.listbox.InheritValueList = value

    @property
    def InSelection(self):
        return self.listbox.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.listbox.InSelection = value

    @property
    def IsVisible(self):
        return self.listbox.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.listbox.IsVisible = value

    def ItemData(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.listbox.ItemData):
            return self.listbox.ItemData(*args, **arguments)
        else:
            return self.listbox.GetItemData(*args, **arguments)

    @property
    def ItemsSelected(self):
        return self.listbox.ItemsSelected

    @property
    def LabelAlign(self):
        return self.listbox.LabelAlign

    @LabelAlign.setter
    def LabelAlign(self, value):
        self.listbox.LabelAlign = value

    @property
    def LabelX(self):
        return self.listbox.LabelX

    @LabelX.setter
    def LabelX(self, value):
        self.listbox.LabelX = value

    @property
    def LabelY(self):
        return self.listbox.LabelY

    @LabelY.setter
    def LabelY(self, value):
        self.listbox.LabelY = value

    @property
    def Layout(self):
        return AcLayoutType(self.listbox.Layout)

    @property
    def LayoutID(self):
        return self.listbox.LayoutID

    @property
    def Left(self):
        return self.listbox.Left

    @Left.setter
    def Left(self, value):
        self.listbox.Left = value

    @property
    def LeftPadding(self):
        return self.listbox.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.listbox.LeftPadding = value

    @property
    def ListCount(self):
        return self.listbox.ListCount

    @ListCount.setter
    def ListCount(self, value):
        self.listbox.ListCount = value

    @property
    def ListIndex(self):
        return self.listbox.ListIndex

    @ListIndex.setter
    def ListIndex(self, value):
        self.listbox.ListIndex = value

    @property
    def ListItemsEditForm(self):
        return self.listbox.ListItemsEditForm

    @ListItemsEditForm.setter
    def ListItemsEditForm(self, value):
        self.listbox.ListItemsEditForm = value

    @property
    def Locked(self):
        return self.listbox.Locked

    @Locked.setter
    def Locked(self, value):
        self.listbox.Locked = value

    @property
    def MultiSelect(self):
        return self.listbox.MultiSelect

    @MultiSelect.setter
    def MultiSelect(self, value):
        self.listbox.MultiSelect = value

    @property
    def Name(self):
        return self.listbox.Name

    @Name.setter
    def Name(self, value):
        self.listbox.Name = value

    @property
    def OldBorderStyle(self):
        return self.listbox.OldBorderStyle

    @OldBorderStyle.setter
    def OldBorderStyle(self, value):
        self.listbox.OldBorderStyle = value

    @property
    def OldValue(self):
        return self.listbox.OldValue

    @property
    def OnClick(self):
        return self.listbox.OnClick

    @OnClick.setter
    def OnClick(self, value):
        self.listbox.OnClick = value

    @property
    def OnDblClick(self):
        return self.listbox.OnDblClick

    @OnDblClick.setter
    def OnDblClick(self, value):
        self.listbox.OnDblClick = value

    @property
    def OnEnter(self):
        return self.listbox.OnEnter

    @OnEnter.setter
    def OnEnter(self, value):
        self.listbox.OnEnter = value

    @property
    def OnExit(self):
        return self.listbox.OnExit

    @OnExit.setter
    def OnExit(self, value):
        self.listbox.OnExit = value

    @property
    def OnGotFocus(self):
        return self.listbox.OnGotFocus

    @OnGotFocus.setter
    def OnGotFocus(self, value):
        self.listbox.OnGotFocus = value

    @property
    def OnKeyDown(self):
        return self.listbox.OnKeyDown

    @OnKeyDown.setter
    def OnKeyDown(self, value):
        self.listbox.OnKeyDown = value

    @property
    def OnKeyPress(self):
        return self.listbox.OnKeyPress

    @OnKeyPress.setter
    def OnKeyPress(self, value):
        self.listbox.OnKeyPress = value

    @property
    def OnKeyUp(self):
        return self.listbox.OnKeyUp

    @OnKeyUp.setter
    def OnKeyUp(self, value):
        self.listbox.OnKeyUp = value

    @property
    def OnLostFocus(self):
        return self.listbox.OnLostFocus

    @OnLostFocus.setter
    def OnLostFocus(self, value):
        self.listbox.OnLostFocus = value

    @property
    def OnMouseDown(self):
        return self.listbox.OnMouseDown

    @OnMouseDown.setter
    def OnMouseDown(self, value):
        self.listbox.OnMouseDown = value

    @property
    def OnMouseMove(self):
        return self.listbox.OnMouseMove

    @OnMouseMove.setter
    def OnMouseMove(self, value):
        self.listbox.OnMouseMove = value

    @property
    def OnMouseUp(self):
        return self.listbox.OnMouseUp

    @OnMouseUp.setter
    def OnMouseUp(self, value):
        self.listbox.OnMouseUp = value

    @property
    def Parent(self):
        return self.listbox.Parent

    @property
    def Properties(self):
        return Properties(self.listbox.Properties)

    @property
    def ReadingOrder(self):
        return self.listbox.ReadingOrder

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.listbox.ReadingOrder = value

    @property
    def Recordset(self):
        return self.listbox.Recordset

    @Recordset.setter
    def Recordset(self, value):
        self.listbox.Recordset = value

    @property
    def RightPadding(self):
        return self.listbox.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.listbox.RightPadding = value

    @property
    def RowSource(self):
        return self.listbox.RowSource

    @RowSource.setter
    def RowSource(self, value):
        self.listbox.RowSource = value

    @property
    def RowSourceType(self):
        return self.listbox.RowSourceType

    @RowSourceType.setter
    def RowSourceType(self, value):
        self.listbox.RowSourceType = value

    @property
    def ScrollBarAlign(self):
        return self.listbox.ScrollBarAlign

    @ScrollBarAlign.setter
    def ScrollBarAlign(self, value):
        self.listbox.ScrollBarAlign = value

    @property
    def Section(self):
        return self.listbox.Section

    @Section.setter
    def Section(self, value):
        self.listbox.Section = value

    @property
    def Selected(self):
        return self.listbox.Selected

    @Selected.setter
    def Selected(self, value):
        self.listbox.Selected = value

    @property
    def ShortcutMenuBar(self):
        return self.listbox.ShortcutMenuBar

    @ShortcutMenuBar.setter
    def ShortcutMenuBar(self, value):
        self.listbox.ShortcutMenuBar = value

    @property
    def ShowOnlyRowSourceValues(self):
        return self.listbox.ShowOnlyRowSourceValues

    @ShowOnlyRowSourceValues.setter
    def ShowOnlyRowSourceValues(self, value):
        self.listbox.ShowOnlyRowSourceValues = value

    @property
    def SmartTags(self):
        return SmartTags(self.listbox.SmartTags)

    @property
    def SpecialEffect(self):
        return self.listbox.SpecialEffect

    @SpecialEffect.setter
    def SpecialEffect(self, value):
        self.listbox.SpecialEffect = value

    @property
    def StatusBarText(self):
        return self.listbox.StatusBarText

    @StatusBarText.setter
    def StatusBarText(self, value):
        self.listbox.StatusBarText = value

    @property
    def TabIndex(self):
        return self.listbox.TabIndex

    @TabIndex.setter
    def TabIndex(self, value):
        self.listbox.TabIndex = value

    @property
    def TabStop(self):
        return self.listbox.TabStop

    @TabStop.setter
    def TabStop(self, value):
        self.listbox.TabStop = value

    @property
    def Tag(self):
        return self.listbox.Tag

    @Tag.setter
    def Tag(self, value):
        self.listbox.Tag = value

    @property
    def ThemeFontIndex(self):
        return self.listbox.ThemeFontIndex

    @ThemeFontIndex.setter
    def ThemeFontIndex(self, value):
        self.listbox.ThemeFontIndex = value

    @property
    def Top(self):
        return self.listbox.Top

    @Top.setter
    def Top(self, value):
        self.listbox.Top = value

    @property
    def TopPadding(self):
        return self.listbox.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.listbox.TopPadding = value

    @property
    def ValidationRule(self):
        return self.listbox.ValidationRule

    @ValidationRule.setter
    def ValidationRule(self, value):
        self.listbox.ValidationRule = value

    @property
    def ValidationText(self):
        return self.listbox.ValidationText

    @ValidationText.setter
    def ValidationText(self, value):
        self.listbox.ValidationText = value

    @property
    def Value(self):
        return self.listbox.Value

    @Value.setter
    def Value(self, value):
        self.listbox.Value = value

    @property
    def VerticalAnchor(self):
        return self.listbox.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.listbox.VerticalAnchor = value

    @property
    def Visible(self):
        return self.listbox.Visible

    @Visible.setter
    def Visible(self, value):
        self.listbox.Visible = value

    @property
    def Width(self):
        return self.listbox.Width

    @Width.setter
    def Width(self, value):
        self.listbox.Width = value

    def AddItem(self, *args, Item=None, Index=None):
        arguments = {"Item": Item, "Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.listbox.AddItem(*args, **arguments)

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.listbox.Move(*args, **arguments)

    def RemoveItem(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.listbox.RemoveItem(*args, **arguments)

    def Requery(self):
        self.listbox.Requery()

    def SetFocus(self):
        return self.listbox.SetFocus()

    def SizeToFit(self):
        self.listbox.SizeToFit()

    def Undo(self):
        self.listbox.Undo()

class MacroError:

    def __init__(self, macroerror=None):
        self.macroerror = macroerror

    @property
    def Arguments(self):
        return self.macroerror.Arguments

    @property
    def Condition(self):
        return self.macroerror.Condition

    @property
    def Description(self):
        return self.macroerror.Description

    @property
    def MacroName(self):
        return self.macroerror.MacroName

    @property
    def Number(self):
        return self.macroerror.Number

class Module:

    def __init__(self, module=None):
        self.module = module

    @property
    def Application(self):
        return self.module.Application

    @property
    def CountOfDeclarationLines(self):
        return self.module.CountOfDeclarationLines

    @property
    def CountOfLines(self):
        return self.module.CountOfLines

    def Lines(self, *args, Line=None, NumLines=None):
        arguments = {"Line": Line, "NumLines": NumLines}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.module.Lines):
            return self.module.Lines(*args, **arguments)
        else:
            return self.module.GetLines(*args, **arguments)

    @property
    def Name(self):
        return self.module.Name

    @Name.setter
    def Name(self, value):
        self.module.Name = value

    @property
    def Parent(self):
        return self.module.Parent

    def ProcBodyLine(self, *args, ProcName=None, ProcKind=None):
        arguments = {"ProcName": ProcName, "ProcKind": ProcKind}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.module.ProcBodyLine):
            return self.module.ProcBodyLine(*args, **arguments)
        else:
            return self.module.GetProcBodyLine(*args, **arguments)

    def ProcCountLines(self, *args, ProcName=None, ProcKind=None):
        arguments = {"ProcName": ProcName, "ProcKind": ProcKind}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.module.ProcCountLines):
            return self.module.ProcCountLines(*args, **arguments)
        else:
            return self.module.GetProcCountLines(*args, **arguments)

    def ProcOfLine(self, *args, Line=None, ProcKind=None):
        arguments = {"Line": Line, "ProcKind": ProcKind}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.module.ProcOfLine):
            return self.module.ProcOfLine(*args, **arguments)
        else:
            return self.module.GetProcOfLine(*args, **arguments)

    def ProcStartLine(self, *args, ProcName=None, ProcKind=None):
        arguments = {"ProcName": ProcName, "ProcKind": ProcKind}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.module.ProcStartLine):
            return self.module.ProcStartLine(*args, **arguments)
        else:
            return self.module.GetProcStartLine(*args, **arguments)

    @property
    def Type(self):
        return self.module.Type

    def AddFromFile(self, *args, FileName=None):
        arguments = {"FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.module.AddFromFile(*args, **arguments)

    def AddFromString(self, *args, String=None):
        arguments = {"String": String}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.module.AddFromString(*args, **arguments)

    def CreateEventProc(self, *args, EventName=None, ObjectName=None):
        arguments = {"EventName": EventName, "ObjectName": ObjectName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.module.CreateEventProc(*args, **arguments)

    def DeleteLines(self, *args, StartLine=None, Count=None):
        arguments = {"StartLine": StartLine, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.module.DeleteLines(*args, **arguments)

    def Find(self, *args, Target=None, StartLine=None, StartColumn=None, EndLine=None, EndColumn=None, WholeWord=None, MatchCase=None, PatternSearch=None):
        arguments = {"Target": Target, "StartLine": StartLine, "StartColumn": StartColumn, "EndLine": EndLine, "EndColumn": EndColumn, "WholeWord": WholeWord, "MatchCase": MatchCase, "PatternSearch": PatternSearch}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.module.Find(*args, **arguments)

    def InsertLines(self, *args, Line=None, String=None):
        arguments = {"Line": Line, "String": String}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.module.InsertLines(*args, **arguments)

    def InsertText(self, *args, Text=None):
        arguments = {"Text": Text}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.module.InsertText(*args, **arguments)

    def ReplaceLine(self, *args, Line=None, String=None):
        arguments = {"Line": Line, "String": String}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.module.ReplaceLine(*args, **arguments)

class Modules:

    def __init__(self, modules=None):
        self.modules = modules

    @property
    def Application(self):
        return self.modules.Application

    @property
    def Count(self):
        return self.modules.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.modules.Item):
            return self.modules.Item(*args, **arguments)
        else:
            return self.modules.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.modules.Parent

class NavigationButton:

    def __init__(self, navigationbutton=None):
        self.navigationbutton = navigationbutton

    @property
    def AddColon(self):
        return self.navigationbutton.AddColon

    @AddColon.setter
    def AddColon(self, value):
        self.navigationbutton.AddColon = value

    @property
    def Alignment(self):
        return self.navigationbutton.Alignment

    @Alignment.setter
    def Alignment(self, value):
        self.navigationbutton.Alignment = value

    @property
    def Application(self):
        return self.navigationbutton.Application

    @property
    def AutoLabel(self):
        return self.navigationbutton.AutoLabel

    @AutoLabel.setter
    def AutoLabel(self, value):
        self.navigationbutton.AutoLabel = value

    @property
    def AutoRepeat(self):
        return self.navigationbutton.AutoRepeat

    @AutoRepeat.setter
    def AutoRepeat(self, value):
        self.navigationbutton.AutoRepeat = value

    @property
    def BackColor(self):
        return self.navigationbutton.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.navigationbutton.BackColor = value

    @property
    def BackShade(self):
        return self.navigationbutton.BackShade

    @BackShade.setter
    def BackShade(self, value):
        self.navigationbutton.BackShade = value

    @property
    def BackStyle(self):
        return self.navigationbutton.BackStyle

    @BackStyle.setter
    def BackStyle(self, value):
        self.navigationbutton.BackStyle = value

    @property
    def BackThemeColorIndex(self):
        return self.navigationbutton.BackThemeColorIndex

    @BackThemeColorIndex.setter
    def BackThemeColorIndex(self, value):
        self.navigationbutton.BackThemeColorIndex = value

    @property
    def BackTint(self):
        return self.navigationbutton.BackTint

    @BackTint.setter
    def BackTint(self, value):
        self.navigationbutton.BackTint = value

    @property
    def Bevel(self):
        return self.navigationbutton.Bevel

    @Bevel.setter
    def Bevel(self, value):
        self.navigationbutton.Bevel = value

    @property
    def BorderColor(self):
        return self.navigationbutton.BorderColor

    @BorderColor.setter
    def BorderColor(self, value):
        self.navigationbutton.BorderColor = value

    @property
    def BorderShade(self):
        return self.navigationbutton.BorderShade

    @BorderShade.setter
    def BorderShade(self, value):
        self.navigationbutton.BorderShade = value

    @property
    def BorderStyle(self):
        return self.navigationbutton.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.navigationbutton.BorderStyle = value

    @property
    def BorderThemeColorIndex(self):
        return self.navigationbutton.BorderThemeColorIndex

    @BorderThemeColorIndex.setter
    def BorderThemeColorIndex(self, value):
        self.navigationbutton.BorderThemeColorIndex = value

    @property
    def BorderTint(self):
        return self.navigationbutton.BorderTint

    @BorderTint.setter
    def BorderTint(self, value):
        self.navigationbutton.BorderTint = value

    @property
    def BorderWidth(self):
        return self.navigationbutton.BorderWidth

    @BorderWidth.setter
    def BorderWidth(self, value):
        self.navigationbutton.BorderWidth = value

    @property
    def BottomPadding(self):
        return self.navigationbutton.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.navigationbutton.BottomPadding = value

    @property
    def Caption(self):
        return self.navigationbutton.Caption

    @Caption.setter
    def Caption(self, value):
        self.navigationbutton.Caption = value

    @property
    def Controls(self):
        return Controls(self.navigationbutton.Controls)

    @property
    def ControlTipText(self):
        return self.navigationbutton.ControlTipText

    @ControlTipText.setter
    def ControlTipText(self, value):
        self.navigationbutton.ControlTipText = value

    @property
    def ControlType(self):
        return self.navigationbutton.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.navigationbutton.ControlType = value

    @property
    def CursorOnHover(self):
        return self.navigationbutton.CursorOnHover

    @CursorOnHover.setter
    def CursorOnHover(self, value):
        self.navigationbutton.CursorOnHover = value

    @property
    def DisplayWhen(self):
        return self.navigationbutton.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.navigationbutton.DisplayWhen = value

    @property
    def Enabled(self):
        return self.navigationbutton.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.navigationbutton.Enabled = value

    @property
    def EventProcPrefix(self):
        return self.navigationbutton.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.navigationbutton.EventProcPrefix = value

    @property
    def FontBold(self):
        return self.navigationbutton.FontBold

    @property
    def FontItalic(self):
        return self.navigationbutton.FontItalic

    @property
    def FontName(self):
        return self.navigationbutton.FontName

    @property
    def FontSize(self):
        return self.navigationbutton.FontSize

    @property
    def FontUnderline(self):
        return self.navigationbutton.FontUnderline

    @property
    def FontWeight(self):
        return self.navigationbutton.FontWeight

    @FontWeight.setter
    def FontWeight(self, value):
        self.navigationbutton.FontWeight = value

    @property
    def ForeColor(self):
        return self.navigationbutton.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.navigationbutton.ForeColor = value

    @property
    def ForeShade(self):
        return self.navigationbutton.ForeShade

    @ForeShade.setter
    def ForeShade(self, value):
        self.navigationbutton.ForeShade = value

    @property
    def ForeThemeColorIndex(self):
        return self.navigationbutton.ForeThemeColorIndex

    @ForeThemeColorIndex.setter
    def ForeThemeColorIndex(self, value):
        self.navigationbutton.ForeThemeColorIndex = value

    @property
    def ForeTint(self):
        return self.navigationbutton.ForeTint

    @ForeTint.setter
    def ForeTint(self, value):
        self.navigationbutton.ForeTint = value

    @property
    def Glow(self):
        return self.navigationbutton.Glow

    @Glow.setter
    def Glow(self, value):
        self.navigationbutton.Glow = value

    @property
    def Gradient(self):
        return self.navigationbutton.Gradient

    @Gradient.setter
    def Gradient(self, value):
        self.navigationbutton.Gradient = value

    @property
    def GridlineColor(self):
        return self.navigationbutton.GridlineColor

    @GridlineColor.setter
    def GridlineColor(self, value):
        self.navigationbutton.GridlineColor = value

    @property
    def GridlineShade(self):
        return self.navigationbutton.GridlineShade

    @GridlineShade.setter
    def GridlineShade(self, value):
        self.navigationbutton.GridlineShade = value

    @property
    def GridlineStyleBottom(self):
        return self.navigationbutton.GridlineStyleBottom

    @GridlineStyleBottom.setter
    def GridlineStyleBottom(self, value):
        self.navigationbutton.GridlineStyleBottom = value

    @property
    def GridlineStyleLeft(self):
        return self.navigationbutton.GridlineStyleLeft

    @GridlineStyleLeft.setter
    def GridlineStyleLeft(self, value):
        self.navigationbutton.GridlineStyleLeft = value

    @property
    def GridlineStyleRight(self):
        return self.navigationbutton.GridlineStyleRight

    @GridlineStyleRight.setter
    def GridlineStyleRight(self, value):
        self.navigationbutton.GridlineStyleRight = value

    @property
    def GridlineStyleTop(self):
        return self.navigationbutton.GridlineStyleTop

    @GridlineStyleTop.setter
    def GridlineStyleTop(self, value):
        self.navigationbutton.GridlineStyleTop = value

    @property
    def GridlineThemeColorIndex(self):
        return self.navigationbutton.GridlineThemeColorIndex

    @GridlineThemeColorIndex.setter
    def GridlineThemeColorIndex(self, value):
        self.navigationbutton.GridlineThemeColorIndex = value

    @property
    def GridlineTint(self):
        return self.navigationbutton.GridlineTint

    @GridlineTint.setter
    def GridlineTint(self, value):
        self.navigationbutton.GridlineTint = value

    @property
    def GridlineWidthBottom(self):
        return self.navigationbutton.GridlineWidthBottom

    @GridlineWidthBottom.setter
    def GridlineWidthBottom(self, value):
        self.navigationbutton.GridlineWidthBottom = value

    @property
    def GridlineWidthLeft(self):
        return self.navigationbutton.GridlineWidthLeft

    @GridlineWidthLeft.setter
    def GridlineWidthLeft(self, value):
        self.navigationbutton.GridlineWidthLeft = value

    @property
    def GridlineWidthRight(self):
        return self.navigationbutton.GridlineWidthRight

    @GridlineWidthRight.setter
    def GridlineWidthRight(self, value):
        self.navigationbutton.GridlineWidthRight = value

    @property
    def GridlineWidthTop(self):
        return self.navigationbutton.GridlineWidthTop

    @GridlineWidthTop.setter
    def GridlineWidthTop(self, value):
        self.navigationbutton.GridlineWidthTop = value

    @property
    def Height(self):
        return self.navigationbutton.Height

    @Height.setter
    def Height(self, value):
        self.navigationbutton.Height = value

    @property
    def HelpContextId(self):
        return self.navigationbutton.HelpContextId

    @HelpContextId.setter
    def HelpContextId(self, value):
        self.navigationbutton.HelpContextId = value

    @property
    def HorizontalAnchor(self):
        return self.navigationbutton.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.navigationbutton.HorizontalAnchor = value

    @property
    def HoverColor(self):
        return self.navigationbutton.HoverColor

    @HoverColor.setter
    def HoverColor(self, value):
        self.navigationbutton.HoverColor = value

    @property
    def HoverForeColor(self):
        return self.navigationbutton.HoverForeColor

    @HoverForeColor.setter
    def HoverForeColor(self, value):
        self.navigationbutton.HoverForeColor = value

    @property
    def HoverForeShade(self):
        return self.navigationbutton.HoverForeShade

    @HoverForeShade.setter
    def HoverForeShade(self, value):
        self.navigationbutton.HoverForeShade = value

    @property
    def HoverForeThemeColorIndex(self):
        return self.navigationbutton.HoverForeThemeColorIndex

    @HoverForeThemeColorIndex.setter
    def HoverForeThemeColorIndex(self, value):
        self.navigationbutton.HoverForeThemeColorIndex = value

    @property
    def HoverForeTint(self):
        return self.navigationbutton.HoverForeTint

    @HoverForeTint.setter
    def HoverForeTint(self, value):
        self.navigationbutton.HoverForeTint = value

    @property
    def HoverShade(self):
        return self.navigationbutton.HoverShade

    @HoverShade.setter
    def HoverShade(self, value):
        self.navigationbutton.HoverShade = value

    @property
    def HoverThemeColorIndex(self):
        return self.navigationbutton.HoverThemeColorIndex

    @HoverThemeColorIndex.setter
    def HoverThemeColorIndex(self, value):
        self.navigationbutton.HoverThemeColorIndex = value

    @property
    def HoverTint(self):
        return self.navigationbutton.HoverTint

    @HoverTint.setter
    def HoverTint(self, value):
        self.navigationbutton.HoverTint = value

    @property
    def Hyperlink(self):
        return self.navigationbutton.Hyperlink

    @property
    def HyperlinkAddress(self):
        return self.navigationbutton.HyperlinkAddress

    @HyperlinkAddress.setter
    def HyperlinkAddress(self, value):
        self.navigationbutton.HyperlinkAddress = value

    @property
    def HyperlinkSubAddress(self):
        return self.navigationbutton.HyperlinkSubAddress

    @HyperlinkSubAddress.setter
    def HyperlinkSubAddress(self, value):
        self.navigationbutton.HyperlinkSubAddress = value

    @property
    def InSelection(self):
        return self.navigationbutton.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.navigationbutton.InSelection = value

    @property
    def IsVisible(self):
        return self.navigationbutton.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.navigationbutton.IsVisible = value

    @property
    def LabelAlign(self):
        return self.navigationbutton.LabelAlign

    @LabelAlign.setter
    def LabelAlign(self, value):
        self.navigationbutton.LabelAlign = value

    @property
    def LabelX(self):
        return self.navigationbutton.LabelX

    @LabelX.setter
    def LabelX(self, value):
        self.navigationbutton.LabelX = value

    @property
    def LabelY(self):
        return self.navigationbutton.LabelY

    @LabelY.setter
    def LabelY(self, value):
        self.navigationbutton.LabelY = value

    @property
    def Layout(self):
        return AcLayoutType(self.navigationbutton.Layout)

    @property
    def LayoutID(self):
        return self.navigationbutton.LayoutID

    @property
    def Left(self):
        return self.navigationbutton.Left

    @Left.setter
    def Left(self, value):
        self.navigationbutton.Left = value

    @property
    def LeftPadding(self):
        return self.navigationbutton.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.navigationbutton.LeftPadding = value

    @property
    def Name(self):
        return self.navigationbutton.Name

    @Name.setter
    def Name(self, value):
        self.navigationbutton.Name = value

    @property
    def NavigationTargetName(self):
        return self.navigationbutton.NavigationTargetName

    @NavigationTargetName.setter
    def NavigationTargetName(self, value):
        self.navigationbutton.NavigationTargetName = value

    @property
    def NavigationWhereClause(self):
        return self.navigationbutton.NavigationWhereClause

    @NavigationWhereClause.setter
    def NavigationWhereClause(self, value):
        self.navigationbutton.NavigationWhereClause = value

    @property
    def ObjectPalette(self):
        return self.navigationbutton.ObjectPalette

    @ObjectPalette.setter
    def ObjectPalette(self, value):
        self.navigationbutton.ObjectPalette = value

    @property
    def OldValue(self):
        return self.navigationbutton.OldValue

    @property
    def OnClick(self):
        return self.navigationbutton.OnClick

    @OnClick.setter
    def OnClick(self, value):
        self.navigationbutton.OnClick = value

    @property
    def OnDblClick(self):
        return self.navigationbutton.OnDblClick

    @OnDblClick.setter
    def OnDblClick(self, value):
        self.navigationbutton.OnDblClick = value

    @property
    def OnEnter(self):
        return self.navigationbutton.OnEnter

    @OnEnter.setter
    def OnEnter(self, value):
        self.navigationbutton.OnEnter = value

    @property
    def OnExit(self):
        return self.navigationbutton.OnExit

    @OnExit.setter
    def OnExit(self, value):
        self.navigationbutton.OnExit = value

    @property
    def OnGotFocus(self):
        return self.navigationbutton.OnGotFocus

    @OnGotFocus.setter
    def OnGotFocus(self, value):
        self.navigationbutton.OnGotFocus = value

    @property
    def OnKeyDown(self):
        return self.navigationbutton.OnKeyDown

    @OnKeyDown.setter
    def OnKeyDown(self, value):
        self.navigationbutton.OnKeyDown = value

    @property
    def OnKeyPress(self):
        return self.navigationbutton.OnKeyPress

    @OnKeyPress.setter
    def OnKeyPress(self, value):
        self.navigationbutton.OnKeyPress = value

    @property
    def OnKeyUp(self):
        return self.navigationbutton.OnKeyUp

    @OnKeyUp.setter
    def OnKeyUp(self, value):
        self.navigationbutton.OnKeyUp = value

    @property
    def OnLostFocus(self):
        return self.navigationbutton.OnLostFocus

    @OnLostFocus.setter
    def OnLostFocus(self, value):
        self.navigationbutton.OnLostFocus = value

    @property
    def OnMouseDown(self):
        return self.navigationbutton.OnMouseDown

    @OnMouseDown.setter
    def OnMouseDown(self, value):
        self.navigationbutton.OnMouseDown = value

    @property
    def OnMouseMove(self):
        return self.navigationbutton.OnMouseMove

    @OnMouseMove.setter
    def OnMouseMove(self, value):
        self.navigationbutton.OnMouseMove = value

    @property
    def OnMouseUp(self):
        return self.navigationbutton.OnMouseUp

    @OnMouseUp.setter
    def OnMouseUp(self, value):
        self.navigationbutton.OnMouseUp = value

    @property
    def OnPush(self):
        return self.navigationbutton.OnPush

    @OnPush.setter
    def OnPush(self, value):
        self.navigationbutton.OnPush = value

    @property
    def Parent(self):
        return self.navigationbutton.Parent

    @property
    def ParentTab(self):
        return self.navigationbutton.ParentTab

    @property
    def Picture(self):
        return self.navigationbutton.Picture

    @Picture.setter
    def Picture(self, value):
        self.navigationbutton.Picture = value

    @property
    def PictureCaptionArrangement(self):
        return self.navigationbutton.PictureCaptionArrangement

    @PictureCaptionArrangement.setter
    def PictureCaptionArrangement(self, value):
        self.navigationbutton.PictureCaptionArrangement = value

    @property
    def PictureData(self):
        return self.navigationbutton.PictureData

    @PictureData.setter
    def PictureData(self, value):
        self.navigationbutton.PictureData = value

    @property
    def PictureType(self):
        return self.navigationbutton.PictureType

    @PictureType.setter
    def PictureType(self, value):
        self.navigationbutton.PictureType = value

    @property
    def PressedColor(self):
        return self.navigationbutton.PressedColor

    @PressedColor.setter
    def PressedColor(self, value):
        self.navigationbutton.PressedColor = value

    @property
    def PressedForeColor(self):
        return self.navigationbutton.PressedForeColor

    @PressedForeColor.setter
    def PressedForeColor(self, value):
        self.navigationbutton.PressedForeColor = value

    @property
    def PressedForeShade(self):
        return self.navigationbutton.PressedForeShade

    @PressedForeShade.setter
    def PressedForeShade(self, value):
        self.navigationbutton.PressedForeShade = value

    @property
    def PressedForeThemeColorIndex(self):
        return self.navigationbutton.PressedForeThemeColorIndex

    @PressedForeThemeColorIndex.setter
    def PressedForeThemeColorIndex(self, value):
        self.navigationbutton.PressedForeThemeColorIndex = value

    @property
    def PressedForeTint(self):
        return self.navigationbutton.PressedForeTint

    @PressedForeTint.setter
    def PressedForeTint(self, value):
        self.navigationbutton.PressedForeTint = value

    @property
    def PressedShade(self):
        return self.navigationbutton.PressedShade

    @PressedShade.setter
    def PressedShade(self, value):
        self.navigationbutton.PressedShade = value

    @property
    def PressedThemeColorIndex(self):
        return self.navigationbutton.PressedThemeColorIndex

    @PressedThemeColorIndex.setter
    def PressedThemeColorIndex(self, value):
        self.navigationbutton.PressedThemeColorIndex = value

    @property
    def PressedTint(self):
        return self.navigationbutton.PressedTint

    @PressedTint.setter
    def PressedTint(self, value):
        self.navigationbutton.PressedTint = value

    @property
    def Properties(self):
        return Properties(self.navigationbutton.Properties)

    @property
    def QuickStyle(self):
        return self.navigationbutton.QuickStyle

    @QuickStyle.setter
    def QuickStyle(self, value):
        self.navigationbutton.QuickStyle = value

    @property
    def ReadingOrder(self):
        return self.navigationbutton.ReadingOrder

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.navigationbutton.ReadingOrder = value

    @property
    def RightPadding(self):
        return self.navigationbutton.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.navigationbutton.RightPadding = value

    @property
    def Section(self):
        return self.navigationbutton.Section

    @Section.setter
    def Section(self, value):
        self.navigationbutton.Section = value

    @property
    def Shadow(self):
        return self.navigationbutton.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.navigationbutton.Shadow = value

    @property
    def Shape(self):
        return self.navigationbutton.Shape

    @Shape.setter
    def Shape(self, value):
        self.navigationbutton.Shape = value

    @property
    def ShortcutMenuBar(self):
        return self.navigationbutton.ShortcutMenuBar

    @ShortcutMenuBar.setter
    def ShortcutMenuBar(self, value):
        self.navigationbutton.ShortcutMenuBar = value

    @property
    def SoftEdges(self):
        return self.navigationbutton.SoftEdges

    @SoftEdges.setter
    def SoftEdges(self, value):
        self.navigationbutton.SoftEdges = value

    @property
    def StatusBarText(self):
        return self.navigationbutton.StatusBarText

    @StatusBarText.setter
    def StatusBarText(self, value):
        self.navigationbutton.StatusBarText = value

    @property
    def TabIndex(self):
        return self.navigationbutton.TabIndex

    @TabIndex.setter
    def TabIndex(self, value):
        self.navigationbutton.TabIndex = value

    @property
    def TabStop(self):
        return self.navigationbutton.TabStop

    @TabStop.setter
    def TabStop(self, value):
        self.navigationbutton.TabStop = value

    @property
    def Tag(self):
        return self.navigationbutton.Tag

    @Tag.setter
    def Tag(self, value):
        self.navigationbutton.Tag = value

    @property
    def ThemeFontIndex(self):
        return self.navigationbutton.ThemeFontIndex

    @ThemeFontIndex.setter
    def ThemeFontIndex(self, value):
        self.navigationbutton.ThemeFontIndex = value

    @property
    def Top(self):
        return self.navigationbutton.Top

    @Top.setter
    def Top(self, value):
        self.navigationbutton.Top = value

    @property
    def TopPadding(self):
        return self.navigationbutton.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.navigationbutton.TopPadding = value

    @property
    def Transparent(self):
        return self.navigationbutton.Transparent

    @Transparent.setter
    def Transparent(self, value):
        self.navigationbutton.Transparent = value

    @property
    def VerticalAnchor(self):
        return self.navigationbutton.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.navigationbutton.VerticalAnchor = value

    @property
    def Visible(self):
        return self.navigationbutton.Visible

    @Visible.setter
    def Visible(self, value):
        self.navigationbutton.Visible = value

    @property
    def Width(self):
        return self.navigationbutton.Width

    @Width.setter
    def Width(self, value):
        self.navigationbutton.Width = value

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.navigationbutton.Move(*args, **arguments)

    def Requery(self):
        self.navigationbutton.Requery()

    def SetFocus(self):
        self.navigationbutton.SetFocus()

    def SizeToFit(self):
        self.navigationbutton.SizeToFit()

class NavigationControl:

    def __init__(self, navigationcontrol=None):
        self.navigationcontrol = navigationcontrol

    @property
    def Application(self):
        return self.navigationcontrol.Application

    @property
    def AutoTab(self):
        return self.navigationcontrol.AutoTab

    @AutoTab.setter
    def AutoTab(self, value):
        self.navigationcontrol.AutoTab = value

    @property
    def BackColor(self):
        return self.navigationcontrol.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.navigationcontrol.BackColor = value

    @property
    def BackShade(self):
        return self.navigationcontrol.BackShade

    @BackShade.setter
    def BackShade(self, value):
        self.navigationcontrol.BackShade = value

    @property
    def BackStyle(self):
        return self.navigationcontrol.BackStyle

    @BackStyle.setter
    def BackStyle(self, value):
        self.navigationcontrol.BackStyle = value

    @property
    def BackThemeColorIndex(self):
        return self.navigationcontrol.BackThemeColorIndex

    @BackThemeColorIndex.setter
    def BackThemeColorIndex(self, value):
        self.navigationcontrol.BackThemeColorIndex = value

    @property
    def BackTint(self):
        return self.navigationcontrol.BackTint

    @BackTint.setter
    def BackTint(self, value):
        self.navigationcontrol.BackTint = value

    @property
    def BorderColor(self):
        return self.navigationcontrol.BorderColor

    @BorderColor.setter
    def BorderColor(self, value):
        self.navigationcontrol.BorderColor = value

    @property
    def BorderShade(self):
        return self.navigationcontrol.BorderShade

    @BorderShade.setter
    def BorderShade(self, value):
        self.navigationcontrol.BorderShade = value

    @property
    def BorderStyle(self):
        return self.navigationcontrol.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.navigationcontrol.BorderStyle = value

    @property
    def BorderThemeColorIndex(self):
        return self.navigationcontrol.BorderThemeColorIndex

    @BorderThemeColorIndex.setter
    def BorderThemeColorIndex(self, value):
        self.navigationcontrol.BorderThemeColorIndex = value

    @property
    def BorderTint(self):
        return self.navigationcontrol.BorderTint

    @BorderTint.setter
    def BorderTint(self, value):
        self.navigationcontrol.BorderTint = value

    @property
    def BorderWidth(self):
        return self.navigationcontrol.BorderWidth

    @BorderWidth.setter
    def BorderWidth(self, value):
        self.navigationcontrol.BorderWidth = value

    @property
    def BottomPadding(self):
        return self.navigationcontrol.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.navigationcontrol.BottomPadding = value

    @property
    def Controls(self):
        return Controls(self.navigationcontrol.Controls)

    @property
    def ControlTipText(self):
        return self.navigationcontrol.ControlTipText

    @ControlTipText.setter
    def ControlTipText(self, value):
        self.navigationcontrol.ControlTipText = value

    @property
    def ControlType(self):
        return self.navigationcontrol.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.navigationcontrol.ControlType = value

    @property
    def DisplayWhen(self):
        return self.navigationcontrol.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.navigationcontrol.DisplayWhen = value

    @property
    def Enabled(self):
        return self.navigationcontrol.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.navigationcontrol.Enabled = value

    @property
    def EventProcPrefix(self):
        return self.navigationcontrol.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.navigationcontrol.EventProcPrefix = value

    @property
    def FilterLookup(self):
        return self.navigationcontrol.FilterLookup

    @FilterLookup.setter
    def FilterLookup(self, value):
        self.navigationcontrol.FilterLookup = value

    @property
    def FormatConditions(self):
        return self.navigationcontrol.FormatConditions

    @property
    def GridlineColor(self):
        return self.navigationcontrol.GridlineColor

    @GridlineColor.setter
    def GridlineColor(self, value):
        self.navigationcontrol.GridlineColor = value

    @property
    def GridlineShade(self):
        return self.navigationcontrol.GridlineShade

    @GridlineShade.setter
    def GridlineShade(self, value):
        self.navigationcontrol.GridlineShade = value

    @property
    def GridlineStyleBottom(self):
        return self.navigationcontrol.GridlineStyleBottom

    @GridlineStyleBottom.setter
    def GridlineStyleBottom(self, value):
        self.navigationcontrol.GridlineStyleBottom = value

    @property
    def GridlineStyleLeft(self):
        return self.navigationcontrol.GridlineStyleLeft

    @GridlineStyleLeft.setter
    def GridlineStyleLeft(self, value):
        self.navigationcontrol.GridlineStyleLeft = value

    @property
    def GridlineStyleRight(self):
        return self.navigationcontrol.GridlineStyleRight

    @GridlineStyleRight.setter
    def GridlineStyleRight(self, value):
        self.navigationcontrol.GridlineStyleRight = value

    @property
    def GridlineStyleTop(self):
        return self.navigationcontrol.GridlineStyleTop

    @GridlineStyleTop.setter
    def GridlineStyleTop(self, value):
        self.navigationcontrol.GridlineStyleTop = value

    @property
    def GridlineThemeColorIndex(self):
        return self.navigationcontrol.GridlineThemeColorIndex

    @GridlineThemeColorIndex.setter
    def GridlineThemeColorIndex(self, value):
        self.navigationcontrol.GridlineThemeColorIndex = value

    @property
    def GridlineTint(self):
        return self.navigationcontrol.GridlineTint

    @GridlineTint.setter
    def GridlineTint(self, value):
        self.navigationcontrol.GridlineTint = value

    @property
    def GridlineWidthBottom(self):
        return self.navigationcontrol.GridlineWidthBottom

    @GridlineWidthBottom.setter
    def GridlineWidthBottom(self, value):
        self.navigationcontrol.GridlineWidthBottom = value

    @property
    def GridlineWidthLeft(self):
        return self.navigationcontrol.GridlineWidthLeft

    @GridlineWidthLeft.setter
    def GridlineWidthLeft(self, value):
        self.navigationcontrol.GridlineWidthLeft = value

    @property
    def GridlineWidthRight(self):
        return self.navigationcontrol.GridlineWidthRight

    @GridlineWidthRight.setter
    def GridlineWidthRight(self, value):
        self.navigationcontrol.GridlineWidthRight = value

    @property
    def GridlineWidthTop(self):
        return self.navigationcontrol.GridlineWidthTop

    @GridlineWidthTop.setter
    def GridlineWidthTop(self, value):
        self.navigationcontrol.GridlineWidthTop = value

    @property
    def Height(self):
        return self.navigationcontrol.Height

    @Height.setter
    def Height(self, value):
        self.navigationcontrol.Height = value

    @property
    def HelpContextId(self):
        return self.navigationcontrol.HelpContextId

    @HelpContextId.setter
    def HelpContextId(self, value):
        self.navigationcontrol.HelpContextId = value

    @property
    def HorizontalAnchor(self):
        return self.navigationcontrol.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.navigationcontrol.HorizontalAnchor = value

    @property
    def Hyperlink(self):
        return self.navigationcontrol.Hyperlink

    @property
    def InSelection(self):
        return self.navigationcontrol.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.navigationcontrol.InSelection = value

    @property
    def IsVisible(self):
        return self.navigationcontrol.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.navigationcontrol.IsVisible = value

    @property
    def Layout(self):
        return AcLayoutType(self.navigationcontrol.Layout)

    @property
    def LayoutID(self):
        return self.navigationcontrol.LayoutID

    @property
    def Left(self):
        return self.navigationcontrol.Left

    @Left.setter
    def Left(self, value):
        self.navigationcontrol.Left = value

    @property
    def LeftPadding(self):
        return self.navigationcontrol.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.navigationcontrol.LeftPadding = value

    @property
    def LineSpacing(self):
        return self.navigationcontrol.LineSpacing

    @LineSpacing.setter
    def LineSpacing(self, value):
        self.navigationcontrol.LineSpacing = value

    @property
    def Name(self):
        return self.navigationcontrol.Name

    @Name.setter
    def Name(self, value):
        self.navigationcontrol.Name = value

    @property
    def OldBorderStyle(self):
        return self.navigationcontrol.OldBorderStyle

    @OldBorderStyle.setter
    def OldBorderStyle(self, value):
        self.navigationcontrol.OldBorderStyle = value

    @property
    def OldValue(self):
        return self.navigationcontrol.OldValue

    @property
    def OnClick(self):
        return self.navigationcontrol.OnClick

    @OnClick.setter
    def OnClick(self, value):
        self.navigationcontrol.OnClick = value

    @property
    def OnDblClick(self):
        return self.navigationcontrol.OnDblClick

    @OnDblClick.setter
    def OnDblClick(self, value):
        self.navigationcontrol.OnDblClick = value

    @property
    def OnGotFocus(self):
        return self.navigationcontrol.OnGotFocus

    @OnGotFocus.setter
    def OnGotFocus(self, value):
        self.navigationcontrol.OnGotFocus = value

    @property
    def OnKeyDown(self):
        return self.navigationcontrol.OnKeyDown

    @OnKeyDown.setter
    def OnKeyDown(self, value):
        self.navigationcontrol.OnKeyDown = value

    @property
    def OnKeyPress(self):
        return self.navigationcontrol.OnKeyPress

    @OnKeyPress.setter
    def OnKeyPress(self, value):
        self.navigationcontrol.OnKeyPress = value

    @property
    def OnKeyUp(self):
        return self.navigationcontrol.OnKeyUp

    @OnKeyUp.setter
    def OnKeyUp(self, value):
        self.navigationcontrol.OnKeyUp = value

    @property
    def OnLostFocus(self):
        return self.navigationcontrol.OnLostFocus

    @OnLostFocus.setter
    def OnLostFocus(self, value):
        self.navigationcontrol.OnLostFocus = value

    @property
    def OnMouseDown(self):
        return self.navigationcontrol.OnMouseDown

    @OnMouseDown.setter
    def OnMouseDown(self, value):
        self.navigationcontrol.OnMouseDown = value

    @property
    def OnMouseMove(self):
        return self.navigationcontrol.OnMouseMove

    @OnMouseMove.setter
    def OnMouseMove(self, value):
        self.navigationcontrol.OnMouseMove = value

    @property
    def OnMouseUp(self):
        return self.navigationcontrol.OnMouseUp

    @OnMouseUp.setter
    def OnMouseUp(self, value):
        self.navigationcontrol.OnMouseUp = value

    @property
    def Parent(self):
        return self.navigationcontrol.Parent

    @property
    def Properties(self):
        return Properties(self.navigationcontrol.Properties)

    @property
    def ReadingOrder(self):
        return self.navigationcontrol.ReadingOrder

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.navigationcontrol.ReadingOrder = value

    @property
    def RightPadding(self):
        return self.navigationcontrol.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.navigationcontrol.RightPadding = value

    @property
    def ScrollBarAlign(self):
        return self.navigationcontrol.ScrollBarAlign

    @ScrollBarAlign.setter
    def ScrollBarAlign(self, value):
        self.navigationcontrol.ScrollBarAlign = value

    @property
    def Section(self):
        return self.navigationcontrol.Section

    @Section.setter
    def Section(self, value):
        self.navigationcontrol.Section = value

    @property
    def SelectedTab(self):
        return self.navigationcontrol.SelectedTab

    @property
    def ShortcutMenuBar(self):
        return self.navigationcontrol.ShortcutMenuBar

    @ShortcutMenuBar.setter
    def ShortcutMenuBar(self, value):
        self.navigationcontrol.ShortcutMenuBar = value

    @property
    def SmartTags(self):
        return SmartTags(self.navigationcontrol.SmartTags)

    @property
    def Span(self):
        return self.navigationcontrol.Span

    @Span.setter
    def Span(self, value):
        self.navigationcontrol.Span = value

    @property
    def SpecialEffect(self):
        return self.navigationcontrol.SpecialEffect

    @SpecialEffect.setter
    def SpecialEffect(self, value):
        self.navigationcontrol.SpecialEffect = value

    @property
    def StatusBarText(self):
        return self.navigationcontrol.StatusBarText

    @StatusBarText.setter
    def StatusBarText(self, value):
        self.navigationcontrol.StatusBarText = value

    @property
    def SubForm(self):
        return self.navigationcontrol.SubForm

    @SubForm.setter
    def SubForm(self, value):
        self.navigationcontrol.SubForm = value

    @property
    def TabIndex(self):
        return self.navigationcontrol.TabIndex

    @TabIndex.setter
    def TabIndex(self, value):
        self.navigationcontrol.TabIndex = value

    @property
    def Tabs(self):
        return self.navigationcontrol.Tabs

    @property
    def TabStop(self):
        return self.navigationcontrol.TabStop

    @TabStop.setter
    def TabStop(self, value):
        self.navigationcontrol.TabStop = value

    @property
    def Tag(self):
        return self.navigationcontrol.Tag

    @Tag.setter
    def Tag(self, value):
        self.navigationcontrol.Tag = value

    @property
    def Top(self):
        return self.navigationcontrol.Top

    @Top.setter
    def Top(self, value):
        self.navigationcontrol.Top = value

    @property
    def TopPadding(self):
        return self.navigationcontrol.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.navigationcontrol.TopPadding = value

    @property
    def Value(self):
        return self.navigationcontrol.Value

    @Value.setter
    def Value(self, value):
        self.navigationcontrol.Value = value

    @property
    def VerticalAnchor(self):
        return self.navigationcontrol.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.navigationcontrol.VerticalAnchor = value

    @property
    def Visible(self):
        return self.navigationcontrol.Visible

    @Visible.setter
    def Visible(self, value):
        self.navigationcontrol.Visible = value

    @property
    def Width(self):
        return self.navigationcontrol.Width

    @Width.setter
    def Width(self, value):
        self.navigationcontrol.Width = value

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.navigationcontrol.Move(*args, **arguments)

    def Requery(self):
        self.navigationcontrol.Requery()

    def SetFocus(self):
        self.navigationcontrol.SetFocus()

    def SizeToFit(self):
        self.navigationcontrol.SizeToFit()

    def Undo(self):
        self.navigationcontrol.Undo()

class ObjectFrame:

    def __init__(self, objectframe=None):
        self.objectframe = objectframe

    @property
    def Action(self):
        return self.objectframe.Action

    @Action.setter
    def Action(self, value):
        self.objectframe.Action = value

    @property
    def Application(self):
        return self.objectframe.Application

    @property
    def AutoActivate(self):
        return self.objectframe.AutoActivate

    @AutoActivate.setter
    def AutoActivate(self, value):
        self.objectframe.AutoActivate = value

    @property
    def BackColor(self):
        return self.objectframe.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.objectframe.BackColor = value

    @property
    def BackShade(self):
        return self.objectframe.BackShade

    @BackShade.setter
    def BackShade(self, value):
        self.objectframe.BackShade = value

    @property
    def BackStyle(self):
        return self.objectframe.BackStyle

    @BackStyle.setter
    def BackStyle(self, value):
        self.objectframe.BackStyle = value

    @property
    def BackThemeColorIndex(self):
        return self.objectframe.BackThemeColorIndex

    @BackThemeColorIndex.setter
    def BackThemeColorIndex(self, value):
        self.objectframe.BackThemeColorIndex = value

    @property
    def BackTint(self):
        return self.objectframe.BackTint

    @BackTint.setter
    def BackTint(self, value):
        self.objectframe.BackTint = value

    @property
    def BorderColor(self):
        return self.objectframe.BorderColor

    @BorderColor.setter
    def BorderColor(self, value):
        self.objectframe.BorderColor = value

    @property
    def BorderShade(self):
        return self.objectframe.BorderShade

    @BorderShade.setter
    def BorderShade(self, value):
        self.objectframe.BorderShade = value

    @property
    def BorderStyle(self):
        return self.objectframe.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.objectframe.BorderStyle = value

    @property
    def BorderThemeColorIndex(self):
        return self.objectframe.BorderThemeColorIndex

    @BorderThemeColorIndex.setter
    def BorderThemeColorIndex(self, value):
        self.objectframe.BorderThemeColorIndex = value

    @property
    def BorderTint(self):
        return self.objectframe.BorderTint

    @BorderTint.setter
    def BorderTint(self, value):
        self.objectframe.BorderTint = value

    @property
    def BorderWidth(self):
        return self.objectframe.BorderWidth

    @BorderWidth.setter
    def BorderWidth(self, value):
        self.objectframe.BorderWidth = value

    @property
    def BottomPadding(self):
        return self.objectframe.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.objectframe.BottomPadding = value

    @property
    def Class(self):
        return self.objectframe.Class

    @Class.setter
    def Class(self, value):
        self.objectframe.Class = value

    @property
    def ColumnCount(self):
        return self.objectframe.ColumnCount

    @ColumnCount.setter
    def ColumnCount(self, value):
        self.objectframe.ColumnCount = value

    @property
    def ColumnHeads(self):
        return self.objectframe.ColumnHeads

    @ColumnHeads.setter
    def ColumnHeads(self, value):
        self.objectframe.ColumnHeads = value

    @property
    def Controls(self):
        return Controls(self.objectframe.Controls)

    @property
    def ControlTipText(self):
        return self.objectframe.ControlTipText

    @ControlTipText.setter
    def ControlTipText(self, value):
        self.objectframe.ControlTipText = value

    @property
    def ControlType(self):
        return self.objectframe.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.objectframe.ControlType = value

    @property
    def DisplayType(self):
        return self.objectframe.DisplayType

    @DisplayType.setter
    def DisplayType(self, value):
        self.objectframe.DisplayType = value

    @property
    def DisplayWhen(self):
        return self.objectframe.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.objectframe.DisplayWhen = value

    @property
    def Enabled(self):
        return self.objectframe.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.objectframe.Enabled = value

    @property
    def EventProcPrefix(self):
        return self.objectframe.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.objectframe.EventProcPrefix = value

    @property
    def GridlineColor(self):
        return self.objectframe.GridlineColor

    @GridlineColor.setter
    def GridlineColor(self, value):
        self.objectframe.GridlineColor = value

    @property
    def GridlineShade(self):
        return self.objectframe.GridlineShade

    @GridlineShade.setter
    def GridlineShade(self, value):
        self.objectframe.GridlineShade = value

    @property
    def GridlineStyleBottom(self):
        return self.objectframe.GridlineStyleBottom

    @GridlineStyleBottom.setter
    def GridlineStyleBottom(self, value):
        self.objectframe.GridlineStyleBottom = value

    @property
    def GridlineStyleLeft(self):
        return self.objectframe.GridlineStyleLeft

    @GridlineStyleLeft.setter
    def GridlineStyleLeft(self, value):
        self.objectframe.GridlineStyleLeft = value

    @property
    def GridlineStyleRight(self):
        return self.objectframe.GridlineStyleRight

    @GridlineStyleRight.setter
    def GridlineStyleRight(self, value):
        self.objectframe.GridlineStyleRight = value

    @property
    def GridlineStyleTop(self):
        return self.objectframe.GridlineStyleTop

    @GridlineStyleTop.setter
    def GridlineStyleTop(self, value):
        self.objectframe.GridlineStyleTop = value

    @property
    def GridlineThemeColorIndex(self):
        return self.objectframe.GridlineThemeColorIndex

    @GridlineThemeColorIndex.setter
    def GridlineThemeColorIndex(self, value):
        self.objectframe.GridlineThemeColorIndex = value

    @property
    def GridlineTint(self):
        return self.objectframe.GridlineTint

    @GridlineTint.setter
    def GridlineTint(self, value):
        self.objectframe.GridlineTint = value

    @property
    def GridlineWidthBottom(self):
        return self.objectframe.GridlineWidthBottom

    @GridlineWidthBottom.setter
    def GridlineWidthBottom(self, value):
        self.objectframe.GridlineWidthBottom = value

    @property
    def GridlineWidthLeft(self):
        return self.objectframe.GridlineWidthLeft

    @GridlineWidthLeft.setter
    def GridlineWidthLeft(self, value):
        self.objectframe.GridlineWidthLeft = value

    @property
    def GridlineWidthRight(self):
        return self.objectframe.GridlineWidthRight

    @GridlineWidthRight.setter
    def GridlineWidthRight(self, value):
        self.objectframe.GridlineWidthRight = value

    @property
    def GridlineWidthTop(self):
        return self.objectframe.GridlineWidthTop

    @GridlineWidthTop.setter
    def GridlineWidthTop(self, value):
        self.objectframe.GridlineWidthTop = value

    @property
    def Height(self):
        return self.objectframe.Height

    @Height.setter
    def Height(self, value):
        self.objectframe.Height = value

    @property
    def HelpContextId(self):
        return self.objectframe.HelpContextId

    @HelpContextId.setter
    def HelpContextId(self, value):
        self.objectframe.HelpContextId = value

    @property
    def HorizontalAnchor(self):
        return self.objectframe.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.objectframe.HorizontalAnchor = value

    @property
    def InSelection(self):
        return self.objectframe.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.objectframe.InSelection = value

    @property
    def IsVisible(self):
        return self.objectframe.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.objectframe.IsVisible = value

    @property
    def Item(self):
        return self.objectframe.Item

    @Item.setter
    def Item(self, value):
        self.objectframe.Item = value

    @property
    def Layout(self):
        return AcLayoutType(self.objectframe.Layout)

    @property
    def LayoutID(self):
        return self.objectframe.LayoutID

    @property
    def Left(self):
        return self.objectframe.Left

    @Left.setter
    def Left(self, value):
        self.objectframe.Left = value

    @property
    def LeftPadding(self):
        return self.objectframe.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.objectframe.LeftPadding = value

    @property
    def LinkChildFields(self):
        return self.objectframe.LinkChildFields

    @LinkChildFields.setter
    def LinkChildFields(self, value):
        self.objectframe.LinkChildFields = value

    @property
    def LinkMasterFields(self):
        return self.objectframe.LinkMasterFields

    @LinkMasterFields.setter
    def LinkMasterFields(self, value):
        self.objectframe.LinkMasterFields = value

    @property
    def Locked(self):
        return self.objectframe.Locked

    @Locked.setter
    def Locked(self, value):
        self.objectframe.Locked = value

    @property
    def Name(self):
        return self.objectframe.Name

    @Name.setter
    def Name(self, value):
        self.objectframe.Name = value

    @property
    def Object(self):
        return self.objectframe.Object

    @property
    def ObjectPalette(self):
        return self.objectframe.ObjectPalette

    @ObjectPalette.setter
    def ObjectPalette(self, value):
        self.objectframe.ObjectPalette = value

    def ObjectVerbs(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.objectframe.ObjectVerbs):
            return self.objectframe.ObjectVerbs(*args, **arguments)
        else:
            return self.objectframe.GetObjectVerbs(*args, **arguments)

    @property
    def ObjectVerbsCount(self):
        return self.objectframe.ObjectVerbsCount

    @property
    def OldBorderStyle(self):
        return self.objectframe.OldBorderStyle

    @OldBorderStyle.setter
    def OldBorderStyle(self, value):
        self.objectframe.OldBorderStyle = value

    @property
    def OldValue(self):
        return self.objectframe.OldValue

    @property
    def OLEClass(self):
        return self.objectframe.OLEClass

    @property
    def OLEType(self):
        return self.objectframe.OLEType

    @OLEType.setter
    def OLEType(self, value):
        self.objectframe.OLEType = value

    @property
    def OLETypeAllowed(self):
        return self.objectframe.OLETypeAllowed

    @OLETypeAllowed.setter
    def OLETypeAllowed(self, value):
        self.objectframe.OLETypeAllowed = value

    @property
    def OnClick(self):
        return self.objectframe.OnClick

    @OnClick.setter
    def OnClick(self, value):
        self.objectframe.OnClick = value

    @property
    def OnDblClick(self):
        return self.objectframe.OnDblClick

    @OnDblClick.setter
    def OnDblClick(self, value):
        self.objectframe.OnDblClick = value

    @property
    def OnEnter(self):
        return self.objectframe.OnEnter

    @OnEnter.setter
    def OnEnter(self, value):
        self.objectframe.OnEnter = value

    @property
    def OnExit(self):
        return self.objectframe.OnExit

    @OnExit.setter
    def OnExit(self, value):
        self.objectframe.OnExit = value

    @property
    def OnGotFocus(self):
        return self.objectframe.OnGotFocus

    @OnGotFocus.setter
    def OnGotFocus(self, value):
        self.objectframe.OnGotFocus = value

    @property
    def OnLostFocus(self):
        return self.objectframe.OnLostFocus

    @OnLostFocus.setter
    def OnLostFocus(self, value):
        self.objectframe.OnLostFocus = value

    @property
    def OnMouseDown(self):
        return self.objectframe.OnMouseDown

    @OnMouseDown.setter
    def OnMouseDown(self, value):
        self.objectframe.OnMouseDown = value

    @property
    def OnMouseMove(self):
        return self.objectframe.OnMouseMove

    @OnMouseMove.setter
    def OnMouseMove(self, value):
        self.objectframe.OnMouseMove = value

    @property
    def OnMouseUp(self):
        return self.objectframe.OnMouseUp

    @OnMouseUp.setter
    def OnMouseUp(self, value):
        self.objectframe.OnMouseUp = value

    @property
    def OnUpdated(self):
        return self.objectframe.OnUpdated

    @OnUpdated.setter
    def OnUpdated(self, value):
        self.objectframe.OnUpdated = value

    @property
    def Parent(self):
        return self.objectframe.Parent

    @property
    def Properties(self):
        return Properties(self.objectframe.Properties)

    @property
    def RightPadding(self):
        return self.objectframe.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.objectframe.RightPadding = value

    @property
    def RowSource(self):
        return self.objectframe.RowSource

    @RowSource.setter
    def RowSource(self, value):
        self.objectframe.RowSource = value

    @property
    def RowSourceType(self):
        return self.objectframe.RowSourceType

    @RowSourceType.setter
    def RowSourceType(self, value):
        self.objectframe.RowSourceType = value

    @property
    def Scaling(self):
        return self.objectframe.Scaling

    @Scaling.setter
    def Scaling(self, value):
        self.objectframe.Scaling = value

    @property
    def Section(self):
        return self.objectframe.Section

    @Section.setter
    def Section(self, value):
        self.objectframe.Section = value

    @property
    def ShortcutMenuBar(self):
        return self.objectframe.ShortcutMenuBar

    @ShortcutMenuBar.setter
    def ShortcutMenuBar(self, value):
        self.objectframe.ShortcutMenuBar = value

    @property
    def SizeMode(self):
        return self.objectframe.SizeMode

    @property
    def SourceDoc(self):
        return self.objectframe.SourceDoc

    @SourceDoc.setter
    def SourceDoc(self, value):
        self.objectframe.SourceDoc = value

    @property
    def SourceItem(self):
        return self.objectframe.SourceItem

    @SourceItem.setter
    def SourceItem(self, value):
        self.objectframe.SourceItem = value

    @property
    def SourceObject(self):
        return self.objectframe.SourceObject

    @property
    def SpecialEffect(self):
        return self.objectframe.SpecialEffect

    @SpecialEffect.setter
    def SpecialEffect(self, value):
        self.objectframe.SpecialEffect = value

    @property
    def StatusBarText(self):
        return self.objectframe.StatusBarText

    @StatusBarText.setter
    def StatusBarText(self, value):
        self.objectframe.StatusBarText = value

    @property
    def TabIndex(self):
        return self.objectframe.TabIndex

    @TabIndex.setter
    def TabIndex(self, value):
        self.objectframe.TabIndex = value

    @property
    def TabStop(self):
        return self.objectframe.TabStop

    @TabStop.setter
    def TabStop(self, value):
        self.objectframe.TabStop = value

    @property
    def Tag(self):
        return self.objectframe.Tag

    @Tag.setter
    def Tag(self, value):
        self.objectframe.Tag = value

    @property
    def Top(self):
        return self.objectframe.Top

    @Top.setter
    def Top(self, value):
        self.objectframe.Top = value

    @property
    def TopPadding(self):
        return self.objectframe.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.objectframe.TopPadding = value

    @property
    def UpdateMethod(self):
        return self.objectframe.UpdateMethod

    @property
    def UpdateOptions(self):
        return self.objectframe.UpdateOptions

    @UpdateOptions.setter
    def UpdateOptions(self, value):
        self.objectframe.UpdateOptions = value

    @property
    def VarOleObject(self):
        return self.objectframe.VarOleObject

    @property
    def Verb(self):
        return self.objectframe.Verb

    @Verb.setter
    def Verb(self, value):
        self.objectframe.Verb = value

    @property
    def VerticalAnchor(self):
        return self.objectframe.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.objectframe.VerticalAnchor = value

    @property
    def Visible(self):
        return self.objectframe.Visible

    @Visible.setter
    def Visible(self, value):
        self.objectframe.Visible = value

    @property
    def Width(self):
        return self.objectframe.Width

    @Width.setter
    def Width(self, value):
        self.objectframe.Width = value

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.objectframe.Move(*args, **arguments)

    def Requery(self):
        self.objectframe.Requery()

    def SetFocus(self):
        return self.objectframe.SetFocus()

    def SizeToFit(self):
        self.objectframe.SizeToFit()

class Operation:

    def __init__(self, operation=None):
        self.operation = operation

    @property
    def Name(self):
        return self.operation.Name

    @Name.setter
    def Name(self, value):
        self.operation.Name = value

    @property
    def Parent(self):
        return self.operation.Parent

    @property
    def WSParameters(self):
        return self.operation.WSParameters

    def Execute(self, *args, bstrParameters=None):
        arguments = {"bstrParameters": bstrParameters}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.operation.Execute(*args, **arguments)

class Operations:

    def __init__(self, operations=None):
        self.operations = operations

    @property
    def Count(self):
        return self.operations.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.operations.Item):
            return self.operations.Item(*args, **arguments)
        else:
            return self.operations.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.operations.Parent

class OptionButton:

    def __init__(self, optionbutton=None):
        self.optionbutton = optionbutton

    @property
    def AddColon(self):
        return self.optionbutton.AddColon

    @AddColon.setter
    def AddColon(self, value):
        self.optionbutton.AddColon = value

    @property
    def Application(self):
        return self.optionbutton.Application

    @property
    def AutoLabel(self):
        return self.optionbutton.AutoLabel

    @AutoLabel.setter
    def AutoLabel(self, value):
        self.optionbutton.AutoLabel = value

    @property
    def BorderColor(self):
        return self.optionbutton.BorderColor

    @BorderColor.setter
    def BorderColor(self, value):
        self.optionbutton.BorderColor = value

    @property
    def BorderShade(self):
        return self.optionbutton.BorderShade

    @BorderShade.setter
    def BorderShade(self, value):
        self.optionbutton.BorderShade = value

    @property
    def BorderStyle(self):
        return self.optionbutton.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.optionbutton.BorderStyle = value

    @property
    def BorderThemeColorIndex(self):
        return self.optionbutton.BorderThemeColorIndex

    @BorderThemeColorIndex.setter
    def BorderThemeColorIndex(self, value):
        self.optionbutton.BorderThemeColorIndex = value

    @property
    def BorderTint(self):
        return self.optionbutton.BorderTint

    @BorderTint.setter
    def BorderTint(self, value):
        self.optionbutton.BorderTint = value

    @property
    def BorderWidth(self):
        return self.optionbutton.BorderWidth

    @BorderWidth.setter
    def BorderWidth(self, value):
        self.optionbutton.BorderWidth = value

    @property
    def BottomPadding(self):
        return self.optionbutton.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.optionbutton.BottomPadding = value

    @property
    def ColumnHidden(self):
        return self.optionbutton.ColumnHidden

    @ColumnHidden.setter
    def ColumnHidden(self, value):
        self.optionbutton.ColumnHidden = value

    @property
    def ColumnOrder(self):
        return self.optionbutton.ColumnOrder

    @ColumnOrder.setter
    def ColumnOrder(self, value):
        self.optionbutton.ColumnOrder = value

    @property
    def ColumnWidth(self):
        return self.optionbutton.ColumnWidth

    @ColumnWidth.setter
    def ColumnWidth(self, value):
        self.optionbutton.ColumnWidth = value

    @property
    def Controls(self):
        return Controls(self.optionbutton.Controls)

    @property
    def ControlSource(self):
        return self.optionbutton.ControlSource

    @ControlSource.setter
    def ControlSource(self, value):
        self.optionbutton.ControlSource = value

    @property
    def ControlTipText(self):
        return self.optionbutton.ControlTipText

    @ControlTipText.setter
    def ControlTipText(self, value):
        self.optionbutton.ControlTipText = value

    @property
    def ControlType(self):
        return self.optionbutton.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.optionbutton.ControlType = value

    @property
    def DefaultValue(self):
        return self.optionbutton.DefaultValue

    @DefaultValue.setter
    def DefaultValue(self, value):
        self.optionbutton.DefaultValue = value

    @property
    def DisplayWhen(self):
        return self.optionbutton.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.optionbutton.DisplayWhen = value

    @property
    def Enabled(self):
        return self.optionbutton.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.optionbutton.Enabled = value

    @property
    def EventProcPrefix(self):
        return self.optionbutton.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.optionbutton.EventProcPrefix = value

    @property
    def GridlineColor(self):
        return self.optionbutton.GridlineColor

    @GridlineColor.setter
    def GridlineColor(self, value):
        self.optionbutton.GridlineColor = value

    @property
    def GridlineShade(self):
        return self.optionbutton.GridlineShade

    @GridlineShade.setter
    def GridlineShade(self, value):
        self.optionbutton.GridlineShade = value

    @property
    def GridlineStyleBottom(self):
        return self.optionbutton.GridlineStyleBottom

    @GridlineStyleBottom.setter
    def GridlineStyleBottom(self, value):
        self.optionbutton.GridlineStyleBottom = value

    @property
    def GridlineStyleLeft(self):
        return self.optionbutton.GridlineStyleLeft

    @GridlineStyleLeft.setter
    def GridlineStyleLeft(self, value):
        self.optionbutton.GridlineStyleLeft = value

    @property
    def GridlineStyleRight(self):
        return self.optionbutton.GridlineStyleRight

    @GridlineStyleRight.setter
    def GridlineStyleRight(self, value):
        self.optionbutton.GridlineStyleRight = value

    @property
    def GridlineStyleTop(self):
        return self.optionbutton.GridlineStyleTop

    @GridlineStyleTop.setter
    def GridlineStyleTop(self, value):
        self.optionbutton.GridlineStyleTop = value

    @property
    def GridlineThemeColorIndex(self):
        return self.optionbutton.GridlineThemeColorIndex

    @GridlineThemeColorIndex.setter
    def GridlineThemeColorIndex(self, value):
        self.optionbutton.GridlineThemeColorIndex = value

    @property
    def GridlineTint(self):
        return self.optionbutton.GridlineTint

    @GridlineTint.setter
    def GridlineTint(self, value):
        self.optionbutton.GridlineTint = value

    @property
    def GridlineWidthBottom(self):
        return self.optionbutton.GridlineWidthBottom

    @GridlineWidthBottom.setter
    def GridlineWidthBottom(self, value):
        self.optionbutton.GridlineWidthBottom = value

    @property
    def GridlineWidthLeft(self):
        return self.optionbutton.GridlineWidthLeft

    @GridlineWidthLeft.setter
    def GridlineWidthLeft(self, value):
        self.optionbutton.GridlineWidthLeft = value

    @property
    def GridlineWidthRight(self):
        return self.optionbutton.GridlineWidthRight

    @GridlineWidthRight.setter
    def GridlineWidthRight(self, value):
        self.optionbutton.GridlineWidthRight = value

    @property
    def GridlineWidthTop(self):
        return self.optionbutton.GridlineWidthTop

    @GridlineWidthTop.setter
    def GridlineWidthTop(self, value):
        self.optionbutton.GridlineWidthTop = value

    @property
    def Height(self):
        return self.optionbutton.Height

    @Height.setter
    def Height(self, value):
        self.optionbutton.Height = value

    @property
    def HelpContextId(self):
        return self.optionbutton.HelpContextId

    @HelpContextId.setter
    def HelpContextId(self, value):
        self.optionbutton.HelpContextId = value

    @property
    def HideDuplicates(self):
        return self.optionbutton.HideDuplicates

    @HideDuplicates.setter
    def HideDuplicates(self, value):
        self.optionbutton.HideDuplicates = value

    @property
    def HorizontalAnchor(self):
        return self.optionbutton.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.optionbutton.HorizontalAnchor = value

    @property
    def InSelection(self):
        return self.optionbutton.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.optionbutton.InSelection = value

    @property
    def IsVisible(self):
        return self.optionbutton.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.optionbutton.IsVisible = value

    @property
    def LabelAlign(self):
        return self.optionbutton.LabelAlign

    @LabelAlign.setter
    def LabelAlign(self, value):
        self.optionbutton.LabelAlign = value

    @property
    def LabelX(self):
        return self.optionbutton.LabelX

    @LabelX.setter
    def LabelX(self, value):
        self.optionbutton.LabelX = value

    @property
    def LabelY(self):
        return self.optionbutton.LabelY

    @LabelY.setter
    def LabelY(self, value):
        self.optionbutton.LabelY = value

    @property
    def Layout(self):
        return AcLayoutType(self.optionbutton.Layout)

    @property
    def LayoutID(self):
        return self.optionbutton.LayoutID

    @property
    def Left(self):
        return self.optionbutton.Left

    @Left.setter
    def Left(self, value):
        self.optionbutton.Left = value

    @property
    def LeftPadding(self):
        return self.optionbutton.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.optionbutton.LeftPadding = value

    @property
    def Locked(self):
        return self.optionbutton.Locked

    @Locked.setter
    def Locked(self, value):
        self.optionbutton.Locked = value

    @property
    def Name(self):
        return self.optionbutton.Name

    @Name.setter
    def Name(self, value):
        self.optionbutton.Name = value

    @property
    def OldBorderStyle(self):
        return self.optionbutton.OldBorderStyle

    @OldBorderStyle.setter
    def OldBorderStyle(self, value):
        self.optionbutton.OldBorderStyle = value

    @property
    def OldValue(self):
        return self.optionbutton.OldValue

    @property
    def OnClick(self):
        return self.optionbutton.OnClick

    @OnClick.setter
    def OnClick(self, value):
        self.optionbutton.OnClick = value

    @property
    def OnDblClick(self):
        return self.optionbutton.OnDblClick

    @OnDblClick.setter
    def OnDblClick(self, value):
        self.optionbutton.OnDblClick = value

    @property
    def OnEnter(self):
        return self.optionbutton.OnEnter

    @OnEnter.setter
    def OnEnter(self, value):
        self.optionbutton.OnEnter = value

    @property
    def OnExit(self):
        return self.optionbutton.OnExit

    @OnExit.setter
    def OnExit(self, value):
        self.optionbutton.OnExit = value

    @property
    def OnGotFocus(self):
        return self.optionbutton.OnGotFocus

    @OnGotFocus.setter
    def OnGotFocus(self, value):
        self.optionbutton.OnGotFocus = value

    @property
    def OnKeyDown(self):
        return self.optionbutton.OnKeyDown

    @OnKeyDown.setter
    def OnKeyDown(self, value):
        self.optionbutton.OnKeyDown = value

    @property
    def OnKeyPress(self):
        return self.optionbutton.OnKeyPress

    @OnKeyPress.setter
    def OnKeyPress(self, value):
        self.optionbutton.OnKeyPress = value

    @property
    def OnKeyUp(self):
        return self.optionbutton.OnKeyUp

    @OnKeyUp.setter
    def OnKeyUp(self, value):
        self.optionbutton.OnKeyUp = value

    @property
    def OnLostFocus(self):
        return self.optionbutton.OnLostFocus

    @OnLostFocus.setter
    def OnLostFocus(self, value):
        self.optionbutton.OnLostFocus = value

    @property
    def OnMouseDown(self):
        return self.optionbutton.OnMouseDown

    @OnMouseDown.setter
    def OnMouseDown(self, value):
        self.optionbutton.OnMouseDown = value

    @property
    def OnMouseMove(self):
        return self.optionbutton.OnMouseMove

    @OnMouseMove.setter
    def OnMouseMove(self, value):
        self.optionbutton.OnMouseMove = value

    @property
    def OnMouseUp(self):
        return self.optionbutton.OnMouseUp

    @OnMouseUp.setter
    def OnMouseUp(self, value):
        self.optionbutton.OnMouseUp = value

    @property
    def OptionValue(self):
        return self.optionbutton.OptionValue

    @OptionValue.setter
    def OptionValue(self, value):
        self.optionbutton.OptionValue = value

    @property
    def Parent(self):
        return self.optionbutton.Parent

    @property
    def Properties(self):
        return Properties(self.optionbutton.Properties)

    @property
    def ReadingOrder(self):
        return self.optionbutton.ReadingOrder

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.optionbutton.ReadingOrder = value

    @property
    def RightPadding(self):
        return self.optionbutton.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.optionbutton.RightPadding = value

    @property
    def Section(self):
        return self.optionbutton.Section

    @Section.setter
    def Section(self, value):
        self.optionbutton.Section = value

    @property
    def ShortcutMenuBar(self):
        return self.optionbutton.ShortcutMenuBar

    @ShortcutMenuBar.setter
    def ShortcutMenuBar(self, value):
        self.optionbutton.ShortcutMenuBar = value

    @property
    def SpecialEffect(self):
        return self.optionbutton.SpecialEffect

    @SpecialEffect.setter
    def SpecialEffect(self, value):
        self.optionbutton.SpecialEffect = value

    @property
    def StatusBarText(self):
        return self.optionbutton.StatusBarText

    @StatusBarText.setter
    def StatusBarText(self, value):
        self.optionbutton.StatusBarText = value

    @property
    def TabIndex(self):
        return self.optionbutton.TabIndex

    @TabIndex.setter
    def TabIndex(self, value):
        self.optionbutton.TabIndex = value

    @property
    def TabStop(self):
        return self.optionbutton.TabStop

    @TabStop.setter
    def TabStop(self, value):
        self.optionbutton.TabStop = value

    @property
    def Tag(self):
        return self.optionbutton.Tag

    @Tag.setter
    def Tag(self, value):
        self.optionbutton.Tag = value

    @property
    def Top(self):
        return self.optionbutton.Top

    @Top.setter
    def Top(self, value):
        self.optionbutton.Top = value

    @property
    def TopPadding(self):
        return self.optionbutton.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.optionbutton.TopPadding = value

    @property
    def TripleState(self):
        return self.optionbutton.TripleState

    @TripleState.setter
    def TripleState(self, value):
        self.optionbutton.TripleState = value

    @property
    def ValidationRule(self):
        return self.optionbutton.ValidationRule

    @ValidationRule.setter
    def ValidationRule(self, value):
        self.optionbutton.ValidationRule = value

    @property
    def ValidationText(self):
        return self.optionbutton.ValidationText

    @ValidationText.setter
    def ValidationText(self, value):
        self.optionbutton.ValidationText = value

    @property
    def Value(self):
        return self.optionbutton.Value

    @Value.setter
    def Value(self, value):
        self.optionbutton.Value = value

    @property
    def VerticalAnchor(self):
        return self.optionbutton.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.optionbutton.VerticalAnchor = value

    @property
    def Visible(self):
        return self.optionbutton.Visible

    @Visible.setter
    def Visible(self, value):
        self.optionbutton.Visible = value

    @property
    def Width(self):
        return self.optionbutton.Width

    @Width.setter
    def Width(self, value):
        self.optionbutton.Width = value

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.optionbutton.Move(*args, **arguments)

    def Requery(self):
        self.optionbutton.Requery()

    def SetFocus(self):
        return self.optionbutton.SetFocus()

    def SizeToFit(self):
        self.optionbutton.SizeToFit()

class OptionGroup:

    def __init__(self, optiongroup=None):
        self.optiongroup = optiongroup

    @property
    def AddColon(self):
        return self.optiongroup.AddColon

    @AddColon.setter
    def AddColon(self, value):
        self.optiongroup.AddColon = value

    @property
    def Application(self):
        return self.optiongroup.Application

    @property
    def AutoLabel(self):
        return self.optiongroup.AutoLabel

    @AutoLabel.setter
    def AutoLabel(self, value):
        self.optiongroup.AutoLabel = value

    @property
    def BackColor(self):
        return self.optiongroup.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.optiongroup.BackColor = value

    @property
    def BackShade(self):
        return self.optiongroup.BackShade

    @BackShade.setter
    def BackShade(self, value):
        self.optiongroup.BackShade = value

    @property
    def BackStyle(self):
        return self.optiongroup.BackStyle

    @BackStyle.setter
    def BackStyle(self, value):
        self.optiongroup.BackStyle = value

    @property
    def BackThemeColorIndex(self):
        return self.optiongroup.BackThemeColorIndex

    @BackThemeColorIndex.setter
    def BackThemeColorIndex(self, value):
        self.optiongroup.BackThemeColorIndex = value

    @property
    def BackTint(self):
        return self.optiongroup.BackTint

    @BackTint.setter
    def BackTint(self, value):
        self.optiongroup.BackTint = value

    @property
    def BorderColor(self):
        return self.optiongroup.BorderColor

    @BorderColor.setter
    def BorderColor(self, value):
        self.optiongroup.BorderColor = value

    @property
    def BorderShade(self):
        return self.optiongroup.BorderShade

    @BorderShade.setter
    def BorderShade(self, value):
        self.optiongroup.BorderShade = value

    @property
    def BorderStyle(self):
        return self.optiongroup.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.optiongroup.BorderStyle = value

    @property
    def BorderThemeColorIndex(self):
        return self.optiongroup.BorderThemeColorIndex

    @BorderThemeColorIndex.setter
    def BorderThemeColorIndex(self, value):
        self.optiongroup.BorderThemeColorIndex = value

    @property
    def BorderTint(self):
        return self.optiongroup.BorderTint

    @BorderTint.setter
    def BorderTint(self, value):
        self.optiongroup.BorderTint = value

    @property
    def BorderWidth(self):
        return self.optiongroup.BorderWidth

    @BorderWidth.setter
    def BorderWidth(self, value):
        self.optiongroup.BorderWidth = value

    @property
    def ColumnHidden(self):
        return self.optiongroup.ColumnHidden

    @ColumnHidden.setter
    def ColumnHidden(self, value):
        self.optiongroup.ColumnHidden = value

    @property
    def ColumnOrder(self):
        return self.optiongroup.ColumnOrder

    @ColumnOrder.setter
    def ColumnOrder(self, value):
        self.optiongroup.ColumnOrder = value

    @property
    def ColumnWidth(self):
        return self.optiongroup.ColumnWidth

    @ColumnWidth.setter
    def ColumnWidth(self, value):
        self.optiongroup.ColumnWidth = value

    @property
    def Controls(self):
        return Controls(self.optiongroup.Controls)

    @property
    def ControlSource(self):
        return self.optiongroup.ControlSource

    @ControlSource.setter
    def ControlSource(self, value):
        self.optiongroup.ControlSource = value

    @property
    def ControlTipText(self):
        return self.optiongroup.ControlTipText

    @ControlTipText.setter
    def ControlTipText(self, value):
        self.optiongroup.ControlTipText = value

    @property
    def ControlType(self):
        return self.optiongroup.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.optiongroup.ControlType = value

    @property
    def DefaultValue(self):
        return self.optiongroup.DefaultValue

    @DefaultValue.setter
    def DefaultValue(self, value):
        self.optiongroup.DefaultValue = value

    @property
    def DisplayWhen(self):
        return self.optiongroup.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.optiongroup.DisplayWhen = value

    @property
    def Enabled(self):
        return self.optiongroup.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.optiongroup.Enabled = value

    @property
    def EventProcPrefix(self):
        return self.optiongroup.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.optiongroup.EventProcPrefix = value

    @property
    def Height(self):
        return self.optiongroup.Height

    @Height.setter
    def Height(self, value):
        self.optiongroup.Height = value

    @property
    def HelpContextId(self):
        return self.optiongroup.HelpContextId

    @HelpContextId.setter
    def HelpContextId(self, value):
        self.optiongroup.HelpContextId = value

    @property
    def HideDuplicates(self):
        return self.optiongroup.HideDuplicates

    @HideDuplicates.setter
    def HideDuplicates(self, value):
        self.optiongroup.HideDuplicates = value

    @property
    def HorizontalAnchor(self):
        return self.optiongroup.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.optiongroup.HorizontalAnchor = value

    @property
    def InSelection(self):
        return self.optiongroup.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.optiongroup.InSelection = value

    @property
    def IsVisible(self):
        return self.optiongroup.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.optiongroup.IsVisible = value

    @property
    def LabelAlign(self):
        return self.optiongroup.LabelAlign

    @LabelAlign.setter
    def LabelAlign(self, value):
        self.optiongroup.LabelAlign = value

    @property
    def LabelX(self):
        return self.optiongroup.LabelX

    @LabelX.setter
    def LabelX(self, value):
        self.optiongroup.LabelX = value

    @property
    def LabelY(self):
        return self.optiongroup.LabelY

    @LabelY.setter
    def LabelY(self, value):
        self.optiongroup.LabelY = value

    @property
    def Left(self):
        return self.optiongroup.Left

    @Left.setter
    def Left(self, value):
        self.optiongroup.Left = value

    @property
    def Locked(self):
        return self.optiongroup.Locked

    @Locked.setter
    def Locked(self, value):
        self.optiongroup.Locked = value

    @property
    def Name(self):
        return self.optiongroup.Name

    @Name.setter
    def Name(self, value):
        self.optiongroup.Name = value

    @property
    def OldBorderStyle(self):
        return self.optiongroup.OldBorderStyle

    @OldBorderStyle.setter
    def OldBorderStyle(self, value):
        self.optiongroup.OldBorderStyle = value

    @property
    def OldValue(self):
        return self.optiongroup.OldValue

    @property
    def OnClick(self):
        return self.optiongroup.OnClick

    @OnClick.setter
    def OnClick(self, value):
        self.optiongroup.OnClick = value

    @property
    def OnDblClick(self):
        return self.optiongroup.OnDblClick

    @OnDblClick.setter
    def OnDblClick(self, value):
        self.optiongroup.OnDblClick = value

    @property
    def OnEnter(self):
        return self.optiongroup.OnEnter

    @OnEnter.setter
    def OnEnter(self, value):
        self.optiongroup.OnEnter = value

    @property
    def OnExit(self):
        return self.optiongroup.OnExit

    @OnExit.setter
    def OnExit(self, value):
        self.optiongroup.OnExit = value

    @property
    def OnMouseDown(self):
        return self.optiongroup.OnMouseDown

    @OnMouseDown.setter
    def OnMouseDown(self, value):
        self.optiongroup.OnMouseDown = value

    @property
    def OnMouseMove(self):
        return self.optiongroup.OnMouseMove

    @OnMouseMove.setter
    def OnMouseMove(self, value):
        self.optiongroup.OnMouseMove = value

    @property
    def OnMouseUp(self):
        return self.optiongroup.OnMouseUp

    @OnMouseUp.setter
    def OnMouseUp(self, value):
        self.optiongroup.OnMouseUp = value

    @property
    def Parent(self):
        return self.optiongroup.Parent

    @property
    def Properties(self):
        return Properties(self.optiongroup.Properties)

    @property
    def Section(self):
        return self.optiongroup.Section

    @Section.setter
    def Section(self, value):
        self.optiongroup.Section = value

    @property
    def ShortcutMenuBar(self):
        return self.optiongroup.ShortcutMenuBar

    @ShortcutMenuBar.setter
    def ShortcutMenuBar(self, value):
        self.optiongroup.ShortcutMenuBar = value

    @property
    def SpecialEffect(self):
        return self.optiongroup.SpecialEffect

    @SpecialEffect.setter
    def SpecialEffect(self, value):
        self.optiongroup.SpecialEffect = value

    @property
    def StatusBarText(self):
        return self.optiongroup.StatusBarText

    @StatusBarText.setter
    def StatusBarText(self, value):
        self.optiongroup.StatusBarText = value

    @property
    def TabIndex(self):
        return self.optiongroup.TabIndex

    @TabIndex.setter
    def TabIndex(self, value):
        self.optiongroup.TabIndex = value

    @property
    def TabStop(self):
        return self.optiongroup.TabStop

    @TabStop.setter
    def TabStop(self, value):
        self.optiongroup.TabStop = value

    @property
    def Tag(self):
        return self.optiongroup.Tag

    @Tag.setter
    def Tag(self, value):
        self.optiongroup.Tag = value

    @property
    def Top(self):
        return self.optiongroup.Top

    @Top.setter
    def Top(self, value):
        self.optiongroup.Top = value

    @property
    def ValidationRule(self):
        return self.optiongroup.ValidationRule

    @ValidationRule.setter
    def ValidationRule(self, value):
        self.optiongroup.ValidationRule = value

    @property
    def ValidationText(self):
        return self.optiongroup.ValidationText

    @ValidationText.setter
    def ValidationText(self, value):
        self.optiongroup.ValidationText = value

    @property
    def Value(self):
        return self.optiongroup.Value

    @Value.setter
    def Value(self, value):
        self.optiongroup.Value = value

    @property
    def VerticalAnchor(self):
        return self.optiongroup.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.optiongroup.VerticalAnchor = value

    @property
    def Visible(self):
        return self.optiongroup.Visible

    @Visible.setter
    def Visible(self, value):
        self.optiongroup.Visible = value

    @property
    def Width(self):
        return self.optiongroup.Width

    @Width.setter
    def Width(self, value):
        self.optiongroup.Width = value

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.optiongroup.Move(*args, **arguments)

    def Requery(self):
        self.optiongroup.Requery()

    def SetFocus(self):
        return self.optiongroup.SetFocus()

    def SizeToFit(self):
        self.optiongroup.SizeToFit()

    def Undo(self):
        self.optiongroup.Undo()

class Page:

    def __init__(self, page=None):
        self.page = page

    @property
    def Application(self):
        return self.page.Application

    @property
    def Caption(self):
        return self.page.Caption

    @Caption.setter
    def Caption(self, value):
        self.page.Caption = value

    @property
    def Controls(self):
        return Controls(self.page.Controls)

    @property
    def ControlTipText(self):
        return self.page.ControlTipText

    @ControlTipText.setter
    def ControlTipText(self, value):
        self.page.ControlTipText = value

    @property
    def ControlType(self):
        return self.page.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.page.ControlType = value

    @property
    def Enabled(self):
        return self.page.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.page.Enabled = value

    @property
    def EventProcPrefix(self):
        return self.page.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.page.EventProcPrefix = value

    @property
    def Height(self):
        return self.page.Height

    @Height.setter
    def Height(self, value):
        self.page.Height = value

    @property
    def HelpContextId(self):
        return self.page.HelpContextId

    @HelpContextId.setter
    def HelpContextId(self, value):
        self.page.HelpContextId = value

    @property
    def InSelection(self):
        return self.page.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.page.InSelection = value

    @property
    def IsVisible(self):
        return self.page.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.page.IsVisible = value

    @property
    def Left(self):
        return self.page.Left

    @Left.setter
    def Left(self, value):
        self.page.Left = value

    @property
    def Name(self):
        return self.page.Name

    @Name.setter
    def Name(self, value):
        self.page.Name = value

    @property
    def OnClick(self):
        return self.page.OnClick

    @OnClick.setter
    def OnClick(self, value):
        self.page.OnClick = value

    @property
    def OnDblClick(self):
        return self.page.OnDblClick

    @OnDblClick.setter
    def OnDblClick(self, value):
        self.page.OnDblClick = value

    @property
    def OnMouseDown(self):
        return self.page.OnMouseDown

    @OnMouseDown.setter
    def OnMouseDown(self, value):
        self.page.OnMouseDown = value

    @property
    def OnMouseMove(self):
        return self.page.OnMouseMove

    @OnMouseMove.setter
    def OnMouseMove(self, value):
        self.page.OnMouseMove = value

    @property
    def OnMouseUp(self):
        return self.page.OnMouseUp

    @OnMouseUp.setter
    def OnMouseUp(self, value):
        self.page.OnMouseUp = value

    @property
    def PageIndex(self):
        return self.page.PageIndex

    @PageIndex.setter
    def PageIndex(self, value):
        self.page.PageIndex = value

    @property
    def Parent(self):
        return self.page.Parent

    @property
    def Picture(self):
        return self.page.Picture

    @Picture.setter
    def Picture(self, value):
        self.page.Picture = value

    @property
    def PictureData(self):
        return self.page.PictureData

    @PictureData.setter
    def PictureData(self, value):
        self.page.PictureData = value

    @property
    def PictureType(self):
        return self.page.PictureType

    @PictureType.setter
    def PictureType(self, value):
        self.page.PictureType = value

    @property
    def Properties(self):
        return Properties(self.page.Properties)

    @property
    def Section(self):
        return self.page.Section

    @Section.setter
    def Section(self, value):
        self.page.Section = value

    @property
    def ShortcutMenuBar(self):
        return self.page.ShortcutMenuBar

    @ShortcutMenuBar.setter
    def ShortcutMenuBar(self, value):
        self.page.ShortcutMenuBar = value

    @property
    def StatusBarText(self):
        return self.page.StatusBarText

    @StatusBarText.setter
    def StatusBarText(self, value):
        self.page.StatusBarText = value

    @property
    def Tag(self):
        return self.page.Tag

    @Tag.setter
    def Tag(self, value):
        self.page.Tag = value

    @property
    def Top(self):
        return self.page.Top

    @Top.setter
    def Top(self, value):
        self.page.Top = value

    @property
    def Visible(self):
        return self.page.Visible

    @Visible.setter
    def Visible(self, value):
        self.page.Visible = value

    @property
    def Width(self):
        return self.page.Width

    @Width.setter
    def Width(self, value):
        self.page.Width = value

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.page.Move(*args, **arguments)

    def Requery(self):
        self.page.Requery()

    def SetFocus(self):
        return self.page.SetFocus()

    def SetTabOrder(self):
        self.page.SetTabOrder()

    def SizeToFit(self):
        self.page.SizeToFit()

class PageBreak:

    def __init__(self, pagebreak=None):
        self.pagebreak = pagebreak

    @property
    def Application(self):
        return self.pagebreak.Application

    @property
    def ControlType(self):
        return self.pagebreak.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.pagebreak.ControlType = value

    @property
    def EventProcPrefix(self):
        return self.pagebreak.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.pagebreak.EventProcPrefix = value

    @property
    def InSelection(self):
        return self.pagebreak.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.pagebreak.InSelection = value

    @property
    def IsVisible(self):
        return self.pagebreak.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.pagebreak.IsVisible = value

    @property
    def Left(self):
        return self.pagebreak.Left

    @Left.setter
    def Left(self, value):
        self.pagebreak.Left = value

    @property
    def Name(self):
        return self.pagebreak.Name

    @Name.setter
    def Name(self, value):
        self.pagebreak.Name = value

    @property
    def Parent(self):
        return self.pagebreak.Parent

    @property
    def Properties(self):
        return Properties(self.pagebreak.Properties)

    @property
    def Section(self):
        return self.pagebreak.Section

    @Section.setter
    def Section(self, value):
        self.pagebreak.Section = value

    @property
    def Tag(self):
        return self.pagebreak.Tag

    @Tag.setter
    def Tag(self, value):
        self.pagebreak.Tag = value

    @property
    def Top(self):
        return self.pagebreak.Top

    @Top.setter
    def Top(self, value):
        self.pagebreak.Top = value

    @property
    def Visible(self):
        return self.pagebreak.Visible

    @Visible.setter
    def Visible(self, value):
        self.pagebreak.Visible = value

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pagebreak.Move(*args, **arguments)

    def SizeToFit(self):
        self.pagebreak.SizeToFit()

class Pages:

    def __init__(self, pages=None):
        self.pages = pages

    @property
    def Count(self):
        return self.pages.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.pages.Item):
            return self.pages.Item(*args, **arguments)
        else:
            return self.pages.GetItem(*args, **arguments)

    def Add(self, *args, Before=None):
        arguments = {"Before": Before}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.pages.Add(*args, **arguments)

    def Remove(self, *args, Item=None):
        arguments = {"Item": Item}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pages.Remove(*args, **arguments)

class Printer:

    def __init__(self, printer=None):
        self.printer = printer

    @property
    def BottomMargin(self):
        return self.printer.BottomMargin

    @BottomMargin.setter
    def BottomMargin(self, value):
        self.printer.BottomMargin = value

    @property
    def ColorMode(self):
        return AcPrintColor(self.printer.ColorMode)

    @ColorMode.setter
    def ColorMode(self, value):
        self.printer.ColorMode = value

    @property
    def ColumnSpacing(self):
        return self.printer.ColumnSpacing

    @ColumnSpacing.setter
    def ColumnSpacing(self, value):
        self.printer.ColumnSpacing = value

    @property
    def Copies(self):
        return self.printer.Copies

    @Copies.setter
    def Copies(self, value):
        self.printer.Copies = value

    @property
    def DataOnly(self):
        return self.printer.DataOnly

    @DataOnly.setter
    def DataOnly(self, value):
        self.printer.DataOnly = value

    @property
    def DefaultSize(self):
        return self.printer.DefaultSize

    @DefaultSize.setter
    def DefaultSize(self, value):
        self.printer.DefaultSize = value

    @property
    def DeviceName(self):
        return self.printer.DeviceName

    @property
    def DriverName(self):
        return self.printer.DriverName

    @property
    def Duplex(self):
        return AcPrintDuplex(self.printer.Duplex)

    @Duplex.setter
    def Duplex(self, value):
        self.printer.Duplex = value

    @property
    def ItemLayout(self):
        return AcPrintItemLayout(self.printer.ItemLayout)

    @ItemLayout.setter
    def ItemLayout(self, value):
        self.printer.ItemLayout = value

    @property
    def ItemsAcross(self):
        return self.printer.ItemsAcross

    @ItemsAcross.setter
    def ItemsAcross(self, value):
        self.printer.ItemsAcross = value

    @property
    def ItemSizeHeight(self):
        return self.printer.ItemSizeHeight

    @ItemSizeHeight.setter
    def ItemSizeHeight(self, value):
        self.printer.ItemSizeHeight = value

    @property
    def ItemSizeWidth(self):
        return self.printer.ItemSizeWidth

    @ItemSizeWidth.setter
    def ItemSizeWidth(self, value):
        self.printer.ItemSizeWidth = value

    @property
    def LeftMargin(self):
        return self.printer.LeftMargin

    @LeftMargin.setter
    def LeftMargin(self, value):
        self.printer.LeftMargin = value

    @property
    def Orientation(self):
        return self.printer.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.printer.Orientation = value

    @property
    def PaperBin(self):
        return AcPrintPaperBin(self.printer.PaperBin)

    @PaperBin.setter
    def PaperBin(self, value):
        self.printer.PaperBin = value

    @property
    def PaperSize(self):
        return AcPrintPaperSize(self.printer.PaperSize)

    @PaperSize.setter
    def PaperSize(self, value):
        self.printer.PaperSize = value

    @property
    def Port(self):
        return self.printer.Port

    @property
    def PrintQuality(self):
        return AcPrintObjQuality(self.printer.PrintQuality)

    @PrintQuality.setter
    def PrintQuality(self, value):
        self.printer.PrintQuality = value

    @property
    def RightMargin(self):
        return self.printer.RightMargin

    @RightMargin.setter
    def RightMargin(self, value):
        self.printer.RightMargin = value

    @property
    def RowSpacing(self):
        return self.printer.RowSpacing

    @RowSpacing.setter
    def RowSpacing(self, value):
        self.printer.RowSpacing = value

    @property
    def TopMargin(self):
        return self.printer.TopMargin

    @TopMargin.setter
    def TopMargin(self, value):
        self.printer.TopMargin = value

class Printers:

    def __init__(self, printers=None):
        self.printers = printers

    @property
    def Application(self):
        return self.printers.Application

    @property
    def Count(self):
        return self.printers.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.printers.Item):
            return self.printers.Item(*args, **arguments)
        else:
            return self.printers.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.printers.Parent

class Properties:

    def __init__(self, properties=None):
        self.properties = properties

    @property
    def Application(self):
        return self.properties.Application

    @property
    def Count(self):
        return self.properties.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.properties.Item):
            return self.properties.Item(*args, **arguments)
        else:
            return self.properties.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.properties.Parent

class Rectangle:

    def __init__(self, rectangle=None):
        self.rectangle = rectangle

    @property
    def Application(self):
        return self.rectangle.Application

    @property
    def BackColor(self):
        return self.rectangle.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.rectangle.BackColor = value

    @property
    def BackShade(self):
        return self.rectangle.BackShade

    @BackShade.setter
    def BackShade(self, value):
        self.rectangle.BackShade = value

    @property
    def BackStyle(self):
        return self.rectangle.BackStyle

    @BackStyle.setter
    def BackStyle(self, value):
        self.rectangle.BackStyle = value

    @property
    def BackThemeColorIndex(self):
        return self.rectangle.BackThemeColorIndex

    @BackThemeColorIndex.setter
    def BackThemeColorIndex(self, value):
        self.rectangle.BackThemeColorIndex = value

    @property
    def BackTint(self):
        return self.rectangle.BackTint

    @BackTint.setter
    def BackTint(self, value):
        self.rectangle.BackTint = value

    @property
    def BorderColor(self):
        return self.rectangle.BorderColor

    @BorderColor.setter
    def BorderColor(self, value):
        self.rectangle.BorderColor = value

    @property
    def BorderShade(self):
        return self.rectangle.BorderShade

    @BorderShade.setter
    def BorderShade(self, value):
        self.rectangle.BorderShade = value

    @property
    def BorderStyle(self):
        return self.rectangle.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.rectangle.BorderStyle = value

    @property
    def BorderThemeColorIndex(self):
        return self.rectangle.BorderThemeColorIndex

    @BorderThemeColorIndex.setter
    def BorderThemeColorIndex(self, value):
        self.rectangle.BorderThemeColorIndex = value

    @property
    def BorderTint(self):
        return self.rectangle.BorderTint

    @BorderTint.setter
    def BorderTint(self, value):
        self.rectangle.BorderTint = value

    @property
    def BorderWidth(self):
        return self.rectangle.BorderWidth

    @BorderWidth.setter
    def BorderWidth(self, value):
        self.rectangle.BorderWidth = value

    @property
    def ControlType(self):
        return self.rectangle.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.rectangle.ControlType = value

    @property
    def DisplayWhen(self):
        return self.rectangle.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.rectangle.DisplayWhen = value

    @property
    def EventProcPrefix(self):
        return self.rectangle.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.rectangle.EventProcPrefix = value

    @property
    def Height(self):
        return self.rectangle.Height

    @Height.setter
    def Height(self, value):
        self.rectangle.Height = value

    @property
    def HorizontalAnchor(self):
        return self.rectangle.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.rectangle.HorizontalAnchor = value

    @property
    def InSelection(self):
        return self.rectangle.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.rectangle.InSelection = value

    @property
    def IsVisible(self):
        return self.rectangle.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.rectangle.IsVisible = value

    @property
    def Left(self):
        return self.rectangle.Left

    @Left.setter
    def Left(self, value):
        self.rectangle.Left = value

    @property
    def Name(self):
        return self.rectangle.Name

    @Name.setter
    def Name(self, value):
        self.rectangle.Name = value

    @property
    def OldBorderStyle(self):
        return self.rectangle.OldBorderStyle

    @OldBorderStyle.setter
    def OldBorderStyle(self, value):
        self.rectangle.OldBorderStyle = value

    @property
    def OnClick(self):
        return self.rectangle.OnClick

    @OnClick.setter
    def OnClick(self, value):
        self.rectangle.OnClick = value

    @property
    def OnDblClick(self):
        return self.rectangle.OnDblClick

    @OnDblClick.setter
    def OnDblClick(self, value):
        self.rectangle.OnDblClick = value

    @property
    def OnMouseDown(self):
        return self.rectangle.OnMouseDown

    @OnMouseDown.setter
    def OnMouseDown(self, value):
        self.rectangle.OnMouseDown = value

    @property
    def OnMouseMove(self):
        return self.rectangle.OnMouseMove

    @OnMouseMove.setter
    def OnMouseMove(self, value):
        self.rectangle.OnMouseMove = value

    @property
    def OnMouseUp(self):
        return self.rectangle.OnMouseUp

    @OnMouseUp.setter
    def OnMouseUp(self, value):
        self.rectangle.OnMouseUp = value

    @property
    def Parent(self):
        return self.rectangle.Parent

    @property
    def Properties(self):
        return Properties(self.rectangle.Properties)

    @property
    def Section(self):
        return self.rectangle.Section

    @Section.setter
    def Section(self, value):
        self.rectangle.Section = value

    @property
    def SpecialEffect(self):
        return self.rectangle.SpecialEffect

    @SpecialEffect.setter
    def SpecialEffect(self, value):
        self.rectangle.SpecialEffect = value

    @property
    def Tag(self):
        return self.rectangle.Tag

    @Tag.setter
    def Tag(self, value):
        self.rectangle.Tag = value

    @property
    def Top(self):
        return self.rectangle.Top

    @Top.setter
    def Top(self, value):
        self.rectangle.Top = value

    @property
    def VerticalAnchor(self):
        return self.rectangle.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.rectangle.VerticalAnchor = value

    @property
    def Visible(self):
        return self.rectangle.Visible

    @Visible.setter
    def Visible(self, value):
        self.rectangle.Visible = value

    @property
    def Width(self):
        return self.rectangle.Width

    @Width.setter
    def Width(self, value):
        self.rectangle.Width = value

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.rectangle.Move(*args, **arguments)

    def SizeToFit(self):
        self.rectangle.SizeToFit()

class Reference:

    def __init__(self, reference=None):
        self.reference = reference

    @property
    def BuiltIn(self):
        return self.reference.BuiltIn

    @property
    def Collection(self):
        return self.reference.Collection

    @property
    def FullPath(self):
        return self.reference.FullPath

    @property
    def Guid(self):
        return self.reference.Guid

    @property
    def IsBroken(self):
        return self.reference.IsBroken

    @property
    def Kind(self):
        return self.reference.Kind

    @property
    def Major(self):
        return self.reference.Major

    @property
    def Minor(self):
        return self.reference.Minor

    @property
    def Name(self):
        return self.reference.Name

class References:

    def __init__(self, references=None):
        self.references = references

    @property
    def Count(self):
        return self.references.Count

    @property
    def Parent(self):
        return self.references.Parent

    def AddFromFile(self, *args, FileName=None):
        arguments = {"FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.references.AddFromFile(*args, **arguments)

    def AddFromGuid(self, *args, Guid=None, Major=None, Minor=None):
        arguments = {"Guid": Guid, "Major": Major, "Minor": Minor}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.references.AddFromGuid(*args, **arguments)

    def Item(self, *args, var=None):
        arguments = {"var": var}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.references.Item(*args, **arguments)

    def Remove(self, *args, Reference=None):
        arguments = {"Reference": Reference}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.references.Remove(*args, **arguments)

class Report:

    def __init__(self, report=None):
        self.report = report

    @property
    def ActiveControl(self):
        return self.report.ActiveControl

    @property
    def AllowLayoutView(self):
        return self.report.AllowLayoutView

    @AllowLayoutView.setter
    def AllowLayoutView(self, value):
        self.report.AllowLayoutView = value

    @property
    def AllowReportView(self):
        return self.report.AllowReportView

    @AllowReportView.setter
    def AllowReportView(self, value):
        self.report.AllowReportView = value

    @property
    def Application(self):
        return self.report.Application

    @property
    def AutoCenter(self):
        return self.report.AutoCenter

    @AutoCenter.setter
    def AutoCenter(self, value):
        self.report.AutoCenter = value

    @property
    def AutoResize(self):
        return self.report.AutoResize

    @AutoResize.setter
    def AutoResize(self, value):
        self.report.AutoResize = value

    @property
    def BorderStyle(self):
        return self.report.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.report.BorderStyle = value

    @property
    def Caption(self):
        return self.report.Caption

    @Caption.setter
    def Caption(self, value):
        self.report.Caption = value

    @property
    def CloseButton(self):
        return self.report.CloseButton

    @CloseButton.setter
    def CloseButton(self, value):
        self.report.CloseButton = value

    @property
    def ControlBox(self):
        return self.report.ControlBox

    @ControlBox.setter
    def ControlBox(self, value):
        self.report.ControlBox = value

    @property
    def Controls(self):
        return Controls(self.report.Controls)

    @property
    def Count(self):
        return self.report.Count

    @property
    def CurrentRecord(self):
        return self.report.CurrentRecord

    @CurrentRecord.setter
    def CurrentRecord(self, value):
        self.report.CurrentRecord = value

    @property
    def CurrentView(self):
        return self.report.CurrentView

    @CurrentView.setter
    def CurrentView(self, value):
        self.report.CurrentView = value

    @property
    def CurrentX(self):
        return self.report.CurrentX

    @CurrentX.setter
    def CurrentX(self, value):
        self.report.CurrentX = value

    @property
    def CurrentY(self):
        return self.report.CurrentY

    @CurrentY.setter
    def CurrentY(self, value):
        self.report.CurrentY = value

    @property
    def Cycle(self):
        return self.report.Cycle

    @Cycle.setter
    def Cycle(self, value):
        self.report.Cycle = value

    @property
    def DateGrouping(self):
        return self.report.DateGrouping

    @DateGrouping.setter
    def DateGrouping(self, value):
        self.report.DateGrouping = value

    def DefaultControl(self, *args, ControlType=None):
        arguments = {"ControlType": ControlType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.report.DefaultControl):
            return self.report.DefaultControl(*args, **arguments)
        else:
            return self.report.GetDefaultControl(*args, **arguments)

    @property
    def DefaultView(self):
        return self.report.DefaultView

    @DefaultView.setter
    def DefaultView(self, value):
        self.report.DefaultView = value

    @property
    def Dirty(self):
        return self.report.Dirty

    @Dirty.setter
    def Dirty(self, value):
        self.report.Dirty = value

    @property
    def DisplayOnSharePointSite(self):
        return self.report.DisplayOnSharePointSite

    @DisplayOnSharePointSite.setter
    def DisplayOnSharePointSite(self, value):
        self.report.DisplayOnSharePointSite = value

    @property
    def DrawMode(self):
        return self.report.DrawMode

    @DrawMode.setter
    def DrawMode(self, value):
        self.report.DrawMode = value

    @property
    def DrawStyle(self):
        return self.report.DrawStyle

    @DrawStyle.setter
    def DrawStyle(self, value):
        self.report.DrawStyle = value

    @property
    def DrawWidth(self):
        return self.report.DrawWidth

    @DrawWidth.setter
    def DrawWidth(self, value):
        self.report.DrawWidth = value

    @property
    def FastLaserPrinting(self):
        return self.report.FastLaserPrinting

    @FastLaserPrinting.setter
    def FastLaserPrinting(self, value):
        self.report.FastLaserPrinting = value

    @property
    def FillColor(self):
        return self.report.FillColor

    @FillColor.setter
    def FillColor(self, value):
        self.report.FillColor = value

    @property
    def FillStyle(self):
        return self.report.FillStyle

    @FillStyle.setter
    def FillStyle(self, value):
        self.report.FillStyle = value

    @property
    def Filter(self):
        return self.report.Filter

    @Filter.setter
    def Filter(self, value):
        self.report.Filter = value

    @property
    def FilterOn(self):
        return self.report.FilterOn

    @FilterOn.setter
    def FilterOn(self, value):
        self.report.FilterOn = value

    @property
    def FilterOnLoad(self):
        return self.report.FilterOnLoad

    @FilterOnLoad.setter
    def FilterOnLoad(self, value):
        self.report.FilterOnLoad = value

    @property
    def FitToPage(self):
        return self.report.FitToPage

    @FitToPage.setter
    def FitToPage(self, value):
        self.report.FitToPage = value

    @property
    def FontBold(self):
        return self.report.FontBold

    @property
    def FontItalic(self):
        return self.report.FontItalic

    @property
    def FontName(self):
        return self.report.FontName

    @property
    def FontSize(self):
        return self.report.FontSize

    @property
    def FontUnderline(self):
        return self.report.FontUnderline

    @property
    def ForeColor(self):
        return self.report.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.report.ForeColor = value

    @property
    def FormatCount(self):
        return self.report.FormatCount

    @FormatCount.setter
    def FormatCount(self, value):
        self.report.FormatCount = value

    @property
    def GridX(self):
        return self.report.GridX

    @GridX.setter
    def GridX(self, value):
        self.report.GridX = value

    @property
    def GridY(self):
        return self.report.GridY

    @GridY.setter
    def GridY(self, value):
        self.report.GridY = value

    def GroupLevel(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.report.GroupLevel):
            return self.report.GroupLevel(*args, **arguments)
        else:
            return self.report.GetGroupLevel(*args, **arguments)

    @property
    def GrpKeepTogether(self):
        return self.report.GrpKeepTogether

    @GrpKeepTogether.setter
    def GrpKeepTogether(self, value):
        self.report.GrpKeepTogether = value

    @property
    def HasData(self):
        return self.report.HasData

    @HasData.setter
    def HasData(self, value):
        self.report.HasData = value

    @property
    def HasModule(self):
        return self.report.HasModule

    @HasModule.setter
    def HasModule(self, value):
        self.report.HasModule = value

    @property
    def Height(self):
        return self.report.Height

    @Height.setter
    def Height(self, value):
        self.report.Height = value

    @property
    def HelpContextId(self):
        return self.report.HelpContextId

    @HelpContextId.setter
    def HelpContextId(self, value):
        self.report.HelpContextId = value

    @property
    def HelpFile(self):
        return self.report.HelpFile

    @HelpFile.setter
    def HelpFile(self, value):
        self.report.HelpFile = value

    @property
    def Hwnd(self):
        return self.report.Hwnd

    @Hwnd.setter
    def Hwnd(self, value):
        self.report.Hwnd = value

    @property
    def KeyPreview(self):
        return self.report.KeyPreview

    @KeyPreview.setter
    def KeyPreview(self, value):
        self.report.KeyPreview = value

    @property
    def LayoutForPrint(self):
        return self.report.LayoutForPrint

    @LayoutForPrint.setter
    def LayoutForPrint(self, value):
        self.report.LayoutForPrint = value

    @property
    def Left(self):
        return self.report.Left

    @Left.setter
    def Left(self, value):
        self.report.Left = value

    @property
    def MenuBar(self):
        return self.report.MenuBar

    @MenuBar.setter
    def MenuBar(self, value):
        self.report.MenuBar = value

    @property
    def MinMaxButtons(self):
        return self.report.MinMaxButtons

    @MinMaxButtons.setter
    def MinMaxButtons(self, value):
        self.report.MinMaxButtons = value

    @property
    def Modal(self):
        return self.report.Modal

    @Modal.setter
    def Modal(self, value):
        self.report.Modal = value

    @property
    def Module(self):
        return self.report.Module

    @property
    def MouseWheel(self):
        return self.report.MouseWheel

    @MouseWheel.setter
    def MouseWheel(self, value):
        self.report.MouseWheel = value

    @property
    def Moveable(self):
        return self.report.Moveable

    @Moveable.setter
    def Moveable(self, value):
        self.report.Moveable = value

    @property
    def MoveLayout(self):
        return self.report.MoveLayout

    @MoveLayout.setter
    def MoveLayout(self, value):
        self.report.MoveLayout = value

    @property
    def Name(self):
        return self.report.Name

    @Name.setter
    def Name(self, value):
        self.report.Name = value

    @property
    def NextRecord(self):
        return self.report.NextRecord

    @NextRecord.setter
    def NextRecord(self, value):
        self.report.NextRecord = value

    @property
    def OnActivate(self):
        return self.report.OnActivate

    @OnActivate.setter
    def OnActivate(self, value):
        self.report.OnActivate = value

    @property
    def OnApplyFilter(self):
        return self.report.OnApplyFilter

    @OnApplyFilter.setter
    def OnApplyFilter(self, value):
        self.report.OnApplyFilter = value

    @property
    def OnClick(self):
        return self.report.OnClick

    @OnClick.setter
    def OnClick(self, value):
        self.report.OnClick = value

    @property
    def OnClose(self):
        return self.report.OnClose

    @OnClose.setter
    def OnClose(self, value):
        self.report.OnClose = value

    @property
    def OnCurrent(self):
        return self.report.OnCurrent

    @OnCurrent.setter
    def OnCurrent(self, value):
        self.report.OnCurrent = value

    @property
    def OnDblClick(self):
        return self.report.OnDblClick

    @OnDblClick.setter
    def OnDblClick(self, value):
        self.report.OnDblClick = value

    @property
    def OnDeactivate(self):
        return self.report.OnDeactivate

    @OnDeactivate.setter
    def OnDeactivate(self, value):
        self.report.OnDeactivate = value

    @property
    def OnError(self):
        return self.report.OnError

    @OnError.setter
    def OnError(self, value):
        self.report.OnError = value

    @property
    def OnFilter(self):
        return self.report.OnFilter

    @OnFilter.setter
    def OnFilter(self, value):
        self.report.OnFilter = value

    @property
    def OnGotFocus(self):
        return self.report.OnGotFocus

    @OnGotFocus.setter
    def OnGotFocus(self, value):
        self.report.OnGotFocus = value

    @property
    def OnKeyDown(self):
        return self.report.OnKeyDown

    @OnKeyDown.setter
    def OnKeyDown(self, value):
        self.report.OnKeyDown = value

    @property
    def OnKeyPress(self):
        return self.report.OnKeyPress

    @OnKeyPress.setter
    def OnKeyPress(self, value):
        self.report.OnKeyPress = value

    @property
    def OnKeyUp(self):
        return self.report.OnKeyUp

    @OnKeyUp.setter
    def OnKeyUp(self, value):
        self.report.OnKeyUp = value

    @property
    def OnLoad(self):
        return self.report.OnLoad

    @OnLoad.setter
    def OnLoad(self, value):
        self.report.OnLoad = value

    @property
    def OnLostFocus(self):
        return self.report.OnLostFocus

    @OnLostFocus.setter
    def OnLostFocus(self, value):
        self.report.OnLostFocus = value

    @property
    def OnMouseDown(self):
        return self.report.OnMouseDown

    @OnMouseDown.setter
    def OnMouseDown(self, value):
        self.report.OnMouseDown = value

    @property
    def OnMouseMove(self):
        return self.report.OnMouseMove

    @OnMouseMove.setter
    def OnMouseMove(self, value):
        self.report.OnMouseMove = value

    @property
    def OnMouseUp(self):
        return self.report.OnMouseUp

    @OnMouseUp.setter
    def OnMouseUp(self, value):
        self.report.OnMouseUp = value

    @property
    def OnNoData(self):
        return self.report.OnNoData

    @OnNoData.setter
    def OnNoData(self, value):
        self.report.OnNoData = value

    @property
    def OnOpen(self):
        return self.report.OnOpen

    @OnOpen.setter
    def OnOpen(self, value):
        self.report.OnOpen = value

    @property
    def OnPage(self):
        return self.report.OnPage

    @OnPage.setter
    def OnPage(self, value):
        self.report.OnPage = value

    @property
    def OnResize(self):
        return self.report.OnResize

    @OnResize.setter
    def OnResize(self, value):
        self.report.OnResize = value

    @property
    def OnTimer(self):
        return self.report.OnTimer

    @OnTimer.setter
    def OnTimer(self, value):
        self.report.OnTimer = value

    @property
    def OnUnload(self):
        return self.report.OnUnload

    @OnUnload.setter
    def OnUnload(self, value):
        self.report.OnUnload = value

    @property
    def OpenArgs(self):
        return self.report.OpenArgs

    @OpenArgs.setter
    def OpenArgs(self, value):
        self.report.OpenArgs = value

    @property
    def OrderBy(self):
        return self.report.OrderBy

    @OrderBy.setter
    def OrderBy(self, value):
        self.report.OrderBy = value

    @property
    def OrderByOn(self):
        return self.report.OrderByOn

    @OrderByOn.setter
    def OrderByOn(self, value):
        self.report.OrderByOn = value

    @property
    def OrderByOnLoad(self):
        return self.report.OrderByOnLoad

    @OrderByOnLoad.setter
    def OrderByOnLoad(self, value):
        self.report.OrderByOnLoad = value

    @property
    def Orientation(self):
        return self.report.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.report.Orientation = value

    @property
    def Page(self):
        return self.report.Page

    @Page.setter
    def Page(self, value):
        self.report.Page = value

    @property
    def PageFooter(self):
        return self.report.PageFooter

    @PageFooter.setter
    def PageFooter(self, value):
        self.report.PageFooter = value

    @property
    def PageHeader(self):
        return self.report.PageHeader

    @PageHeader.setter
    def PageHeader(self, value):
        self.report.PageHeader = value

    @property
    def Pages(self):
        return self.report.Pages

    @Pages.setter
    def Pages(self, value):
        self.report.Pages = value

    @property
    def Painting(self):
        return self.report.Painting

    @Painting.setter
    def Painting(self, value):
        self.report.Painting = value

    @property
    def PaintPalette(self):
        return self.report.PaintPalette

    @PaintPalette.setter
    def PaintPalette(self, value):
        self.report.PaintPalette = value

    @property
    def PaletteSource(self):
        return self.report.PaletteSource

    @PaletteSource.setter
    def PaletteSource(self, value):
        self.report.PaletteSource = value

    @property
    def Parent(self):
        return self.report.Parent

    @property
    def Picture(self):
        return self.report.Picture

    @Picture.setter
    def Picture(self, value):
        self.report.Picture = value

    @property
    def PictureAlignment(self):
        return self.report.PictureAlignment

    @PictureAlignment.setter
    def PictureAlignment(self, value):
        self.report.PictureAlignment = value

    @property
    def PictureData(self):
        return self.report.PictureData

    @PictureData.setter
    def PictureData(self, value):
        self.report.PictureData = value

    @property
    def PicturePages(self):
        return self.report.PicturePages

    @PicturePages.setter
    def PicturePages(self, value):
        self.report.PicturePages = value

    @property
    def PicturePalette(self):
        return self.report.PicturePalette

    @PicturePalette.setter
    def PicturePalette(self, value):
        self.report.PicturePalette = value

    @property
    def PictureSizeMode(self):
        return self.report.PictureSizeMode

    @PictureSizeMode.setter
    def PictureSizeMode(self, value):
        self.report.PictureSizeMode = value

    @property
    def PictureTiling(self):
        return self.report.PictureTiling

    @PictureTiling.setter
    def PictureTiling(self, value):
        self.report.PictureTiling = value

    @property
    def PictureType(self):
        return self.report.PictureType

    @PictureType.setter
    def PictureType(self, value):
        self.report.PictureType = value

    @property
    def PopUp(self):
        return self.report.PopUp

    @PopUp.setter
    def PopUp(self, value):
        self.report.PopUp = value

    @property
    def PrintCount(self):
        return self.report.PrintCount

    @PrintCount.setter
    def PrintCount(self, value):
        self.report.PrintCount = value

    @property
    def Printer(self):
        return Printer(self.report.Printer)

    @Printer.setter
    def Printer(self, value):
        self.report.Printer = value

    @property
    def PrintSection(self):
        return self.report.PrintSection

    @PrintSection.setter
    def PrintSection(self, value):
        self.report.PrintSection = value

    @property
    def Properties(self):
        return Properties(self.report.Properties)

    @property
    def PrtDevMode(self):
        return self.report.PrtDevMode

    @PrtDevMode.setter
    def PrtDevMode(self, value):
        self.report.PrtDevMode = value

    @property
    def PrtDevNames(self):
        return self.report.PrtDevNames

    @PrtDevNames.setter
    def PrtDevNames(self, value):
        self.report.PrtDevNames = value

    @property
    def PrtMip(self):
        return self.report.PrtMip

    @property
    def RecordLocks(self):
        return self.report.RecordLocks

    @RecordLocks.setter
    def RecordLocks(self, value):
        self.report.RecordLocks = value

    @property
    def Recordset(self):
        return self.report.Recordset

    @Recordset.setter
    def Recordset(self, value):
        self.report.Recordset = value

    @property
    def RecordSource(self):
        return self.report.RecordSource

    @RecordSource.setter
    def RecordSource(self, value):
        self.report.RecordSource = value

    @property
    def RecordSourceQualifier(self):
        return self.report.RecordSourceQualifier

    @RecordSourceQualifier.setter
    def RecordSourceQualifier(self, value):
        self.report.RecordSourceQualifier = value

    @property
    def Report(self):
        return self.report.Report

    @property
    def RibbonName(self):
        return self.report.RibbonName

    @RibbonName.setter
    def RibbonName(self, value):
        self.report.RibbonName = value

    @property
    def ScaleHeight(self):
        return self.report.ScaleHeight

    @ScaleHeight.setter
    def ScaleHeight(self, value):
        self.report.ScaleHeight = value

    @property
    def ScaleLeft(self):
        return self.report.ScaleLeft

    @ScaleLeft.setter
    def ScaleLeft(self, value):
        self.report.ScaleLeft = value

    @property
    def ScaleMode(self):
        return self.report.ScaleMode

    @ScaleMode.setter
    def ScaleMode(self, value):
        self.report.ScaleMode = value

    @property
    def ScaleTop(self):
        return self.report.ScaleTop

    @ScaleTop.setter
    def ScaleTop(self, value):
        self.report.ScaleTop = value

    @property
    def ScaleWidth(self):
        return self.report.ScaleWidth

    @ScaleWidth.setter
    def ScaleWidth(self, value):
        self.report.ScaleWidth = value

    @property
    def ScrollBars(self):
        return self.report.ScrollBars

    @ScrollBars.setter
    def ScrollBars(self, value):
        self.report.ScrollBars = value

    def Section(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.report.Section):
            return self.report.Section(*args, **arguments)
        else:
            return self.report.GetSection(*args, **arguments)

    @property
    def ServerFilter(self):
        return self.report.ServerFilter

    @ServerFilter.setter
    def ServerFilter(self, value):
        self.report.ServerFilter = value

    @property
    def ShortcutMenuBar(self):
        return self.report.ShortcutMenuBar

    @ShortcutMenuBar.setter
    def ShortcutMenuBar(self, value):
        self.report.ShortcutMenuBar = value

    @property
    def ShowPageMargins(self):
        return self.report.ShowPageMargins

    @ShowPageMargins.setter
    def ShowPageMargins(self, value):
        self.report.ShowPageMargins = value

    @property
    def Tag(self):
        return self.report.Tag

    @Tag.setter
    def Tag(self, value):
        self.report.Tag = value

    @property
    def TimerInterval(self):
        return self.report.TimerInterval

    @TimerInterval.setter
    def TimerInterval(self, value):
        self.report.TimerInterval = value

    @property
    def Toolbar(self):
        return self.report.Toolbar

    @Toolbar.setter
    def Toolbar(self, value):
        self.report.Toolbar = value

    @property
    def Top(self):
        return self.report.Top

    @Top.setter
    def Top(self, value):
        self.report.Top = value

    @property
    def UseDefaultPrinter(self):
        return self.report.UseDefaultPrinter

    @UseDefaultPrinter.setter
    def UseDefaultPrinter(self, value):
        self.report.UseDefaultPrinter = value

    @property
    def Visible(self):
        return self.report.Visible

    @Visible.setter
    def Visible(self, value):
        self.report.Visible = value

    @property
    def Width(self):
        return self.report.Width

    @Width.setter
    def Width(self, value):
        self.report.Width = value

    @property
    def WindowHeight(self):
        return self.report.WindowHeight

    @property
    def WindowLeft(self):
        return self.report.WindowLeft

    @property
    def WindowTop(self):
        return self.report.WindowTop

    @property
    def WindowWidth(self):
        return self.report.WindowWidth

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.report.Move(*args, **arguments)

    def Print(self, *args, Expr=None):
        arguments = {"Expr": Expr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.report.Print(*args, **arguments)

    def PSet(self, *args, Flags=None, x=None, y=None, Color=None):
        arguments = {"Flags": Flags, "x": x, "y": y, "Color": Color}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.report.PSet(*args, **arguments)

    def Requery(self):
        self.report.Requery()

    def Scale(self, *args, Flags=None, x1=None, y1=None, x2=None, y2=None):
        arguments = {"Flags": Flags, "x1": x1, "y1": y1, "x2": x2, "y2": y2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.report.Scale(*args, **arguments)

    def TextHeight(self, *args, Expr=None):
        arguments = {"Expr": Expr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.report.TextHeight(*args, **arguments)

    def TextWidth(self, *args, Expr=None):
        arguments = {"Expr": Expr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.report.TextWidth(*args, **arguments)

class Reports:

    def __init__(self, reports=None):
        self.reports = reports

    @property
    def Application(self):
        return self.reports.Application

    @property
    def Count(self):
        return self.reports.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.reports.Item):
            return self.reports.Item(*args, **arguments)
        else:
            return self.reports.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.reports.Parent

class ReturnVar:

    def __init__(self, returnvar=None):
        self.returnvar = returnvar

    @property
    def Name(self):
        return self.returnvar.Name

    @property
    def Value(self):
        return self.returnvar.Value

class ReturnVars:

    def __init__(self, returnvars=None):
        self.returnvars = returnvars

    @property
    def Application(self):
        return self.returnvars.Application

    @property
    def Count(self):
        return self.returnvars.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.returnvars.Item):
            return self.returnvars.Item(*args, **arguments)
        else:
            return self.returnvars.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.returnvars.Parent

class Screen:

    def __init__(self, screen=None):
        self.screen = screen

    @property
    def ActiveControl(self):
        return self.screen.ActiveControl

    @property
    def ActiveDatasheet(self):
        return self.screen.ActiveDatasheet

    @property
    def ActiveForm(self):
        return self.screen.ActiveForm

    @property
    def ActiveReport(self):
        return self.screen.ActiveReport

    @property
    def Application(self):
        return self.screen.Application

    @property
    def MousePointer(self):
        return self.screen.MousePointer

    @MousePointer.setter
    def MousePointer(self, value):
        self.screen.MousePointer = value

    @property
    def Parent(self):
        return self.screen.Parent

    @property
    def PreviousControl(self):
        return self.screen.PreviousControl

class Section:

    def __init__(self, section=None):
        self.section = section

    @property
    def AlternateBackColor(self):
        return self.section.AlternateBackColor

    @AlternateBackColor.setter
    def AlternateBackColor(self, value):
        self.section.AlternateBackColor = value

    @property
    def AlternateBackShade(self):
        return self.section.AlternateBackShade

    @AlternateBackShade.setter
    def AlternateBackShade(self, value):
        self.section.AlternateBackShade = value

    @property
    def AlternateBackThemeColorIndex(self):
        return self.section.AlternateBackThemeColorIndex

    @AlternateBackThemeColorIndex.setter
    def AlternateBackThemeColorIndex(self, value):
        self.section.AlternateBackThemeColorIndex = value

    @property
    def AlternateBackTint(self):
        return self.section.AlternateBackTint

    @AlternateBackTint.setter
    def AlternateBackTint(self, value):
        self.section.AlternateBackTint = value

    @property
    def Application(self):
        return self.section.Application

    @property
    def AutoHeight(self):
        return self.section.AutoHeight

    @AutoHeight.setter
    def AutoHeight(self, value):
        self.section.AutoHeight = value

    @property
    def BackColor(self):
        return self.section.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.section.BackColor = value

    @property
    def BackShade(self):
        return self.section.BackShade

    @BackShade.setter
    def BackShade(self, value):
        self.section.BackShade = value

    @property
    def BackThemeColorIndex(self):
        return self.section.BackThemeColorIndex

    @BackThemeColorIndex.setter
    def BackThemeColorIndex(self, value):
        self.section.BackThemeColorIndex = value

    @property
    def BackTint(self):
        return self.section.BackTint

    @BackTint.setter
    def BackTint(self, value):
        self.section.BackTint = value

    @property
    def CanGrow(self):
        return self.section.CanGrow

    @CanGrow.setter
    def CanGrow(self, value):
        self.section.CanGrow = value

    @property
    def CanShrink(self):
        return self.section.CanShrink

    @CanShrink.setter
    def CanShrink(self, value):
        self.section.CanShrink = value

    @property
    def Controls(self):
        return Controls(self.section.Controls)

    @property
    def DisplayWhen(self):
        return self.section.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.section.DisplayWhen = value

    @property
    def EventProcPrefix(self):
        return self.section.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.section.EventProcPrefix = value

    @property
    def ForceNewPage(self):
        return self.section.ForceNewPage

    @ForceNewPage.setter
    def ForceNewPage(self, value):
        self.section.ForceNewPage = value

    @property
    def HasContinued(self):
        return self.section.HasContinued

    @HasContinued.setter
    def HasContinued(self, value):
        self.section.HasContinued = value

    @property
    def Height(self):
        return self.section.Height

    @Height.setter
    def Height(self, value):
        self.section.Height = value

    @property
    def InSelection(self):
        return self.section.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.section.InSelection = value

    @property
    def KeepTogether(self):
        return self.section.KeepTogether

    @KeepTogether.setter
    def KeepTogether(self, value):
        self.section.KeepTogether = value

    @property
    def Name(self):
        return self.section.Name

    @Name.setter
    def Name(self, value):
        self.section.Name = value

    @property
    def NewRowOrCol(self):
        return self.section.NewRowOrCol

    @NewRowOrCol.setter
    def NewRowOrCol(self, value):
        self.section.NewRowOrCol = value

    @property
    def OnClick(self):
        return self.section.OnClick

    @OnClick.setter
    def OnClick(self, value):
        self.section.OnClick = value

    @property
    def OnDblClick(self):
        return self.section.OnDblClick

    @OnDblClick.setter
    def OnDblClick(self, value):
        self.section.OnDblClick = value

    @property
    def OnFormat(self):
        return self.section.OnFormat

    @OnFormat.setter
    def OnFormat(self, value):
        self.section.OnFormat = value

    @property
    def OnMouseDown(self):
        return self.section.OnMouseDown

    @OnMouseDown.setter
    def OnMouseDown(self, value):
        self.section.OnMouseDown = value

    @property
    def OnMouseMove(self):
        return self.section.OnMouseMove

    @OnMouseMove.setter
    def OnMouseMove(self, value):
        self.section.OnMouseMove = value

    @property
    def OnMouseUp(self):
        return self.section.OnMouseUp

    @OnMouseUp.setter
    def OnMouseUp(self, value):
        self.section.OnMouseUp = value

    @property
    def OnPaint(self):
        return self.section.OnPaint

    @OnPaint.setter
    def OnPaint(self, value):
        self.section.OnPaint = value

    @property
    def OnPrint(self):
        return self.section.OnPrint

    @OnPrint.setter
    def OnPrint(self, value):
        self.section.OnPrint = value

    @property
    def OnRetreat(self):
        return self.section.OnRetreat

    @OnRetreat.setter
    def OnRetreat(self, value):
        self.section.OnRetreat = value

    @property
    def Parent(self):
        return self.section.Parent

    @property
    def Properties(self):
        return Properties(self.section.Properties)

    @property
    def RepeatSection(self):
        return self.section.RepeatSection

    @RepeatSection.setter
    def RepeatSection(self, value):
        self.section.RepeatSection = value

    @property
    def SpecialEffect(self):
        return self.section.SpecialEffect

    @SpecialEffect.setter
    def SpecialEffect(self, value):
        self.section.SpecialEffect = value

    @property
    def Tag(self):
        return self.section.Tag

    @Tag.setter
    def Tag(self, value):
        self.section.Tag = value

    @property
    def Visible(self):
        return self.section.Visible

    @Visible.setter
    def Visible(self, value):
        self.section.Visible = value

    @property
    def WillContinue(self):
        return self.section.WillContinue

    @WillContinue.setter
    def WillContinue(self, value):
        self.section.WillContinue = value

    def SetTabOrder(self):
        self.section.SetTabOrder()

class SharedResource:

    def __init__(self, sharedresource=None):
        self.sharedresource = sharedresource

    @property
    def Name(self):
        return self.sharedresource.Name

    @Name.setter
    def Name(self, value):
        self.sharedresource.Name = value

    @property
    def Parent(self):
        return self.sharedresource.Parent

    @property
    def Type(self):
        return self.sharedresource.Type

    def Delete(self):
        self.sharedresource.Delete()

class SharedResources:

    def __init__(self, sharedresources=None):
        self.sharedresources = sharedresources

    @property
    def Application(self):
        return self.sharedresources.Application

    @property
    def Count(self):
        return self.sharedresources.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.sharedresources.Item):
            return self.sharedresources.Item(*args, **arguments)
        else:
            return self.sharedresources.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.sharedresources.Parent

class SmartTag:

    def __init__(self, smarttag=None):
        self.smarttag = smarttag

    @property
    def Application(self):
        return self.smarttag.Application

    @property
    def IsMissing(self):
        return self.smarttag.IsMissing

    @property
    def Name(self):
        return self.smarttag.Name

    @property
    def Parent(self):
        return self.smarttag.Parent

    @property
    def Properties(self):
        return SmartTagProperties(self.smarttag.Properties)

    @property
    def SmartTagActions(self):
        return SmartTagActions(self.smarttag.SmartTagActions)

    @property
    def XML(self):
        return self.smarttag.XML

    def Delete(self):
        return self.smarttag.Delete()

class SmartTagAction:

    def __init__(self, smarttagaction=None):
        self.smarttagaction = smarttagaction

    @property
    def Application(self):
        return self.smarttagaction.Application

    @property
    def Name(self):
        return self.smarttagaction.Name

    @property
    def Parent(self):
        return self.smarttagaction.Parent

    def Execute(self):
        self.smarttagaction.Execute()

class SmartTagActions:

    def __init__(self, smarttagactions=None):
        self.smarttagactions = smarttagactions

    @property
    def Application(self):
        return self.smarttagactions.Application

    @property
    def Count(self):
        return self.smarttagactions.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.smarttagactions.Item):
            return self.smarttagactions.Item(*args, **arguments)
        else:
            return self.smarttagactions.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.smarttagactions.Parent

class SmartTagProperties:

    def __init__(self, smarttagproperties=None):
        self.smarttagproperties = smarttagproperties

    def __call__(self, item):
        return SmartTagPropertie(self.smarttagproperties(item))

    @property
    def Application(self):
        return self.smarttagproperties.Application

    @property
    def Count(self):
        return self.smarttagproperties.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.smarttagproperties.Item):
            return self.smarttagproperties.Item(*args, **arguments)
        else:
            return self.smarttagproperties.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.smarttagproperties.Parent

    def Add(self, *args, Name=None, Value=None):
        arguments = {"Name": Name, "Value": Value}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return SmartTagPropertie(self.smarttagproperties.Add(*args, **arguments))

class SmartTagProperty:

    def __init__(self, smarttagproperty=None):
        self.smarttagproperty = smarttagproperty

    @property
    def Name(self):
        return self.smarttagproperty.Name

    @Name.setter
    def Name(self, value):
        self.smarttagproperty.Name = value

    @property
    def Value(self):
        return self.smarttagproperty.Value

    @Value.setter
    def Value(self, value):
        self.smarttagproperty.Value = value

    def Delete(self):
        return self.smarttagproperty.Delete()

class SmartTags:

    def __init__(self, smarttags=None):
        self.smarttags = smarttags

    @property
    def Application(self):
        return self.smarttags.Application

    @property
    def Count(self):
        return self.smarttags.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.smarttags.Item):
            return self.smarttags.Item(*args, **arguments)
        else:
            return self.smarttags.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.smarttags.Parent

    def Add(self, *args, Name=None):
        arguments = {"Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.smarttags.Add(*args, **arguments)

class SubForm:

    def __init__(self, subform=None):
        self.subform = subform

    @property
    def AddColon(self):
        return self.subform.AddColon

    @AddColon.setter
    def AddColon(self, value):
        self.subform.AddColon = value

    @property
    def Application(self):
        return self.subform.Application

    @property
    def AutoLabel(self):
        return self.subform.AutoLabel

    @AutoLabel.setter
    def AutoLabel(self, value):
        self.subform.AutoLabel = value

    @property
    def BorderColor(self):
        return self.subform.BorderColor

    @BorderColor.setter
    def BorderColor(self, value):
        self.subform.BorderColor = value

    @property
    def BorderShade(self):
        return self.subform.BorderShade

    @BorderShade.setter
    def BorderShade(self, value):
        self.subform.BorderShade = value

    @property
    def BorderStyle(self):
        return self.subform.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.subform.BorderStyle = value

    @property
    def BorderThemeColorIndex(self):
        return self.subform.BorderThemeColorIndex

    @BorderThemeColorIndex.setter
    def BorderThemeColorIndex(self, value):
        self.subform.BorderThemeColorIndex = value

    @property
    def BorderTint(self):
        return self.subform.BorderTint

    @BorderTint.setter
    def BorderTint(self, value):
        self.subform.BorderTint = value

    @property
    def BorderWidth(self):
        return self.subform.BorderWidth

    @BorderWidth.setter
    def BorderWidth(self, value):
        self.subform.BorderWidth = value

    @property
    def BottomPadding(self):
        return self.subform.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.subform.BottomPadding = value

    @property
    def CanGrow(self):
        return self.subform.CanGrow

    @CanGrow.setter
    def CanGrow(self, value):
        self.subform.CanGrow = value

    @property
    def CanShrink(self):
        return self.subform.CanShrink

    @CanShrink.setter
    def CanShrink(self, value):
        self.subform.CanShrink = value

    @property
    def Controls(self):
        return Controls(self.subform.Controls)

    @property
    def ControlType(self):
        return self.subform.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.subform.ControlType = value

    @property
    def DisplayWhen(self):
        return self.subform.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.subform.DisplayWhen = value

    @property
    def Enabled(self):
        return self.subform.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.subform.Enabled = value

    @property
    def EventProcPrefix(self):
        return self.subform.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.subform.EventProcPrefix = value

    @property
    def FilterOnEmptyMaster(self):
        return self.subform.FilterOnEmptyMaster

    @FilterOnEmptyMaster.setter
    def FilterOnEmptyMaster(self, value):
        self.subform.FilterOnEmptyMaster = value

    @property
    def Form(self):
        return self.subform.Form

    @property
    def GridlineColor(self):
        return self.subform.GridlineColor

    @GridlineColor.setter
    def GridlineColor(self, value):
        self.subform.GridlineColor = value

    @property
    def GridlineShade(self):
        return self.subform.GridlineShade

    @GridlineShade.setter
    def GridlineShade(self, value):
        self.subform.GridlineShade = value

    @property
    def GridlineStyleBottom(self):
        return self.subform.GridlineStyleBottom

    @GridlineStyleBottom.setter
    def GridlineStyleBottom(self, value):
        self.subform.GridlineStyleBottom = value

    @property
    def GridlineStyleLeft(self):
        return self.subform.GridlineStyleLeft

    @GridlineStyleLeft.setter
    def GridlineStyleLeft(self, value):
        self.subform.GridlineStyleLeft = value

    @property
    def GridlineStyleRight(self):
        return self.subform.GridlineStyleRight

    @GridlineStyleRight.setter
    def GridlineStyleRight(self, value):
        self.subform.GridlineStyleRight = value

    @property
    def GridlineStyleTop(self):
        return self.subform.GridlineStyleTop

    @GridlineStyleTop.setter
    def GridlineStyleTop(self, value):
        self.subform.GridlineStyleTop = value

    @property
    def GridlineThemeColorIndex(self):
        return self.subform.GridlineThemeColorIndex

    @GridlineThemeColorIndex.setter
    def GridlineThemeColorIndex(self, value):
        self.subform.GridlineThemeColorIndex = value

    @property
    def GridlineTint(self):
        return self.subform.GridlineTint

    @GridlineTint.setter
    def GridlineTint(self, value):
        self.subform.GridlineTint = value

    @property
    def GridlineWidthBottom(self):
        return self.subform.GridlineWidthBottom

    @GridlineWidthBottom.setter
    def GridlineWidthBottom(self, value):
        self.subform.GridlineWidthBottom = value

    @property
    def GridlineWidthLeft(self):
        return self.subform.GridlineWidthLeft

    @GridlineWidthLeft.setter
    def GridlineWidthLeft(self, value):
        self.subform.GridlineWidthLeft = value

    @property
    def GridlineWidthRight(self):
        return self.subform.GridlineWidthRight

    @GridlineWidthRight.setter
    def GridlineWidthRight(self, value):
        self.subform.GridlineWidthRight = value

    @property
    def GridlineWidthTop(self):
        return self.subform.GridlineWidthTop

    @GridlineWidthTop.setter
    def GridlineWidthTop(self, value):
        self.subform.GridlineWidthTop = value

    @property
    def Height(self):
        return self.subform.Height

    @Height.setter
    def Height(self, value):
        self.subform.Height = value

    @property
    def HorizontalAnchor(self):
        return self.subform.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.subform.HorizontalAnchor = value

    @property
    def InSelection(self):
        return self.subform.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.subform.InSelection = value

    @property
    def IsVisible(self):
        return self.subform.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.subform.IsVisible = value

    @property
    def LabelAlign(self):
        return self.subform.LabelAlign

    @LabelAlign.setter
    def LabelAlign(self, value):
        self.subform.LabelAlign = value

    @property
    def LabelX(self):
        return self.subform.LabelX

    @LabelX.setter
    def LabelX(self, value):
        self.subform.LabelX = value

    @property
    def LabelY(self):
        return self.subform.LabelY

    @LabelY.setter
    def LabelY(self, value):
        self.subform.LabelY = value

    @property
    def Layout(self):
        return AcLayoutType(self.subform.Layout)

    @property
    def LayoutID(self):
        return self.subform.LayoutID

    @property
    def Left(self):
        return self.subform.Left

    @Left.setter
    def Left(self, value):
        self.subform.Left = value

    @property
    def LeftPadding(self):
        return self.subform.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.subform.LeftPadding = value

    @property
    def LinkChildFields(self):
        return self.subform.LinkChildFields

    @LinkChildFields.setter
    def LinkChildFields(self, value):
        self.subform.LinkChildFields = value

    @property
    def LinkMasterFields(self):
        return self.subform.LinkMasterFields

    @LinkMasterFields.setter
    def LinkMasterFields(self, value):
        self.subform.LinkMasterFields = value

    @property
    def Locked(self):
        return self.subform.Locked

    @Locked.setter
    def Locked(self, value):
        self.subform.Locked = value

    @property
    def Name(self):
        return self.subform.Name

    @Name.setter
    def Name(self, value):
        self.subform.Name = value

    @property
    def OldBorderStyle(self):
        return self.subform.OldBorderStyle

    @OldBorderStyle.setter
    def OldBorderStyle(self, value):
        self.subform.OldBorderStyle = value

    @property
    def OnEnter(self):
        return self.subform.OnEnter

    @OnEnter.setter
    def OnEnter(self, value):
        self.subform.OnEnter = value

    @property
    def OnExit(self):
        return self.subform.OnExit

    @OnExit.setter
    def OnExit(self, value):
        self.subform.OnExit = value

    @property
    def Parent(self):
        return self.subform.Parent

    @property
    def Properties(self):
        return Properties(self.subform.Properties)

    @property
    def Report(self):
        return self.subform.Report

    @property
    def RightPadding(self):
        return self.subform.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.subform.RightPadding = value

    @property
    def Section(self):
        return self.subform.Section

    @Section.setter
    def Section(self, value):
        self.subform.Section = value

    @property
    def SourceObject(self):
        return self.subform.SourceObject

    @SourceObject.setter
    def SourceObject(self, value):
        self.subform.SourceObject = value

    @property
    def SpecialEffect(self):
        return self.subform.SpecialEffect

    @SpecialEffect.setter
    def SpecialEffect(self, value):
        self.subform.SpecialEffect = value

    @property
    def StatusBarText(self):
        return self.subform.StatusBarText

    @StatusBarText.setter
    def StatusBarText(self, value):
        self.subform.StatusBarText = value

    @property
    def TabIndex(self):
        return self.subform.TabIndex

    @TabIndex.setter
    def TabIndex(self, value):
        self.subform.TabIndex = value

    @property
    def TabStop(self):
        return self.subform.TabStop

    @TabStop.setter
    def TabStop(self, value):
        self.subform.TabStop = value

    @property
    def Tag(self):
        return self.subform.Tag

    @Tag.setter
    def Tag(self, value):
        self.subform.Tag = value

    @property
    def Top(self):
        return self.subform.Top

    @Top.setter
    def Top(self, value):
        self.subform.Top = value

    @property
    def TopPadding(self):
        return self.subform.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.subform.TopPadding = value

    @property
    def VerticalAnchor(self):
        return self.subform.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.subform.VerticalAnchor = value

    @property
    def Visible(self):
        return self.subform.Visible

    @Visible.setter
    def Visible(self, value):
        self.subform.Visible = value

    @property
    def Width(self):
        return self.subform.Width

    @Width.setter
    def Width(self, value):
        self.subform.Width = value

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.subform.Move(*args, **arguments)

    def Requery(self):
        self.subform.Requery()

    def SetFocus(self):
        return self.subform.SetFocus()

    def SizeToFit(self):
        self.subform.SizeToFit()

class TabControl:

    def __init__(self, tabcontrol=None):
        self.tabcontrol = tabcontrol

    @property
    def Application(self):
        return self.tabcontrol.Application

    @property
    def BackColor(self):
        return self.tabcontrol.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.tabcontrol.BackColor = value

    @property
    def BackShade(self):
        return self.tabcontrol.BackShade

    @BackShade.setter
    def BackShade(self, value):
        self.tabcontrol.BackShade = value

    @property
    def BackStyle(self):
        return self.tabcontrol.BackStyle

    @BackStyle.setter
    def BackStyle(self, value):
        self.tabcontrol.BackStyle = value

    @property
    def BackThemeColorIndex(self):
        return self.tabcontrol.BackThemeColorIndex

    @BackThemeColorIndex.setter
    def BackThemeColorIndex(self, value):
        self.tabcontrol.BackThemeColorIndex = value

    @property
    def BackTint(self):
        return self.tabcontrol.BackTint

    @BackTint.setter
    def BackTint(self, value):
        self.tabcontrol.BackTint = value

    @property
    def BorderColor(self):
        return self.tabcontrol.BorderColor

    @BorderColor.setter
    def BorderColor(self, value):
        self.tabcontrol.BorderColor = value

    @property
    def BorderShade(self):
        return self.tabcontrol.BorderShade

    @BorderShade.setter
    def BorderShade(self, value):
        self.tabcontrol.BorderShade = value

    @property
    def BorderStyle(self):
        return self.tabcontrol.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.tabcontrol.BorderStyle = value

    @property
    def BorderThemeColorIndex(self):
        return self.tabcontrol.BorderThemeColorIndex

    @BorderThemeColorIndex.setter
    def BorderThemeColorIndex(self, value):
        self.tabcontrol.BorderThemeColorIndex = value

    @property
    def BorderTint(self):
        return self.tabcontrol.BorderTint

    @BorderTint.setter
    def BorderTint(self, value):
        self.tabcontrol.BorderTint = value

    @property
    def BottomPadding(self):
        return self.tabcontrol.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.tabcontrol.BottomPadding = value

    @property
    def ControlType(self):
        return self.tabcontrol.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.tabcontrol.ControlType = value

    @property
    def DisplayWhen(self):
        return self.tabcontrol.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.tabcontrol.DisplayWhen = value

    @property
    def Enabled(self):
        return self.tabcontrol.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.tabcontrol.Enabled = value

    @property
    def EventProcPrefix(self):
        return self.tabcontrol.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.tabcontrol.EventProcPrefix = value

    @property
    def FontBold(self):
        return self.tabcontrol.FontBold

    @property
    def FontItalic(self):
        return self.tabcontrol.FontItalic

    @property
    def FontName(self):
        return self.tabcontrol.FontName

    @property
    def FontSize(self):
        return self.tabcontrol.FontSize

    @property
    def FontUnderline(self):
        return self.tabcontrol.FontUnderline

    @property
    def FontWeight(self):
        return self.tabcontrol.FontWeight

    @FontWeight.setter
    def FontWeight(self, value):
        self.tabcontrol.FontWeight = value

    @property
    def ForeColor(self):
        return self.tabcontrol.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.tabcontrol.ForeColor = value

    @property
    def ForeShade(self):
        return self.tabcontrol.ForeShade

    @ForeShade.setter
    def ForeShade(self, value):
        self.tabcontrol.ForeShade = value

    @property
    def ForeThemeColorIndex(self):
        return self.tabcontrol.ForeThemeColorIndex

    @ForeThemeColorIndex.setter
    def ForeThemeColorIndex(self, value):
        self.tabcontrol.ForeThemeColorIndex = value

    @property
    def ForeTint(self):
        return self.tabcontrol.ForeTint

    @ForeTint.setter
    def ForeTint(self, value):
        self.tabcontrol.ForeTint = value

    @property
    def Gradient(self):
        return self.tabcontrol.Gradient

    @Gradient.setter
    def Gradient(self, value):
        self.tabcontrol.Gradient = value

    @property
    def GridlineColor(self):
        return self.tabcontrol.GridlineColor

    @GridlineColor.setter
    def GridlineColor(self, value):
        self.tabcontrol.GridlineColor = value

    @property
    def GridlineShade(self):
        return self.tabcontrol.GridlineShade

    @GridlineShade.setter
    def GridlineShade(self, value):
        self.tabcontrol.GridlineShade = value

    @property
    def GridlineStyleBottom(self):
        return self.tabcontrol.GridlineStyleBottom

    @GridlineStyleBottom.setter
    def GridlineStyleBottom(self, value):
        self.tabcontrol.GridlineStyleBottom = value

    @property
    def GridlineStyleLeft(self):
        return self.tabcontrol.GridlineStyleLeft

    @GridlineStyleLeft.setter
    def GridlineStyleLeft(self, value):
        self.tabcontrol.GridlineStyleLeft = value

    @property
    def GridlineStyleRight(self):
        return self.tabcontrol.GridlineStyleRight

    @GridlineStyleRight.setter
    def GridlineStyleRight(self, value):
        self.tabcontrol.GridlineStyleRight = value

    @property
    def GridlineStyleTop(self):
        return self.tabcontrol.GridlineStyleTop

    @GridlineStyleTop.setter
    def GridlineStyleTop(self, value):
        self.tabcontrol.GridlineStyleTop = value

    @property
    def GridlineThemeColorIndex(self):
        return self.tabcontrol.GridlineThemeColorIndex

    @GridlineThemeColorIndex.setter
    def GridlineThemeColorIndex(self, value):
        self.tabcontrol.GridlineThemeColorIndex = value

    @property
    def GridlineTint(self):
        return self.tabcontrol.GridlineTint

    @GridlineTint.setter
    def GridlineTint(self, value):
        self.tabcontrol.GridlineTint = value

    @property
    def GridlineWidthBottom(self):
        return self.tabcontrol.GridlineWidthBottom

    @GridlineWidthBottom.setter
    def GridlineWidthBottom(self, value):
        self.tabcontrol.GridlineWidthBottom = value

    @property
    def GridlineWidthLeft(self):
        return self.tabcontrol.GridlineWidthLeft

    @GridlineWidthLeft.setter
    def GridlineWidthLeft(self, value):
        self.tabcontrol.GridlineWidthLeft = value

    @property
    def GridlineWidthRight(self):
        return self.tabcontrol.GridlineWidthRight

    @GridlineWidthRight.setter
    def GridlineWidthRight(self, value):
        self.tabcontrol.GridlineWidthRight = value

    @property
    def GridlineWidthTop(self):
        return self.tabcontrol.GridlineWidthTop

    @GridlineWidthTop.setter
    def GridlineWidthTop(self, value):
        self.tabcontrol.GridlineWidthTop = value

    @property
    def Height(self):
        return self.tabcontrol.Height

    @Height.setter
    def Height(self, value):
        self.tabcontrol.Height = value

    @property
    def HelpContextId(self):
        return self.tabcontrol.HelpContextId

    @HelpContextId.setter
    def HelpContextId(self, value):
        self.tabcontrol.HelpContextId = value

    @property
    def HorizontalAnchor(self):
        return self.tabcontrol.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.tabcontrol.HorizontalAnchor = value

    @property
    def HoverColor(self):
        return self.tabcontrol.HoverColor

    @HoverColor.setter
    def HoverColor(self, value):
        self.tabcontrol.HoverColor = value

    @property
    def HoverForeColor(self):
        return self.tabcontrol.HoverForeColor

    @HoverForeColor.setter
    def HoverForeColor(self, value):
        self.tabcontrol.HoverForeColor = value

    @property
    def HoverForeShade(self):
        return self.tabcontrol.HoverForeShade

    @HoverForeShade.setter
    def HoverForeShade(self, value):
        self.tabcontrol.HoverForeShade = value

    @property
    def HoverForeThemeColorIndex(self):
        return self.tabcontrol.HoverForeThemeColorIndex

    @HoverForeThemeColorIndex.setter
    def HoverForeThemeColorIndex(self, value):
        self.tabcontrol.HoverForeThemeColorIndex = value

    @property
    def HoverForeTint(self):
        return self.tabcontrol.HoverForeTint

    @HoverForeTint.setter
    def HoverForeTint(self, value):
        self.tabcontrol.HoverForeTint = value

    @property
    def HoverShade(self):
        return self.tabcontrol.HoverShade

    @HoverShade.setter
    def HoverShade(self, value):
        self.tabcontrol.HoverShade = value

    @property
    def HoverThemeColorIndex(self):
        return self.tabcontrol.HoverThemeColorIndex

    @HoverThemeColorIndex.setter
    def HoverThemeColorIndex(self, value):
        self.tabcontrol.HoverThemeColorIndex = value

    @property
    def HoverTint(self):
        return self.tabcontrol.HoverTint

    @HoverTint.setter
    def HoverTint(self, value):
        self.tabcontrol.HoverTint = value

    @property
    def InSelection(self):
        return self.tabcontrol.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.tabcontrol.InSelection = value

    @property
    def IsVisible(self):
        return self.tabcontrol.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.tabcontrol.IsVisible = value

    @property
    def Layout(self):
        return AcLayoutType(self.tabcontrol.Layout)

    @property
    def LayoutID(self):
        return self.tabcontrol.LayoutID

    @property
    def Left(self):
        return self.tabcontrol.Left

    @Left.setter
    def Left(self, value):
        self.tabcontrol.Left = value

    @property
    def LeftPadding(self):
        return self.tabcontrol.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.tabcontrol.LeftPadding = value

    @property
    def MultiRow(self):
        return self.tabcontrol.MultiRow

    @MultiRow.setter
    def MultiRow(self, value):
        self.tabcontrol.MultiRow = value

    @property
    def Name(self):
        return self.tabcontrol.Name

    @Name.setter
    def Name(self, value):
        self.tabcontrol.Name = value

    @property
    def OldValue(self):
        return self.tabcontrol.OldValue

    @property
    def OnChange(self):
        return self.tabcontrol.OnChange

    @OnChange.setter
    def OnChange(self, value):
        self.tabcontrol.OnChange = value

    @property
    def OnClick(self):
        return self.tabcontrol.OnClick

    @OnClick.setter
    def OnClick(self, value):
        self.tabcontrol.OnClick = value

    @property
    def OnDblClick(self):
        return self.tabcontrol.OnDblClick

    @OnDblClick.setter
    def OnDblClick(self, value):
        self.tabcontrol.OnDblClick = value

    @property
    def OnKeyDown(self):
        return self.tabcontrol.OnKeyDown

    @OnKeyDown.setter
    def OnKeyDown(self, value):
        self.tabcontrol.OnKeyDown = value

    @property
    def OnKeyPress(self):
        return self.tabcontrol.OnKeyPress

    @OnKeyPress.setter
    def OnKeyPress(self, value):
        self.tabcontrol.OnKeyPress = value

    @property
    def OnKeyUp(self):
        return self.tabcontrol.OnKeyUp

    @OnKeyUp.setter
    def OnKeyUp(self, value):
        self.tabcontrol.OnKeyUp = value

    @property
    def OnMouseDown(self):
        return self.tabcontrol.OnMouseDown

    @OnMouseDown.setter
    def OnMouseDown(self, value):
        self.tabcontrol.OnMouseDown = value

    @property
    def OnMouseMove(self):
        return self.tabcontrol.OnMouseMove

    @OnMouseMove.setter
    def OnMouseMove(self, value):
        self.tabcontrol.OnMouseMove = value

    @property
    def OnMouseUp(self):
        return self.tabcontrol.OnMouseUp

    @OnMouseUp.setter
    def OnMouseUp(self, value):
        self.tabcontrol.OnMouseUp = value

    @property
    def Pages(self):
        return Pages(self.tabcontrol.Pages)

    @property
    def Parent(self):
        return self.tabcontrol.Parent

    @property
    def PressedColor(self):
        return self.tabcontrol.PressedColor

    @PressedColor.setter
    def PressedColor(self, value):
        self.tabcontrol.PressedColor = value

    @property
    def PressedForeColor(self):
        return self.tabcontrol.PressedForeColor

    @PressedForeColor.setter
    def PressedForeColor(self, value):
        self.tabcontrol.PressedForeColor = value

    @property
    def PressedForeShade(self):
        return self.tabcontrol.PressedForeShade

    @PressedForeShade.setter
    def PressedForeShade(self, value):
        self.tabcontrol.PressedForeShade = value

    @property
    def PressedForeThemeColorIndex(self):
        return self.tabcontrol.PressedForeThemeColorIndex

    @PressedForeThemeColorIndex.setter
    def PressedForeThemeColorIndex(self, value):
        self.tabcontrol.PressedForeThemeColorIndex = value

    @property
    def PressedForeTint(self):
        return self.tabcontrol.PressedForeTint

    @PressedForeTint.setter
    def PressedForeTint(self, value):
        self.tabcontrol.PressedForeTint = value

    @property
    def PressedShade(self):
        return self.tabcontrol.PressedShade

    @PressedShade.setter
    def PressedShade(self, value):
        self.tabcontrol.PressedShade = value

    @property
    def PressedThemeColorIndex(self):
        return self.tabcontrol.PressedThemeColorIndex

    @PressedThemeColorIndex.setter
    def PressedThemeColorIndex(self, value):
        self.tabcontrol.PressedThemeColorIndex = value

    @property
    def PressedTint(self):
        return self.tabcontrol.PressedTint

    @PressedTint.setter
    def PressedTint(self, value):
        self.tabcontrol.PressedTint = value

    @property
    def Properties(self):
        return Properties(self.tabcontrol.Properties)

    @property
    def RightPadding(self):
        return self.tabcontrol.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.tabcontrol.RightPadding = value

    @property
    def Section(self):
        return self.tabcontrol.Section

    @Section.setter
    def Section(self, value):
        self.tabcontrol.Section = value

    @property
    def Shape(self):
        return self.tabcontrol.Shape

    @Shape.setter
    def Shape(self, value):
        self.tabcontrol.Shape = value

    @property
    def ShortcutMenuBar(self):
        return self.tabcontrol.ShortcutMenuBar

    @ShortcutMenuBar.setter
    def ShortcutMenuBar(self, value):
        self.tabcontrol.ShortcutMenuBar = value

    @property
    def StatusBarText(self):
        return self.tabcontrol.StatusBarText

    @StatusBarText.setter
    def StatusBarText(self, value):
        self.tabcontrol.StatusBarText = value

    @property
    def Style(self):
        return self.tabcontrol.Style

    @Style.setter
    def Style(self, value):
        self.tabcontrol.Style = value

    @property
    def TabFixedHeight(self):
        return self.tabcontrol.TabFixedHeight

    @TabFixedHeight.setter
    def TabFixedHeight(self, value):
        self.tabcontrol.TabFixedHeight = value

    @property
    def TabFixedWidth(self):
        return self.tabcontrol.TabFixedWidth

    @TabFixedWidth.setter
    def TabFixedWidth(self, value):
        self.tabcontrol.TabFixedWidth = value

    @property
    def TabIndex(self):
        return self.tabcontrol.TabIndex

    @TabIndex.setter
    def TabIndex(self, value):
        self.tabcontrol.TabIndex = value

    @property
    def TabStop(self):
        return self.tabcontrol.TabStop

    @TabStop.setter
    def TabStop(self, value):
        self.tabcontrol.TabStop = value

    @property
    def Tag(self):
        return self.tabcontrol.Tag

    @Tag.setter
    def Tag(self, value):
        self.tabcontrol.Tag = value

    @property
    def ThemeFontIndex(self):
        return self.tabcontrol.ThemeFontIndex

    @ThemeFontIndex.setter
    def ThemeFontIndex(self, value):
        self.tabcontrol.ThemeFontIndex = value

    @property
    def Top(self):
        return self.tabcontrol.Top

    @Top.setter
    def Top(self, value):
        self.tabcontrol.Top = value

    @property
    def TopPadding(self):
        return self.tabcontrol.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.tabcontrol.TopPadding = value

    @property
    def UseTheme(self):
        return self.tabcontrol.UseTheme

    @UseTheme.setter
    def UseTheme(self, value):
        self.tabcontrol.UseTheme = value

    @property
    def Value(self):
        return self.tabcontrol.Value

    @Value.setter
    def Value(self, value):
        self.tabcontrol.Value = value

    @property
    def VerticalAnchor(self):
        return self.tabcontrol.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.tabcontrol.VerticalAnchor = value

    @property
    def Visible(self):
        return self.tabcontrol.Visible

    @Visible.setter
    def Visible(self, value):
        self.tabcontrol.Visible = value

    @property
    def Width(self):
        return self.tabcontrol.Width

    @Width.setter
    def Width(self, value):
        self.tabcontrol.Width = value

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.tabcontrol.Move(*args, **arguments)

    def SizeToFit(self):
        self.tabcontrol.SizeToFit()

class TempVar:

    def __init__(self, tempvar=None):
        self.tempvar = tempvar

    @property
    def Name(self):
        return self.tempvar.Name

    @property
    def Value(self):
        return self.tempvar.Value

    @Value.setter
    def Value(self, value):
        self.tempvar.Value = value

class TempVars:

    def __init__(self, tempvars=None):
        self.tempvars = tempvars

    @property
    def Application(self):
        return self.tempvars.Application

    @property
    def Count(self):
        return self.tempvars.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.tempvars.Item):
            return self.tempvars.Item(*args, **arguments)
        else:
            return self.tempvars.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.tempvars.Parent

    def Add(self, *args, Name=None, Value=None):
        arguments = {"Name": Name, "Value": Value}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.tempvars.Add(*args, **arguments)

    def Remove(self, *args, var=None):
        arguments = {"var": var}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.tempvars.Remove(*args, **arguments)

    def RemoveAll(self):
        self.tempvars.RemoveAll()

class TextBox:

    def __init__(self, textbox=None):
        self.textbox = textbox

    @property
    def AddColon(self):
        return self.textbox.AddColon

    @AddColon.setter
    def AddColon(self, value):
        self.textbox.AddColon = value

    @property
    def AllowAutoCorrect(self):
        return self.textbox.AllowAutoCorrect

    @AllowAutoCorrect.setter
    def AllowAutoCorrect(self, value):
        self.textbox.AllowAutoCorrect = value

    @property
    def Application(self):
        return self.textbox.Application

    @property
    def AsianLineBreak(self):
        return self.textbox.AsianLineBreak

    @AsianLineBreak.setter
    def AsianLineBreak(self, value):
        self.textbox.AsianLineBreak = value

    @property
    def AutoLabel(self):
        return self.textbox.AutoLabel

    @AutoLabel.setter
    def AutoLabel(self, value):
        self.textbox.AutoLabel = value

    @property
    def AutoTab(self):
        return self.textbox.AutoTab

    @AutoTab.setter
    def AutoTab(self, value):
        self.textbox.AutoTab = value

    @property
    def BackColor(self):
        return self.textbox.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.textbox.BackColor = value

    @property
    def BackShade(self):
        return self.textbox.BackShade

    @BackShade.setter
    def BackShade(self, value):
        self.textbox.BackShade = value

    @property
    def BackStyle(self):
        return self.textbox.BackStyle

    @BackStyle.setter
    def BackStyle(self, value):
        self.textbox.BackStyle = value

    @property
    def BackThemeColorIndex(self):
        return self.textbox.BackThemeColorIndex

    @BackThemeColorIndex.setter
    def BackThemeColorIndex(self, value):
        self.textbox.BackThemeColorIndex = value

    @property
    def BackTint(self):
        return self.textbox.BackTint

    @BackTint.setter
    def BackTint(self, value):
        self.textbox.BackTint = value

    @property
    def BorderColor(self):
        return self.textbox.BorderColor

    @BorderColor.setter
    def BorderColor(self, value):
        self.textbox.BorderColor = value

    @property
    def BorderShade(self):
        return self.textbox.BorderShade

    @BorderShade.setter
    def BorderShade(self, value):
        self.textbox.BorderShade = value

    @property
    def BorderStyle(self):
        return self.textbox.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.textbox.BorderStyle = value

    @property
    def BorderThemeColorIndex(self):
        return self.textbox.BorderThemeColorIndex

    @BorderThemeColorIndex.setter
    def BorderThemeColorIndex(self, value):
        self.textbox.BorderThemeColorIndex = value

    @property
    def BorderTint(self):
        return self.textbox.BorderTint

    @BorderTint.setter
    def BorderTint(self, value):
        self.textbox.BorderTint = value

    @property
    def BorderWidth(self):
        return self.textbox.BorderWidth

    @BorderWidth.setter
    def BorderWidth(self, value):
        self.textbox.BorderWidth = value

    @property
    def BottomMargin(self):
        return self.textbox.BottomMargin

    @BottomMargin.setter
    def BottomMargin(self, value):
        self.textbox.BottomMargin = value

    @property
    def BottomPadding(self):
        return self.textbox.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.textbox.BottomPadding = value

    @property
    def CanGrow(self):
        return self.textbox.CanGrow

    @CanGrow.setter
    def CanGrow(self, value):
        self.textbox.CanGrow = value

    @property
    def CanShrink(self):
        return self.textbox.CanShrink

    @CanShrink.setter
    def CanShrink(self, value):
        self.textbox.CanShrink = value

    @property
    def ColumnHidden(self):
        return self.textbox.ColumnHidden

    @ColumnHidden.setter
    def ColumnHidden(self, value):
        self.textbox.ColumnHidden = value

    @property
    def ColumnOrder(self):
        return self.textbox.ColumnOrder

    @ColumnOrder.setter
    def ColumnOrder(self, value):
        self.textbox.ColumnOrder = value

    @property
    def ColumnWidth(self):
        return self.textbox.ColumnWidth

    @ColumnWidth.setter
    def ColumnWidth(self, value):
        self.textbox.ColumnWidth = value

    @property
    def Controls(self):
        return Controls(self.textbox.Controls)

    @property
    def ControlSource(self):
        return self.textbox.ControlSource

    @ControlSource.setter
    def ControlSource(self, value):
        self.textbox.ControlSource = value

    @property
    def ControlTipText(self):
        return self.textbox.ControlTipText

    @ControlTipText.setter
    def ControlTipText(self, value):
        self.textbox.ControlTipText = value

    @property
    def ControlType(self):
        return self.textbox.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.textbox.ControlType = value

    @property
    def DecimalPlaces(self):
        return self.textbox.DecimalPlaces

    @DecimalPlaces.setter
    def DecimalPlaces(self, value):
        self.textbox.DecimalPlaces = value

    @property
    def DefaultValue(self):
        return self.textbox.DefaultValue

    @DefaultValue.setter
    def DefaultValue(self, value):
        self.textbox.DefaultValue = value

    @property
    def DisplayAsHyperlink(self):
        return self.textbox.DisplayAsHyperlink

    @DisplayAsHyperlink.setter
    def DisplayAsHyperlink(self, value):
        self.textbox.DisplayAsHyperlink = value

    @property
    def DisplayWhen(self):
        return self.textbox.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.textbox.DisplayWhen = value

    @property
    def Enabled(self):
        return self.textbox.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.textbox.Enabled = value

    @property
    def EnterKeyBehavior(self):
        return self.textbox.EnterKeyBehavior

    @EnterKeyBehavior.setter
    def EnterKeyBehavior(self, value):
        self.textbox.EnterKeyBehavior = value

    @property
    def EventProcPrefix(self):
        return self.textbox.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.textbox.EventProcPrefix = value

    @property
    def FilterLookup(self):
        return self.textbox.FilterLookup

    @FilterLookup.setter
    def FilterLookup(self, value):
        self.textbox.FilterLookup = value

    @property
    def FontBold(self):
        return self.textbox.FontBold

    @property
    def FontItalic(self):
        return self.textbox.FontItalic

    @property
    def FontName(self):
        return self.textbox.FontName

    @property
    def FontSize(self):
        return self.textbox.FontSize

    @property
    def FontUnderline(self):
        return self.textbox.FontUnderline

    @property
    def FontWeight(self):
        return self.textbox.FontWeight

    @FontWeight.setter
    def FontWeight(self, value):
        self.textbox.FontWeight = value

    @property
    def ForeColor(self):
        return self.textbox.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.textbox.ForeColor = value

    @property
    def ForeShade(self):
        return self.textbox.ForeShade

    @ForeShade.setter
    def ForeShade(self, value):
        self.textbox.ForeShade = value

    @property
    def ForeThemeColorIndex(self):
        return self.textbox.ForeThemeColorIndex

    @ForeThemeColorIndex.setter
    def ForeThemeColorIndex(self, value):
        self.textbox.ForeThemeColorIndex = value

    @property
    def ForeTint(self):
        return self.textbox.ForeTint

    @ForeTint.setter
    def ForeTint(self, value):
        self.textbox.ForeTint = value

    @property
    def Format(self):
        return self.textbox.Format

    @Format.setter
    def Format(self, value):
        self.textbox.Format = value

    @property
    def FormatConditions(self):
        return self.textbox.FormatConditions

    @property
    def GridlineColor(self):
        return self.textbox.GridlineColor

    @GridlineColor.setter
    def GridlineColor(self, value):
        self.textbox.GridlineColor = value

    @property
    def GridlineShade(self):
        return self.textbox.GridlineShade

    @GridlineShade.setter
    def GridlineShade(self, value):
        self.textbox.GridlineShade = value

    @property
    def GridlineStyleBottom(self):
        return self.textbox.GridlineStyleBottom

    @GridlineStyleBottom.setter
    def GridlineStyleBottom(self, value):
        self.textbox.GridlineStyleBottom = value

    @property
    def GridlineStyleLeft(self):
        return self.textbox.GridlineStyleLeft

    @GridlineStyleLeft.setter
    def GridlineStyleLeft(self, value):
        self.textbox.GridlineStyleLeft = value

    @property
    def GridlineStyleRight(self):
        return self.textbox.GridlineStyleRight

    @GridlineStyleRight.setter
    def GridlineStyleRight(self, value):
        self.textbox.GridlineStyleRight = value

    @property
    def GridlineStyleTop(self):
        return self.textbox.GridlineStyleTop

    @GridlineStyleTop.setter
    def GridlineStyleTop(self, value):
        self.textbox.GridlineStyleTop = value

    @property
    def GridlineThemeColorIndex(self):
        return self.textbox.GridlineThemeColorIndex

    @GridlineThemeColorIndex.setter
    def GridlineThemeColorIndex(self, value):
        self.textbox.GridlineThemeColorIndex = value

    @property
    def GridlineTint(self):
        return self.textbox.GridlineTint

    @GridlineTint.setter
    def GridlineTint(self, value):
        self.textbox.GridlineTint = value

    @property
    def GridlineWidthBottom(self):
        return self.textbox.GridlineWidthBottom

    @GridlineWidthBottom.setter
    def GridlineWidthBottom(self, value):
        self.textbox.GridlineWidthBottom = value

    @property
    def GridlineWidthLeft(self):
        return self.textbox.GridlineWidthLeft

    @GridlineWidthLeft.setter
    def GridlineWidthLeft(self, value):
        self.textbox.GridlineWidthLeft = value

    @property
    def GridlineWidthRight(self):
        return self.textbox.GridlineWidthRight

    @GridlineWidthRight.setter
    def GridlineWidthRight(self, value):
        self.textbox.GridlineWidthRight = value

    @property
    def GridlineWidthTop(self):
        return self.textbox.GridlineWidthTop

    @GridlineWidthTop.setter
    def GridlineWidthTop(self, value):
        self.textbox.GridlineWidthTop = value

    @property
    def Height(self):
        return self.textbox.Height

    @Height.setter
    def Height(self, value):
        self.textbox.Height = value

    @property
    def HelpContextId(self):
        return self.textbox.HelpContextId

    @HelpContextId.setter
    def HelpContextId(self, value):
        self.textbox.HelpContextId = value

    @property
    def HideDuplicates(self):
        return self.textbox.HideDuplicates

    @HideDuplicates.setter
    def HideDuplicates(self, value):
        self.textbox.HideDuplicates = value

    @property
    def HorizontalAnchor(self):
        return self.textbox.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.textbox.HorizontalAnchor = value

    @property
    def Hyperlink(self):
        return self.textbox.Hyperlink

    @property
    def IMEHold(self):
        return self.textbox.IMEHold

    @IMEHold.setter
    def IMEHold(self, value):
        self.textbox.IMEHold = value

    @property
    def InputMask(self):
        return self.textbox.InputMask

    @InputMask.setter
    def InputMask(self, value):
        self.textbox.InputMask = value

    @property
    def InSelection(self):
        return self.textbox.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.textbox.InSelection = value

    @property
    def IsHyperlink(self):
        return self.textbox.IsHyperlink

    @IsHyperlink.setter
    def IsHyperlink(self, value):
        self.textbox.IsHyperlink = value

    @property
    def IsVisible(self):
        return self.textbox.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.textbox.IsVisible = value

    @property
    def LabelAlign(self):
        return self.textbox.LabelAlign

    @LabelAlign.setter
    def LabelAlign(self, value):
        self.textbox.LabelAlign = value

    @property
    def LabelX(self):
        return self.textbox.LabelX

    @LabelX.setter
    def LabelX(self, value):
        self.textbox.LabelX = value

    @property
    def LabelY(self):
        return self.textbox.LabelY

    @LabelY.setter
    def LabelY(self, value):
        self.textbox.LabelY = value

    @property
    def Layout(self):
        return AcLayoutType(self.textbox.Layout)

    @property
    def LayoutID(self):
        return self.textbox.LayoutID

    @property
    def Left(self):
        return self.textbox.Left

    @Left.setter
    def Left(self, value):
        self.textbox.Left = value

    @property
    def LeftMargin(self):
        return self.textbox.LeftMargin

    @LeftMargin.setter
    def LeftMargin(self, value):
        self.textbox.LeftMargin = value

    @property
    def LeftPadding(self):
        return self.textbox.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.textbox.LeftPadding = value

    @property
    def LineSpacing(self):
        return self.textbox.LineSpacing

    @LineSpacing.setter
    def LineSpacing(self, value):
        self.textbox.LineSpacing = value

    @property
    def Locked(self):
        return self.textbox.Locked

    @Locked.setter
    def Locked(self, value):
        self.textbox.Locked = value

    @property
    def Name(self):
        return self.textbox.Name

    @Name.setter
    def Name(self, value):
        self.textbox.Name = value

    @property
    def OldBorderStyle(self):
        return self.textbox.OldBorderStyle

    @OldBorderStyle.setter
    def OldBorderStyle(self, value):
        self.textbox.OldBorderStyle = value

    @property
    def OldValue(self):
        return self.textbox.OldValue

    @property
    def OnChange(self):
        return self.textbox.OnChange

    @OnChange.setter
    def OnChange(self, value):
        self.textbox.OnChange = value

    @property
    def OnClick(self):
        return self.textbox.OnClick

    @OnClick.setter
    def OnClick(self, value):
        self.textbox.OnClick = value

    @property
    def OnDblClick(self):
        return self.textbox.OnDblClick

    @OnDblClick.setter
    def OnDblClick(self, value):
        self.textbox.OnDblClick = value

    @property
    def OnDirty(self):
        return self.textbox.OnDirty

    @OnDirty.setter
    def OnDirty(self, value):
        self.textbox.OnDirty = value

    @property
    def OnEnter(self):
        return self.textbox.OnEnter

    @OnEnter.setter
    def OnEnter(self, value):
        self.textbox.OnEnter = value

    @property
    def OnExit(self):
        return self.textbox.OnExit

    @OnExit.setter
    def OnExit(self, value):
        self.textbox.OnExit = value

    @property
    def OnGotFocus(self):
        return self.textbox.OnGotFocus

    @OnGotFocus.setter
    def OnGotFocus(self, value):
        self.textbox.OnGotFocus = value

    @property
    def OnKeyDown(self):
        return self.textbox.OnKeyDown

    @OnKeyDown.setter
    def OnKeyDown(self, value):
        self.textbox.OnKeyDown = value

    @property
    def OnKeyPress(self):
        return self.textbox.OnKeyPress

    @OnKeyPress.setter
    def OnKeyPress(self, value):
        self.textbox.OnKeyPress = value

    @property
    def OnKeyUp(self):
        return self.textbox.OnKeyUp

    @OnKeyUp.setter
    def OnKeyUp(self, value):
        self.textbox.OnKeyUp = value

    @property
    def OnLostFocus(self):
        return self.textbox.OnLostFocus

    @OnLostFocus.setter
    def OnLostFocus(self, value):
        self.textbox.OnLostFocus = value

    @property
    def OnMouseDown(self):
        return self.textbox.OnMouseDown

    @OnMouseDown.setter
    def OnMouseDown(self, value):
        self.textbox.OnMouseDown = value

    @property
    def OnMouseMove(self):
        return self.textbox.OnMouseMove

    @OnMouseMove.setter
    def OnMouseMove(self, value):
        self.textbox.OnMouseMove = value

    @property
    def OnMouseUp(self):
        return self.textbox.OnMouseUp

    @OnMouseUp.setter
    def OnMouseUp(self, value):
        self.textbox.OnMouseUp = value

    @property
    def OnUndo(self):
        return self.textbox.OnUndo

    @OnUndo.setter
    def OnUndo(self, value):
        self.textbox.OnUndo = value

    @property
    def Parent(self):
        return self.textbox.Parent

    @property
    def PostalAddress(self):
        return self.textbox.PostalAddress

    @PostalAddress.setter
    def PostalAddress(self, value):
        self.textbox.PostalAddress = value

    @property
    def Properties(self):
        return Properties(self.textbox.Properties)

    @property
    def ReadingOrder(self):
        return self.textbox.ReadingOrder

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.textbox.ReadingOrder = value

    @property
    def RightMargin(self):
        return self.textbox.RightMargin

    @RightMargin.setter
    def RightMargin(self, value):
        self.textbox.RightMargin = value

    @property
    def RightPadding(self):
        return self.textbox.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.textbox.RightPadding = value

    @property
    def RunningSum(self):
        return self.textbox.RunningSum

    @RunningSum.setter
    def RunningSum(self, value):
        self.textbox.RunningSum = value

    @property
    def ScrollBarAlign(self):
        return self.textbox.ScrollBarAlign

    @ScrollBarAlign.setter
    def ScrollBarAlign(self, value):
        self.textbox.ScrollBarAlign = value

    @property
    def ScrollBars(self):
        return self.textbox.ScrollBars

    @ScrollBars.setter
    def ScrollBars(self, value):
        self.textbox.ScrollBars = value

    @property
    def Section(self):
        return self.textbox.Section

    @Section.setter
    def Section(self, value):
        self.textbox.Section = value

    @property
    def SelLength(self):
        return self.textbox.SelLength

    @SelLength.setter
    def SelLength(self, value):
        self.textbox.SelLength = value

    @property
    def SelStart(self):
        return self.textbox.SelStart

    @SelStart.setter
    def SelStart(self, value):
        self.textbox.SelStart = value

    @property
    def SelText(self):
        return self.textbox.SelText

    @SelText.setter
    def SelText(self, value):
        self.textbox.SelText = value

    @property
    def ShortcutMenuBar(self):
        return self.textbox.ShortcutMenuBar

    @ShortcutMenuBar.setter
    def ShortcutMenuBar(self, value):
        self.textbox.ShortcutMenuBar = value

    @property
    def ShowDatePicker(self):
        return self.textbox.ShowDatePicker

    @ShowDatePicker.setter
    def ShowDatePicker(self, value):
        self.textbox.ShowDatePicker = value

    @property
    def SmartTags(self):
        return SmartTags(self.textbox.SmartTags)

    @property
    def SpecialEffect(self):
        return self.textbox.SpecialEffect

    @SpecialEffect.setter
    def SpecialEffect(self, value):
        self.textbox.SpecialEffect = value

    @property
    def StatusBarText(self):
        return self.textbox.StatusBarText

    @StatusBarText.setter
    def StatusBarText(self, value):
        self.textbox.StatusBarText = value

    @property
    def TabIndex(self):
        return self.textbox.TabIndex

    @TabIndex.setter
    def TabIndex(self, value):
        self.textbox.TabIndex = value

    @property
    def TabStop(self):
        return self.textbox.TabStop

    @TabStop.setter
    def TabStop(self, value):
        self.textbox.TabStop = value

    @property
    def Tag(self):
        return self.textbox.Tag

    @Tag.setter
    def Tag(self, value):
        self.textbox.Tag = value

    @property
    def Text(self):
        return self.textbox.Text

    @Text.setter
    def Text(self, value):
        self.textbox.Text = value

    @property
    def TextAlign(self):
        return self.textbox.TextAlign

    @TextAlign.setter
    def TextAlign(self, value):
        self.textbox.TextAlign = value

    @property
    def TextFormat(self):
        return self.textbox.TextFormat

    @TextFormat.setter
    def TextFormat(self, value):
        self.textbox.TextFormat = value

    @property
    def ThemeFontIndex(self):
        return self.textbox.ThemeFontIndex

    @ThemeFontIndex.setter
    def ThemeFontIndex(self, value):
        self.textbox.ThemeFontIndex = value

    @property
    def Top(self):
        return self.textbox.Top

    @Top.setter
    def Top(self, value):
        self.textbox.Top = value

    @property
    def TopMargin(self):
        return self.textbox.TopMargin

    @TopMargin.setter
    def TopMargin(self, value):
        self.textbox.TopMargin = value

    @property
    def TopPadding(self):
        return self.textbox.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.textbox.TopPadding = value

    @property
    def ValidationRule(self):
        return self.textbox.ValidationRule

    @ValidationRule.setter
    def ValidationRule(self, value):
        self.textbox.ValidationRule = value

    @property
    def ValidationText(self):
        return self.textbox.ValidationText

    @ValidationText.setter
    def ValidationText(self, value):
        self.textbox.ValidationText = value

    @property
    def Value(self):
        return self.textbox.Value

    @Value.setter
    def Value(self, value):
        self.textbox.Value = value

    @property
    def Vertical(self):
        return self.textbox.Vertical

    @Vertical.setter
    def Vertical(self, value):
        self.textbox.Vertical = value

    @property
    def VerticalAnchor(self):
        return self.textbox.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.textbox.VerticalAnchor = value

    @property
    def Visible(self):
        return self.textbox.Visible

    @Visible.setter
    def Visible(self, value):
        self.textbox.Visible = value

    @property
    def Width(self):
        return self.textbox.Width

    @Width.setter
    def Width(self, value):
        self.textbox.Width = value

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.textbox.Move(*args, **arguments)

    def Requery(self):
        self.textbox.Requery()

    def SetFocus(self):
        return self.textbox.SetFocus()

    def SizeToFit(self):
        self.textbox.SizeToFit()

    def Undo(self):
        self.textbox.Undo()

class ToggleButton:

    def __init__(self, togglebutton=None):
        self.togglebutton = togglebutton

    @property
    def AddColon(self):
        return self.togglebutton.AddColon

    @AddColon.setter
    def AddColon(self, value):
        self.togglebutton.AddColon = value

    @property
    def Application(self):
        return self.togglebutton.Application

    @property
    def AutoLabel(self):
        return self.togglebutton.AutoLabel

    @AutoLabel.setter
    def AutoLabel(self, value):
        self.togglebutton.AutoLabel = value

    @property
    def BackColor(self):
        return self.togglebutton.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.togglebutton.BackColor = value

    @property
    def BackShade(self):
        return self.togglebutton.BackShade

    @BackShade.setter
    def BackShade(self, value):
        self.togglebutton.BackShade = value

    @property
    def BackThemeColorIndex(self):
        return self.togglebutton.BackThemeColorIndex

    @BackThemeColorIndex.setter
    def BackThemeColorIndex(self, value):
        self.togglebutton.BackThemeColorIndex = value

    @property
    def BackTint(self):
        return self.togglebutton.BackTint

    @BackTint.setter
    def BackTint(self, value):
        self.togglebutton.BackTint = value

    @property
    def Bevel(self):
        return self.togglebutton.Bevel

    @Bevel.setter
    def Bevel(self, value):
        self.togglebutton.Bevel = value

    @property
    def BorderColor(self):
        return self.togglebutton.BorderColor

    @BorderColor.setter
    def BorderColor(self, value):
        self.togglebutton.BorderColor = value

    @property
    def BorderShade(self):
        return self.togglebutton.BorderShade

    @BorderShade.setter
    def BorderShade(self, value):
        self.togglebutton.BorderShade = value

    @property
    def BorderStyle(self):
        return self.togglebutton.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.togglebutton.BorderStyle = value

    @property
    def BorderThemeColorIndex(self):
        return self.togglebutton.BorderThemeColorIndex

    @BorderThemeColorIndex.setter
    def BorderThemeColorIndex(self, value):
        self.togglebutton.BorderThemeColorIndex = value

    @property
    def BorderTint(self):
        return self.togglebutton.BorderTint

    @BorderTint.setter
    def BorderTint(self, value):
        self.togglebutton.BorderTint = value

    @property
    def BorderWidth(self):
        return self.togglebutton.BorderWidth

    @BorderWidth.setter
    def BorderWidth(self, value):
        self.togglebutton.BorderWidth = value

    @property
    def BottomPadding(self):
        return self.togglebutton.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.togglebutton.BottomPadding = value

    @property
    def Caption(self):
        return self.togglebutton.Caption

    @Caption.setter
    def Caption(self, value):
        self.togglebutton.Caption = value

    @property
    def ColumnHidden(self):
        return self.togglebutton.ColumnHidden

    @ColumnHidden.setter
    def ColumnHidden(self, value):
        self.togglebutton.ColumnHidden = value

    @property
    def ColumnOrder(self):
        return self.togglebutton.ColumnOrder

    @ColumnOrder.setter
    def ColumnOrder(self, value):
        self.togglebutton.ColumnOrder = value

    @property
    def ColumnWidth(self):
        return self.togglebutton.ColumnWidth

    @ColumnWidth.setter
    def ColumnWidth(self, value):
        self.togglebutton.ColumnWidth = value

    @property
    def Controls(self):
        return Controls(self.togglebutton.Controls)

    @property
    def ControlSource(self):
        return self.togglebutton.ControlSource

    @ControlSource.setter
    def ControlSource(self, value):
        self.togglebutton.ControlSource = value

    @property
    def ControlTipText(self):
        return self.togglebutton.ControlTipText

    @ControlTipText.setter
    def ControlTipText(self, value):
        self.togglebutton.ControlTipText = value

    @property
    def ControlType(self):
        return self.togglebutton.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.togglebutton.ControlType = value

    @property
    def DefaultValue(self):
        return self.togglebutton.DefaultValue

    @DefaultValue.setter
    def DefaultValue(self, value):
        self.togglebutton.DefaultValue = value

    @property
    def DisplayWhen(self):
        return self.togglebutton.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.togglebutton.DisplayWhen = value

    @property
    def Enabled(self):
        return self.togglebutton.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.togglebutton.Enabled = value

    @property
    def EventProcPrefix(self):
        return self.togglebutton.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.togglebutton.EventProcPrefix = value

    @property
    def FontBold(self):
        return self.togglebutton.FontBold

    @property
    def FontItalic(self):
        return self.togglebutton.FontItalic

    @property
    def FontName(self):
        return self.togglebutton.FontName

    @property
    def FontSize(self):
        return self.togglebutton.FontSize

    @property
    def FontUnderline(self):
        return self.togglebutton.FontUnderline

    @property
    def FontWeight(self):
        return self.togglebutton.FontWeight

    @FontWeight.setter
    def FontWeight(self, value):
        self.togglebutton.FontWeight = value

    @property
    def ForeColor(self):
        return self.togglebutton.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.togglebutton.ForeColor = value

    @property
    def ForeShade(self):
        return self.togglebutton.ForeShade

    @ForeShade.setter
    def ForeShade(self, value):
        self.togglebutton.ForeShade = value

    @property
    def ForeThemeColorIndex(self):
        return self.togglebutton.ForeThemeColorIndex

    @ForeThemeColorIndex.setter
    def ForeThemeColorIndex(self, value):
        self.togglebutton.ForeThemeColorIndex = value

    @property
    def ForeTint(self):
        return self.togglebutton.ForeTint

    @ForeTint.setter
    def ForeTint(self, value):
        self.togglebutton.ForeTint = value

    @property
    def Glow(self):
        return self.togglebutton.Glow

    @Glow.setter
    def Glow(self, value):
        self.togglebutton.Glow = value

    @property
    def Gradient(self):
        return self.togglebutton.Gradient

    @Gradient.setter
    def Gradient(self, value):
        self.togglebutton.Gradient = value

    @property
    def GridlineColor(self):
        return self.togglebutton.GridlineColor

    @GridlineColor.setter
    def GridlineColor(self, value):
        self.togglebutton.GridlineColor = value

    @property
    def GridlineShade(self):
        return self.togglebutton.GridlineShade

    @GridlineShade.setter
    def GridlineShade(self, value):
        self.togglebutton.GridlineShade = value

    @property
    def GridlineStyleBottom(self):
        return self.togglebutton.GridlineStyleBottom

    @GridlineStyleBottom.setter
    def GridlineStyleBottom(self, value):
        self.togglebutton.GridlineStyleBottom = value

    @property
    def GridlineStyleLeft(self):
        return self.togglebutton.GridlineStyleLeft

    @GridlineStyleLeft.setter
    def GridlineStyleLeft(self, value):
        self.togglebutton.GridlineStyleLeft = value

    @property
    def GridlineStyleRight(self):
        return self.togglebutton.GridlineStyleRight

    @GridlineStyleRight.setter
    def GridlineStyleRight(self, value):
        self.togglebutton.GridlineStyleRight = value

    @property
    def GridlineStyleTop(self):
        return self.togglebutton.GridlineStyleTop

    @GridlineStyleTop.setter
    def GridlineStyleTop(self, value):
        self.togglebutton.GridlineStyleTop = value

    @property
    def GridlineThemeColorIndex(self):
        return self.togglebutton.GridlineThemeColorIndex

    @GridlineThemeColorIndex.setter
    def GridlineThemeColorIndex(self, value):
        self.togglebutton.GridlineThemeColorIndex = value

    @property
    def GridlineTint(self):
        return self.togglebutton.GridlineTint

    @GridlineTint.setter
    def GridlineTint(self, value):
        self.togglebutton.GridlineTint = value

    @property
    def GridlineWidthBottom(self):
        return self.togglebutton.GridlineWidthBottom

    @GridlineWidthBottom.setter
    def GridlineWidthBottom(self, value):
        self.togglebutton.GridlineWidthBottom = value

    @property
    def GridlineWidthLeft(self):
        return self.togglebutton.GridlineWidthLeft

    @GridlineWidthLeft.setter
    def GridlineWidthLeft(self, value):
        self.togglebutton.GridlineWidthLeft = value

    @property
    def GridlineWidthRight(self):
        return self.togglebutton.GridlineWidthRight

    @GridlineWidthRight.setter
    def GridlineWidthRight(self, value):
        self.togglebutton.GridlineWidthRight = value

    @property
    def GridlineWidthTop(self):
        return self.togglebutton.GridlineWidthTop

    @GridlineWidthTop.setter
    def GridlineWidthTop(self, value):
        self.togglebutton.GridlineWidthTop = value

    @property
    def Height(self):
        return self.togglebutton.Height

    @Height.setter
    def Height(self, value):
        self.togglebutton.Height = value

    @property
    def HelpContextId(self):
        return self.togglebutton.HelpContextId

    @HelpContextId.setter
    def HelpContextId(self, value):
        self.togglebutton.HelpContextId = value

    @property
    def HideDuplicates(self):
        return self.togglebutton.HideDuplicates

    @HideDuplicates.setter
    def HideDuplicates(self, value):
        self.togglebutton.HideDuplicates = value

    @property
    def HorizontalAnchor(self):
        return self.togglebutton.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.togglebutton.HorizontalAnchor = value

    @property
    def HoverColor(self):
        return self.togglebutton.HoverColor

    @HoverColor.setter
    def HoverColor(self, value):
        self.togglebutton.HoverColor = value

    @property
    def HoverForeColor(self):
        return self.togglebutton.HoverForeColor

    @HoverForeColor.setter
    def HoverForeColor(self, value):
        self.togglebutton.HoverForeColor = value

    @property
    def HoverForeShade(self):
        return self.togglebutton.HoverForeShade

    @HoverForeShade.setter
    def HoverForeShade(self, value):
        self.togglebutton.HoverForeShade = value

    @property
    def HoverForeThemeColorIndex(self):
        return self.togglebutton.HoverForeThemeColorIndex

    @HoverForeThemeColorIndex.setter
    def HoverForeThemeColorIndex(self, value):
        self.togglebutton.HoverForeThemeColorIndex = value

    @property
    def HoverForeTint(self):
        return self.togglebutton.HoverForeTint

    @HoverForeTint.setter
    def HoverForeTint(self, value):
        self.togglebutton.HoverForeTint = value

    @property
    def HoverShade(self):
        return self.togglebutton.HoverShade

    @HoverShade.setter
    def HoverShade(self, value):
        self.togglebutton.HoverShade = value

    @property
    def HoverThemeColorIndex(self):
        return self.togglebutton.HoverThemeColorIndex

    @HoverThemeColorIndex.setter
    def HoverThemeColorIndex(self, value):
        self.togglebutton.HoverThemeColorIndex = value

    @property
    def HoverTint(self):
        return self.togglebutton.HoverTint

    @HoverTint.setter
    def HoverTint(self, value):
        self.togglebutton.HoverTint = value

    @property
    def InSelection(self):
        return self.togglebutton.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.togglebutton.InSelection = value

    @property
    def IsVisible(self):
        return self.togglebutton.IsVisible

    @IsVisible.setter
    def IsVisible(self, value):
        self.togglebutton.IsVisible = value

    @property
    def LabelAlign(self):
        return self.togglebutton.LabelAlign

    @LabelAlign.setter
    def LabelAlign(self, value):
        self.togglebutton.LabelAlign = value

    @property
    def LabelX(self):
        return self.togglebutton.LabelX

    @LabelX.setter
    def LabelX(self, value):
        self.togglebutton.LabelX = value

    @property
    def LabelY(self):
        return self.togglebutton.LabelY

    @LabelY.setter
    def LabelY(self, value):
        self.togglebutton.LabelY = value

    @property
    def Layout(self):
        return AcLayoutType(self.togglebutton.Layout)

    @property
    def LayoutID(self):
        return self.togglebutton.LayoutID

    @property
    def Left(self):
        return self.togglebutton.Left

    @Left.setter
    def Left(self, value):
        self.togglebutton.Left = value

    @property
    def LeftPadding(self):
        return self.togglebutton.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.togglebutton.LeftPadding = value

    @property
    def Locked(self):
        return self.togglebutton.Locked

    @Locked.setter
    def Locked(self, value):
        self.togglebutton.Locked = value

    @property
    def Name(self):
        return self.togglebutton.Name

    @Name.setter
    def Name(self, value):
        self.togglebutton.Name = value

    @property
    def ObjectPalette(self):
        return self.togglebutton.ObjectPalette

    @ObjectPalette.setter
    def ObjectPalette(self, value):
        self.togglebutton.ObjectPalette = value

    @property
    def OldValue(self):
        return self.togglebutton.OldValue

    @property
    def OnClick(self):
        return self.togglebutton.OnClick

    @OnClick.setter
    def OnClick(self, value):
        self.togglebutton.OnClick = value

    @property
    def OnDblClick(self):
        return self.togglebutton.OnDblClick

    @OnDblClick.setter
    def OnDblClick(self, value):
        self.togglebutton.OnDblClick = value

    @property
    def OnEnter(self):
        return self.togglebutton.OnEnter

    @OnEnter.setter
    def OnEnter(self, value):
        self.togglebutton.OnEnter = value

    @property
    def OnExit(self):
        return self.togglebutton.OnExit

    @OnExit.setter
    def OnExit(self, value):
        self.togglebutton.OnExit = value

    @property
    def OnGotFocus(self):
        return self.togglebutton.OnGotFocus

    @OnGotFocus.setter
    def OnGotFocus(self, value):
        self.togglebutton.OnGotFocus = value

    @property
    def OnKeyDown(self):
        return self.togglebutton.OnKeyDown

    @OnKeyDown.setter
    def OnKeyDown(self, value):
        self.togglebutton.OnKeyDown = value

    @property
    def OnKeyPress(self):
        return self.togglebutton.OnKeyPress

    @OnKeyPress.setter
    def OnKeyPress(self, value):
        self.togglebutton.OnKeyPress = value

    @property
    def OnKeyUp(self):
        return self.togglebutton.OnKeyUp

    @OnKeyUp.setter
    def OnKeyUp(self, value):
        self.togglebutton.OnKeyUp = value

    @property
    def OnLostFocus(self):
        return self.togglebutton.OnLostFocus

    @OnLostFocus.setter
    def OnLostFocus(self, value):
        self.togglebutton.OnLostFocus = value

    @property
    def OnMouseDown(self):
        return self.togglebutton.OnMouseDown

    @OnMouseDown.setter
    def OnMouseDown(self, value):
        self.togglebutton.OnMouseDown = value

    @property
    def OnMouseMove(self):
        return self.togglebutton.OnMouseMove

    @OnMouseMove.setter
    def OnMouseMove(self, value):
        self.togglebutton.OnMouseMove = value

    @property
    def OnMouseUp(self):
        return self.togglebutton.OnMouseUp

    @OnMouseUp.setter
    def OnMouseUp(self, value):
        self.togglebutton.OnMouseUp = value

    @property
    def OptionValue(self):
        return self.togglebutton.OptionValue

    @OptionValue.setter
    def OptionValue(self, value):
        self.togglebutton.OptionValue = value

    @property
    def Parent(self):
        return self.togglebutton.Parent

    @property
    def Picture(self):
        return self.togglebutton.Picture

    @Picture.setter
    def Picture(self, value):
        self.togglebutton.Picture = value

    @property
    def PictureData(self):
        return self.togglebutton.PictureData

    @PictureData.setter
    def PictureData(self, value):
        self.togglebutton.PictureData = value

    @property
    def PictureType(self):
        return self.togglebutton.PictureType

    @PictureType.setter
    def PictureType(self, value):
        self.togglebutton.PictureType = value

    @property
    def PressedColor(self):
        return self.togglebutton.PressedColor

    @PressedColor.setter
    def PressedColor(self, value):
        self.togglebutton.PressedColor = value

    @property
    def PressedForeColor(self):
        return self.togglebutton.PressedForeColor

    @PressedForeColor.setter
    def PressedForeColor(self, value):
        self.togglebutton.PressedForeColor = value

    @property
    def PressedForeShade(self):
        return self.togglebutton.PressedForeShade

    @PressedForeShade.setter
    def PressedForeShade(self, value):
        self.togglebutton.PressedForeShade = value

    @property
    def PressedForeThemeColorIndex(self):
        return self.togglebutton.PressedForeThemeColorIndex

    @PressedForeThemeColorIndex.setter
    def PressedForeThemeColorIndex(self, value):
        self.togglebutton.PressedForeThemeColorIndex = value

    @property
    def PressedForeTint(self):
        return self.togglebutton.PressedForeTint

    @PressedForeTint.setter
    def PressedForeTint(self, value):
        self.togglebutton.PressedForeTint = value

    @property
    def PressedShade(self):
        return self.togglebutton.PressedShade

    @PressedShade.setter
    def PressedShade(self, value):
        self.togglebutton.PressedShade = value

    @property
    def PressedThemeColorIndex(self):
        return self.togglebutton.PressedThemeColorIndex

    @PressedThemeColorIndex.setter
    def PressedThemeColorIndex(self, value):
        self.togglebutton.PressedThemeColorIndex = value

    @property
    def PressedTint(self):
        return self.togglebutton.PressedTint

    @PressedTint.setter
    def PressedTint(self, value):
        self.togglebutton.PressedTint = value

    @property
    def Properties(self):
        return Properties(self.togglebutton.Properties)

    @property
    def QuickStyle(self):
        return self.togglebutton.QuickStyle

    @QuickStyle.setter
    def QuickStyle(self, value):
        self.togglebutton.QuickStyle = value

    @property
    def ReadingOrder(self):
        return self.togglebutton.ReadingOrder

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.togglebutton.ReadingOrder = value

    @property
    def RightPadding(self):
        return self.togglebutton.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.togglebutton.RightPadding = value

    @property
    def Section(self):
        return self.togglebutton.Section

    @Section.setter
    def Section(self, value):
        self.togglebutton.Section = value

    @property
    def Shadow(self):
        return self.togglebutton.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.togglebutton.Shadow = value

    @property
    def Shape(self):
        return self.togglebutton.Shape

    @Shape.setter
    def Shape(self, value):
        self.togglebutton.Shape = value

    @property
    def ShortcutMenuBar(self):
        return self.togglebutton.ShortcutMenuBar

    @ShortcutMenuBar.setter
    def ShortcutMenuBar(self, value):
        self.togglebutton.ShortcutMenuBar = value

    @property
    def SoftEdges(self):
        return self.togglebutton.SoftEdges

    @SoftEdges.setter
    def SoftEdges(self, value):
        self.togglebutton.SoftEdges = value

    @property
    def StatusBarText(self):
        return self.togglebutton.StatusBarText

    @StatusBarText.setter
    def StatusBarText(self, value):
        self.togglebutton.StatusBarText = value

    @property
    def TabIndex(self):
        return self.togglebutton.TabIndex

    @TabIndex.setter
    def TabIndex(self, value):
        self.togglebutton.TabIndex = value

    @property
    def TabStop(self):
        return self.togglebutton.TabStop

    @TabStop.setter
    def TabStop(self, value):
        self.togglebutton.TabStop = value

    @property
    def Tag(self):
        return self.togglebutton.Tag

    @Tag.setter
    def Tag(self, value):
        self.togglebutton.Tag = value

    @property
    def ThemeFontIndex(self):
        return self.togglebutton.ThemeFontIndex

    @ThemeFontIndex.setter
    def ThemeFontIndex(self, value):
        self.togglebutton.ThemeFontIndex = value

    @property
    def Top(self):
        return self.togglebutton.Top

    @Top.setter
    def Top(self, value):
        self.togglebutton.Top = value

    @property
    def TopPadding(self):
        return self.togglebutton.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.togglebutton.TopPadding = value

    @property
    def TripleState(self):
        return self.togglebutton.TripleState

    @TripleState.setter
    def TripleState(self, value):
        self.togglebutton.TripleState = value

    @property
    def UseTheme(self):
        return self.togglebutton.UseTheme

    @UseTheme.setter
    def UseTheme(self, value):
        self.togglebutton.UseTheme = value

    @property
    def ValidationRule(self):
        return self.togglebutton.ValidationRule

    @ValidationRule.setter
    def ValidationRule(self, value):
        self.togglebutton.ValidationRule = value

    @property
    def ValidationText(self):
        return self.togglebutton.ValidationText

    @ValidationText.setter
    def ValidationText(self, value):
        self.togglebutton.ValidationText = value

    @property
    def Value(self):
        return self.togglebutton.Value

    @Value.setter
    def Value(self, value):
        self.togglebutton.Value = value

    @property
    def VerticalAnchor(self):
        return self.togglebutton.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.togglebutton.VerticalAnchor = value

    @property
    def Visible(self):
        return self.togglebutton.Visible

    @Visible.setter
    def Visible(self, value):
        self.togglebutton.Visible = value

    @property
    def Width(self):
        return self.togglebutton.Width

    @Width.setter
    def Width(self, value):
        self.togglebutton.Width = value

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.togglebutton.Move(*args, **arguments)

    def Requery(self):
        self.togglebutton.Requery()

    def SetFocus(self):
        return self.togglebutton.SetFocus()

    def SizeToFit(self):
        self.togglebutton.SizeToFit()

    def Undo(self):
        self.togglebutton.Undo()

class WebBrowserControl:

    def __init__(self, webbrowsercontrol=None):
        self.webbrowsercontrol = webbrowsercontrol

    @property
    def Application(self):
        return self.webbrowsercontrol.Application

    @property
    def BorderColor(self):
        return self.webbrowsercontrol.BorderColor

    @BorderColor.setter
    def BorderColor(self, value):
        self.webbrowsercontrol.BorderColor = value

    @property
    def BorderShade(self):
        return self.webbrowsercontrol.BorderShade

    @BorderShade.setter
    def BorderShade(self, value):
        self.webbrowsercontrol.BorderShade = value

    @property
    def BorderStyle(self):
        return self.webbrowsercontrol.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.webbrowsercontrol.BorderStyle = value

    @property
    def BorderThemeColorIndex(self):
        return self.webbrowsercontrol.BorderThemeColorIndex

    @BorderThemeColorIndex.setter
    def BorderThemeColorIndex(self, value):
        self.webbrowsercontrol.BorderThemeColorIndex = value

    @property
    def BorderTint(self):
        return self.webbrowsercontrol.BorderTint

    @BorderTint.setter
    def BorderTint(self, value):
        self.webbrowsercontrol.BorderTint = value

    @property
    def BorderWidth(self):
        return self.webbrowsercontrol.BorderWidth

    @BorderWidth.setter
    def BorderWidth(self, value):
        self.webbrowsercontrol.BorderWidth = value

    @property
    def BottomPadding(self):
        return self.webbrowsercontrol.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.webbrowsercontrol.BottomPadding = value

    @property
    def Controls(self):
        return Controls(self.webbrowsercontrol.Controls)

    @property
    def ControlSource(self):
        return self.webbrowsercontrol.ControlSource

    @ControlSource.setter
    def ControlSource(self, value):
        self.webbrowsercontrol.ControlSource = value

    @property
    def ControlTipText(self):
        return self.webbrowsercontrol.ControlTipText

    @ControlTipText.setter
    def ControlTipText(self, value):
        self.webbrowsercontrol.ControlTipText = value

    @property
    def ControlType(self):
        return self.webbrowsercontrol.ControlType

    @ControlType.setter
    def ControlType(self, value):
        self.webbrowsercontrol.ControlType = value

    @property
    def DisplayWhen(self):
        return self.webbrowsercontrol.DisplayWhen

    @DisplayWhen.setter
    def DisplayWhen(self, value):
        self.webbrowsercontrol.DisplayWhen = value

    @property
    def Enabled(self):
        return self.webbrowsercontrol.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.webbrowsercontrol.Enabled = value

    @property
    def EventProcPrefix(self):
        return self.webbrowsercontrol.EventProcPrefix

    @EventProcPrefix.setter
    def EventProcPrefix(self, value):
        self.webbrowsercontrol.EventProcPrefix = value

    @property
    def GridlineColor(self):
        return self.webbrowsercontrol.GridlineColor

    @GridlineColor.setter
    def GridlineColor(self, value):
        self.webbrowsercontrol.GridlineColor = value

    @property
    def GridlineShade(self):
        return self.webbrowsercontrol.GridlineShade

    @GridlineShade.setter
    def GridlineShade(self, value):
        self.webbrowsercontrol.GridlineShade = value

    @property
    def GridlineStyleBottom(self):
        return self.webbrowsercontrol.GridlineStyleBottom

    @GridlineStyleBottom.setter
    def GridlineStyleBottom(self, value):
        self.webbrowsercontrol.GridlineStyleBottom = value

    @property
    def GridlineStyleLeft(self):
        return self.webbrowsercontrol.GridlineStyleLeft

    @GridlineStyleLeft.setter
    def GridlineStyleLeft(self, value):
        self.webbrowsercontrol.GridlineStyleLeft = value

    @property
    def GridlineStyleRight(self):
        return self.webbrowsercontrol.GridlineStyleRight

    @GridlineStyleRight.setter
    def GridlineStyleRight(self, value):
        self.webbrowsercontrol.GridlineStyleRight = value

    @property
    def GridlineStyleTop(self):
        return self.webbrowsercontrol.GridlineStyleTop

    @GridlineStyleTop.setter
    def GridlineStyleTop(self, value):
        self.webbrowsercontrol.GridlineStyleTop = value

    @property
    def GridlineThemeColorIndex(self):
        return self.webbrowsercontrol.GridlineThemeColorIndex

    @GridlineThemeColorIndex.setter
    def GridlineThemeColorIndex(self, value):
        self.webbrowsercontrol.GridlineThemeColorIndex = value

    @property
    def GridlineTint(self):
        return self.webbrowsercontrol.GridlineTint

    @GridlineTint.setter
    def GridlineTint(self, value):
        self.webbrowsercontrol.GridlineTint = value

    @property
    def GridlineWidthBottom(self):
        return self.webbrowsercontrol.GridlineWidthBottom

    @GridlineWidthBottom.setter
    def GridlineWidthBottom(self, value):
        self.webbrowsercontrol.GridlineWidthBottom = value

    @property
    def GridlineWidthLeft(self):
        return self.webbrowsercontrol.GridlineWidthLeft

    @GridlineWidthLeft.setter
    def GridlineWidthLeft(self, value):
        self.webbrowsercontrol.GridlineWidthLeft = value

    @property
    def GridlineWidthRight(self):
        return self.webbrowsercontrol.GridlineWidthRight

    @GridlineWidthRight.setter
    def GridlineWidthRight(self, value):
        self.webbrowsercontrol.GridlineWidthRight = value

    @property
    def GridlineWidthTop(self):
        return self.webbrowsercontrol.GridlineWidthTop

    @GridlineWidthTop.setter
    def GridlineWidthTop(self, value):
        self.webbrowsercontrol.GridlineWidthTop = value

    @property
    def Height(self):
        return self.webbrowsercontrol.Height

    @Height.setter
    def Height(self, value):
        self.webbrowsercontrol.Height = value

    @property
    def HelpContextId(self):
        return self.webbrowsercontrol.HelpContextId

    @HelpContextId.setter
    def HelpContextId(self, value):
        self.webbrowsercontrol.HelpContextId = value

    @property
    def HorizontalAnchor(self):
        return self.webbrowsercontrol.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.webbrowsercontrol.HorizontalAnchor = value

    @property
    def Hyperlink(self):
        return self.webbrowsercontrol.Hyperlink

    @property
    def InSelection(self):
        return self.webbrowsercontrol.InSelection

    @InSelection.setter
    def InSelection(self, value):
        self.webbrowsercontrol.InSelection = value

    @property
    def Layout(self):
        return AcLayoutType(self.webbrowsercontrol.Layout)

    @property
    def LayoutID(self):
        return self.webbrowsercontrol.LayoutID

    @property
    def Left(self):
        return self.webbrowsercontrol.Left

    @Left.setter
    def Left(self, value):
        self.webbrowsercontrol.Left = value

    @property
    def LeftPadding(self):
        return self.webbrowsercontrol.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.webbrowsercontrol.LeftPadding = value

    @property
    def LocationURL(self):
        return self.webbrowsercontrol.LocationURL

    @property
    def Name(self):
        return self.webbrowsercontrol.Name

    @Name.setter
    def Name(self, value):
        self.webbrowsercontrol.Name = value

    @property
    def Object(self):
        return self.webbrowsercontrol.Object

    @property
    def OldValue(self):
        return self.webbrowsercontrol.OldValue

    @property
    def OnBeforeNavigate(self):
        return self.webbrowsercontrol.OnBeforeNavigate

    @OnBeforeNavigate.setter
    def OnBeforeNavigate(self, value):
        self.webbrowsercontrol.OnBeforeNavigate = value

    @property
    def OnDocumentComplete(self):
        return self.webbrowsercontrol.OnDocumentComplete

    @OnDocumentComplete.setter
    def OnDocumentComplete(self, value):
        self.webbrowsercontrol.OnDocumentComplete = value

    @property
    def OnKeyDown(self):
        return self.webbrowsercontrol.OnKeyDown

    @OnKeyDown.setter
    def OnKeyDown(self, value):
        self.webbrowsercontrol.OnKeyDown = value

    @property
    def OnKeyPress(self):
        return self.webbrowsercontrol.OnKeyPress

    @OnKeyPress.setter
    def OnKeyPress(self, value):
        self.webbrowsercontrol.OnKeyPress = value

    @property
    def OnKeyUp(self):
        return self.webbrowsercontrol.OnKeyUp

    @OnKeyUp.setter
    def OnKeyUp(self, value):
        self.webbrowsercontrol.OnKeyUp = value

    @property
    def OnMouseDown(self):
        return self.webbrowsercontrol.OnMouseDown

    @OnMouseDown.setter
    def OnMouseDown(self, value):
        self.webbrowsercontrol.OnMouseDown = value

    @property
    def OnMouseMove(self):
        return self.webbrowsercontrol.OnMouseMove

    @OnMouseMove.setter
    def OnMouseMove(self, value):
        self.webbrowsercontrol.OnMouseMove = value

    @property
    def OnMouseUp(self):
        return self.webbrowsercontrol.OnMouseUp

    @OnMouseUp.setter
    def OnMouseUp(self, value):
        self.webbrowsercontrol.OnMouseUp = value

    @property
    def OnNavigateError(self):
        return self.webbrowsercontrol.OnNavigateError

    @OnNavigateError.setter
    def OnNavigateError(self, value):
        self.webbrowsercontrol.OnNavigateError = value

    @property
    def OnProgressChange(self):
        return self.webbrowsercontrol.OnProgressChange

    @OnProgressChange.setter
    def OnProgressChange(self, value):
        self.webbrowsercontrol.OnProgressChange = value

    @property
    def OnUpdated(self):
        return self.webbrowsercontrol.OnUpdated

    @OnUpdated.setter
    def OnUpdated(self, value):
        self.webbrowsercontrol.OnUpdated = value

    @property
    def Parent(self):
        return self.webbrowsercontrol.Parent

    @property
    def Progress(self):
        return self.webbrowsercontrol.Progress

    @property
    def Properties(self):
        return Properties(self.webbrowsercontrol.Properties)

    @property
    def ReadyState(self):
        return self.webbrowsercontrol.ReadyState

    @property
    def RightPadding(self):
        return self.webbrowsercontrol.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.webbrowsercontrol.RightPadding = value

    @property
    def ScrollBars(self):
        return self.webbrowsercontrol.ScrollBars

    @ScrollBars.setter
    def ScrollBars(self, value):
        self.webbrowsercontrol.ScrollBars = value

    @property
    def ScrollLeft(self):
        return self.webbrowsercontrol.ScrollLeft

    @ScrollLeft.setter
    def ScrollLeft(self, value):
        self.webbrowsercontrol.ScrollLeft = value

    @property
    def ScrollTop(self):
        return self.webbrowsercontrol.ScrollTop

    @ScrollTop.setter
    def ScrollTop(self, value):
        self.webbrowsercontrol.ScrollTop = value

    @property
    def Section(self):
        return self.webbrowsercontrol.Section

    @Section.setter
    def Section(self, value):
        self.webbrowsercontrol.Section = value

    @property
    def SpecialEffect(self):
        return self.webbrowsercontrol.SpecialEffect

    @SpecialEffect.setter
    def SpecialEffect(self, value):
        self.webbrowsercontrol.SpecialEffect = value

    @property
    def StatusBarText(self):
        return self.webbrowsercontrol.StatusBarText

    @StatusBarText.setter
    def StatusBarText(self, value):
        self.webbrowsercontrol.StatusBarText = value

    @property
    def TabIndex(self):
        return self.webbrowsercontrol.TabIndex

    @TabIndex.setter
    def TabIndex(self, value):
        self.webbrowsercontrol.TabIndex = value

    @property
    def TabStop(self):
        return self.webbrowsercontrol.TabStop

    @TabStop.setter
    def TabStop(self, value):
        self.webbrowsercontrol.TabStop = value

    @property
    def Tag(self):
        return self.webbrowsercontrol.Tag

    @Tag.setter
    def Tag(self, value):
        self.webbrowsercontrol.Tag = value

    @property
    def Top(self):
        return self.webbrowsercontrol.Top

    @Top.setter
    def Top(self, value):
        self.webbrowsercontrol.Top = value

    @property
    def TopPadding(self):
        return self.webbrowsercontrol.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.webbrowsercontrol.TopPadding = value

    @property
    def Transform(self):
        return self.webbrowsercontrol.Transform

    @Transform.setter
    def Transform(self, value):
        self.webbrowsercontrol.Transform = value

    @property
    def Value(self):
        return self.webbrowsercontrol.Value

    @Value.setter
    def Value(self, value):
        self.webbrowsercontrol.Value = value

    @property
    def VerticalAnchor(self):
        return self.webbrowsercontrol.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.webbrowsercontrol.VerticalAnchor = value

    @property
    def Visible(self):
        return self.webbrowsercontrol.Visible

    @Visible.setter
    def Visible(self, value):
        self.webbrowsercontrol.Visible = value

    @property
    def Width(self):
        return self.webbrowsercontrol.Width

    @Width.setter
    def Width(self, value):
        self.webbrowsercontrol.Width = value

    def Move(self, *args, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.webbrowsercontrol.Move(*args, **arguments)

    def Requery(self):
        self.webbrowsercontrol.Requery()

    def SetFocus(self):
        self.webbrowsercontrol.SetFocus()

    def SizeToFit(self):
        self.webbrowsercontrol.SizeToFit()

    def Undo(self):
        self.webbrowsercontrol.Undo()

class WebService:

    def __init__(self, webservice=None):
        self.webservice = webservice

    @property
    def Entities(self):
        return self.webservice.Entities

    @property
    def Name(self):
        return self.webservice.Name

    @Name.setter
    def Name(self, value):
        self.webservice.Name = value

    @property
    def Parent(self):
        return self.webservice.Parent

class WebServices:

    def __init__(self, webservices=None):
        self.webservices = webservices

    @property
    def Application(self):
        return self.webservices.Application

    @property
    def Count(self):
        return self.webservices.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.webservices.Item):
            return self.webservices.Item(*args, **arguments)
        else:
            return self.webservices.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.webservices.Parent

class WSParameter:

    def __init__(self, wsparameter=None):
        self.wsparameter = wsparameter

    @property
    def Name(self):
        return self.wsparameter.Name

    @Name.setter
    def Name(self, value):
        self.wsparameter.Name = value

    @property
    def Parent(self):
        return self.wsparameter.Parent

    @property
    def Type(self):
        return self.wsparameter.Type

class WSParameters:

    def __init__(self, wsparameters=None):
        self.wsparameters = wsparameters

    @property
    def Count(self):
        return self.wsparameters.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.wsparameters.Item):
            return self.wsparameters.Item(*args, **arguments)
        else:
            return self.wsparameters.GetItem(*args, **arguments)

    @property
    def Parent(self):
        return self.wsparameters.Parent
