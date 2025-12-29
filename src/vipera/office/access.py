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
        return self.accessobjectproperties.Item(*args, **arguments)

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

class AdditionalData:

    def __init__(self, additionaldata=None):
        self.additionaldata = additionaldata

    @property
    def Count(self):
        return self.additionaldata.Count

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.additionaldata.Item(*args, **arguments)

    @property
    def Name(self):
        return self.additionaldata.Name

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
        return self.alldatabasediagrams.Item(*args, **arguments)

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
        return self.allforms.Item(*args, **arguments)

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
        return self.allfunctions.Item(*args, **arguments)

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
        return self.allmodules.Item(*args, **arguments)

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
        return self.allqueries.Item(*args, **arguments)

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
        return self.allreports.Item(*args, **arguments)

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
        return self.allstoredprocedures.Item(*args, **arguments)

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
        return self.alltables.Item(*args, **arguments)

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
        return self.allviews.Item(*args, **arguments)

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
    def  AppTitle(self):
        return self.application. AppTitle

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

    def FileDialog(self, *args, dialogType=None):
        arguments = {"dialogType": dialogType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.FileDialog(*args, **arguments)

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

    @property
    def TempVars(self):
        return TempVar(self.application.TempVars)

    @property
    def UserControl(self):
        return self.application.UserControl

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

    @property
    def Application(self):
        return self.attachment.Application

    @property
    def AttachmentCount(self):
        return self.attachment.AttachmentCount

    @property
    def AutoLabel(self):
        return self.attachment.AutoLabel

    @property
    def BackColor(self):
        return self.attachment.BackColor

    @property
    def BackShade(self):
        return self.attachment.BackShade

    @property
    def BackStyle(self):
        return self.attachment.BackStyle

    @property
    def BackThemeColorIndex(self):
        return self.attachment.BackThemeColorIndex

    @property
    def BackTint(self):
        return self.attachment.BackTint

    @property
    def BorderColor(self):
        return self.attachment.BorderColor

    @property
    def BorderShade(self):
        return self.attachment.BorderShade

    @property
    def BorderStyle(self):
        return self.attachment.BorderStyle

    @property
    def BorderThemeColorIndex(self):
        return self.attachment.BorderThemeColorIndex

    @property
    def BorderTint(self):
        return self.attachment.BorderTint

    @property
    def BorderWidth(self):
        return self.attachment.BorderWidth

    @property
    def BottomPadding(self):
        return self.attachment.BottomPadding

    @property
    def ColumnHidden(self):
        return self.attachment.ColumnHidden

    @property
    def ColumnOrder(self):
        return self.attachment.ColumnOrder

    @property
    def ColumnWidth(self):
        return self.attachment.ColumnWidth

    @property
    def Controls(self):
        return Controls(self.attachment.Controls)

    @property
    def ControlSource(self):
        return self.attachment.ControlSource

    @property
    def ControlTipText(self):
        return self.attachment.ControlTipText

    @property
    def ControlType(self):
        return self.attachment.ControlType

    @property
    def CurrentAttachment(self):
        return self.attachment.CurrentAttachment

    @property
    def DefaultPicture(self):
        return self.attachment.DefaultPicture

    @property
    def DefaultPictureType(self):
        return self.attachment.DefaultPictureType

    @property
    def DisplayAs(self):
        return self.attachment.DisplayAs

    @property
    def DisplayWhen(self):
        return self.attachment.DisplayWhen

    @property
    def Enabled(self):
        return self.attachment.Enabled

    @property
    def EventProcPrefix(self):
        return self.attachment.EventProcPrefix

    def FileName(self, *args, var=None):
        arguments = {"var": var}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.attachment.FileName(*args, **arguments)

    def FileType(self, *args, var=None):
        arguments = {"var": var}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.attachment.FileType(*args, **arguments)

    def FileURL(self, *args, var=None):
        arguments = {"var": var}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.attachment.FileURL(*args, **arguments)

    @property
    def GridlineColor(self):
        return self.attachment.GridlineColor

    @property
    def GridlineShade(self):
        return self.attachment.GridlineShade

    @property
    def GridlineStyleBottom(self):
        return self.attachment.GridlineStyleBottom

    @property
    def GridlineStyleLeft(self):
        return self.attachment.GridlineStyleLeft

    @property
    def GridlineStyleRight(self):
        return self.attachment.GridlineStyleRight

    @property
    def GridlineStyleTop(self):
        return self.attachment.GridlineStyleTop

    @property
    def GridlineThemeColorIndex(self):
        return self.attachment.GridlineThemeColorIndex

    @property
    def GridlineTint(self):
        return self.attachment.GridlineTint

    @property
    def GridlineWidthBottom(self):
        return self.attachment.GridlineWidthBottom

    @property
    def GridlineWidthLeft(self):
        return self.attachment.GridlineWidthLeft

    @property
    def GridlineWidthRight(self):
        return self.attachment.GridlineWidthRight

    @property
    def GridlineWidthTop(self):
        return self.attachment.GridlineWidthTop

    @property
    def Height(self):
        return self.attachment.Height

    @property
    def HelpContextId(self):
        return self.attachment.HelpContextId

    @property
    def HorizontalAnchor(self):
        return self.attachment.HorizontalAnchor

    @property
    def InSelection(self):
        return self.attachment.InSelection

    @property
    def IsVisible(self):
        return self.attachment.IsVisible

    @property
    def LabelAlign(self):
        return self.attachment.LabelAlign

    @property
    def LabelX(self):
        return self.attachment.LabelX

    @property
    def LabelY(self):
        return self.attachment.LabelY

    @property
    def Layout(self):
        return AcLayoutType(self.attachment.Layout)

    @property
    def LayoutID(self):
        return self.attachment.LayoutID

    @property
    def Left(self):
        return self.attachment.Left

    @property
    def LeftPadding(self):
        return self.attachment.LeftPadding

    @property
    def Locked(self):
        return self.attachment.Locked

    @property
    def Name(self):
        return self.attachment.Name

    @property
    def OldBorderStyle(self):
        return self.attachment.OldBorderStyle

    @property
    def OldValue(self):
        return self.attachment.OldValue

    @property
    def OnAttachmentCurrent(self):
        return self.attachment.OnAttachmentCurrent

    @property
    def OnChange(self):
        return self.attachment.OnChange

    @property
    def OnClick(self):
        return self.attachment.OnClick

    @property
    def OnDblClick(self):
        return self.attachment.OnDblClick

    @property
    def OnDirty(self):
        return self.attachment.OnDirty

    @property
    def OnEnter(self):
        return self.attachment.OnEnter

    @property
    def OnExit(self):
        return self.attachment.OnExit

    @property
    def OnGotFocus(self):
        return self.attachment.OnGotFocus

    @property
    def OnKeyDown(self):
        return self.attachment.OnKeyDown

    @property
    def OnKeyPress(self):
        return self.attachment.OnKeyPress

    @property
    def OnKeyUp(self):
        return self.attachment.OnKeyUp

    @property
    def OnLostFocus(self):
        return self.attachment.OnLostFocus

    @property
    def OnMouseDown(self):
        return self.attachment.OnMouseDown

    @property
    def OnMouseMove(self):
        return self.attachment.OnMouseMove

    @property
    def OnMouseUp(self):
        return self.attachment.OnMouseUp

    @property
    def Parent(self):
        return self.attachment.Parent

    @property
    def PictureAlignment(self):
        return self.attachment.PictureAlignment

    @property
    def PictureSizeMode(self):
        return self.attachment.PictureSizeMode

    @property
    def PictureTiling(self):
        return self.attachment.PictureTiling

    @property
    def Properties(self):
        return Properties(self.attachment.Properties)

    @property
    def RightPadding(self):
        return self.attachment.RightPadding

    @property
    def Section(self):
        return self.attachment.Section

    @property
    def ShortcutMenuBar(self):
        return self.attachment.ShortcutMenuBar

    @property
    def SpecialEffect(self):
        return self.attachment.SpecialEffect

    @property
    def StatusBarText(self):
        return self.attachment.StatusBarText

    @property
    def TabIndex(self):
        return self.attachment.TabIndex

    @property
    def TabStop(self):
        return self.attachment.TabStop

    @property
    def Tag(self):
        return self.attachment.Tag

    @property
    def Top(self):
        return self.attachment.Top

    @property
    def TopPadding(self):
        return self.attachment.TopPadding

    @property
    def VerticalAnchor(self):
        return self.attachment.VerticalAnchor

    @property
    def Visible(self):
        return self.attachment.Visible

    @Visible.setter
    def Visible(self, value):
        self.attachment.Visible = value

    @property
    def Width(self):
        return self.attachment.Width

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

    @property
    def AddColon(self):
        return self.boundobjectframe.AddColon

    @property
    def Application(self):
        return self.boundobjectframe.Application

    @property
    def AutoActivate(self):
        return self.boundobjectframe.AutoActivate

    @property
    def AutoLabel(self):
        return self.boundobjectframe.AutoLabel

    @property
    def BackColor(self):
        return self.boundobjectframe.BackColor

    @property
    def BackShade(self):
        return self.boundobjectframe.BackShade

    @property
    def BackStyle(self):
        return self.boundobjectframe.BackStyle

    @property
    def BackThemeColorIndex(self):
        return self.boundobjectframe.BackThemeColorIndex

    @property
    def BackTint(self):
        return self.boundobjectframe.BackTint

    @property
    def BorderColor(self):
        return self.boundobjectframe.BorderColor

    @property
    def BorderShade(self):
        return self.boundobjectframe.BorderShade

    @property
    def BorderStyle(self):
        return self.boundobjectframe.BorderStyle

    @property
    def BorderThemeColorIndex(self):
        return self.boundobjectframe.BorderThemeColorIndex

    @property
    def BorderTint(self):
        return self.boundobjectframe.BorderTint

    @property
    def BorderWidth(self):
        return self.boundobjectframe.BorderWidth

    @property
    def BottomPadding(self):
        return self.boundobjectframe.BottomPadding

    @property
    def Class(self):
        return self.boundobjectframe.Class

    @property
    def ColumnHidden(self):
        return self.boundobjectframe.ColumnHidden

    @property
    def ColumnOrder(self):
        return self.boundobjectframe.ColumnOrder

    @property
    def ColumnWidth(self):
        return self.boundobjectframe.ColumnWidth

    @property
    def Controls(self):
        return Controls(self.boundobjectframe.Controls)

    @property
    def ControlSource(self):
        return self.boundobjectframe.ControlSource

    @property
    def ControlTipText(self):
        return self.boundobjectframe.ControlTipText

    @property
    def ControlType(self):
        return self.boundobjectframe.ControlType

    @property
    def DisplayType(self):
        return self.boundobjectframe.DisplayType

    @property
    def DisplayWhen(self):
        return self.boundobjectframe.DisplayWhen

    @property
    def Enabled(self):
        return self.boundobjectframe.Enabled

    @property
    def EventProcPrefix(self):
        return self.boundobjectframe.EventProcPrefix

    @property
    def GridlineColor(self):
        return self.boundobjectframe.GridlineColor

    @property
    def GridlineShade(self):
        return self.boundobjectframe.GridlineShade

    @property
    def GridlineStyleBottom(self):
        return self.boundobjectframe.GridlineStyleBottom

    @property
    def GridlineStyleLeft(self):
        return self.boundobjectframe.GridlineStyleLeft

    @property
    def GridlineStyleRight(self):
        return self.boundobjectframe.GridlineStyleRight

    @property
    def GridlineStyleTop(self):
        return self.boundobjectframe.GridlineStyleTop

    @property
    def GridlineThemeColorIndex(self):
        return self.boundobjectframe.GridlineThemeColorIndex

    @property
    def GridlineTint(self):
        return self.boundobjectframe.GridlineTint

    @property
    def GridlineWidthBottom(self):
        return self.boundobjectframe.GridlineWidthBottom

    @property
    def GridlineWidthLeft(self):
        return self.boundobjectframe.GridlineWidthLeft

    @property
    def GridlineWidthRight(self):
        return self.boundobjectframe.GridlineWidthRight

    @property
    def GridlineWidthTop(self):
        return self.boundobjectframe.GridlineWidthTop

    @property
    def Height(self):
        return self.boundobjectframe.Height

    @property
    def HelpContextId(self):
        return self.boundobjectframe.HelpContextId

    @property
    def HorizontalAnchor(self):
        return self.boundobjectframe.HorizontalAnchor

    @property
    def InSelection(self):
        return self.boundobjectframe.InSelection

    @property
    def IsVisible(self):
        return self.boundobjectframe.IsVisible

    @property
    def LabelAlign(self):
        return self.boundobjectframe.LabelAlign

    @property
    def LabelX(self):
        return self.boundobjectframe.LabelX

    @property
    def LabelY(self):
        return self.boundobjectframe.LabelY

    @property
    def Layout(self):
        return AcLayoutType(self.boundobjectframe.Layout)

    @property
    def LayoutID(self):
        return self.boundobjectframe.LayoutID

    @property
    def Left(self):
        return self.boundobjectframe.Left

    @property
    def LeftPadding(self):
        return self.boundobjectframe.LeftPadding

    @property
    def Locked(self):
        return self.boundobjectframe.Locked

    @property
    def Name(self):
        return self.boundobjectframe.Name

    @property
    def Object(self):
        return self.boundobjectframe.Object

    @property
    def ObjectPalette(self):
        return self.boundobjectframe.ObjectPalette

    def ObjectVerbs(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.boundobjectframe.ObjectVerbs(*args, **arguments)

    @property
    def ObjectVerbsCount(self):
        return self.boundobjectframe.ObjectVerbsCount

    @property
    def OldBorderStyle(self):
        return self.boundobjectframe.OldBorderStyle

    @property
    def OldValue(self):
        return self.boundobjectframe.OldValue

    @property
    def OLEType(self):
        return self.boundobjectframe.OLEType

    @property
    def OLETypeAllowed(self):
        return self.boundobjectframe.OLETypeAllowed

    @property
    def OnClick(self):
        return self.boundobjectframe.OnClick

    @property
    def OnDblClick(self):
        return self.boundobjectframe.OnDblClick

    @property
    def OnEnter(self):
        return self.boundobjectframe.OnEnter

    @property
    def OnExit(self):
        return self.boundobjectframe.OnExit

    @property
    def OnGotFocus(self):
        return self.boundobjectframe.OnGotFocus

    @property
    def OnKeyDown(self):
        return self.boundobjectframe.OnKeyDown

    @property
    def OnKeyPress(self):
        return self.boundobjectframe.OnKeyPress

    @property
    def OnKeyUp(self):
        return self.boundobjectframe.OnKeyUp

    @property
    def OnLostFocus(self):
        return self.boundobjectframe.OnLostFocus

    @property
    def OnMouseDown(self):
        return self.boundobjectframe.OnMouseDown

    @property
    def OnMouseMove(self):
        return self.boundobjectframe.OnMouseMove

    @property
    def OnMouseUp(self):
        return self.boundobjectframe.OnMouseUp

    @property
    def OnUpdated(self):
        return self.boundobjectframe.OnUpdated

    @property
    def Parent(self):
        return self.boundobjectframe.Parent

    @property
    def Properties(self):
        return Properties(self.boundobjectframe.Properties)

    @property
    def RightPadding(self):
        return self.boundobjectframe.RightPadding

    @property
    def Scaling(self):
        return self.boundobjectframe.Scaling

    @property
    def Section(self):
        return self.boundobjectframe.Section

    @property
    def ShortcutMenuBar(self):
        return self.boundobjectframe.ShortcutMenuBar

    @property
    def SizeMode(self):
        return self.boundobjectframe.SizeMode

    @property
    def SourceDoc(self):
        return self.boundobjectframe.SourceDoc

    @property
    def SourceItem(self):
        return self.boundobjectframe.SourceItem

    @property
    def SpecialEffect(self):
        return self.boundobjectframe.SpecialEffect

    @property
    def StatusBarText(self):
        return self.boundobjectframe.StatusBarText

    @property
    def TabIndex(self):
        return self.boundobjectframe.TabIndex

    @property
    def TabStop(self):
        return self.boundobjectframe.TabStop

    @property
    def Tag(self):
        return self.boundobjectframe.Tag

    @property
    def Top(self):
        return self.boundobjectframe.Top

    @property
    def TopPadding(self):
        return self.boundobjectframe.TopPadding

    @property
    def UpdateOptions(self):
        return self.boundobjectframe.UpdateOptions

    @property
    def Value(self):
        return self.boundobjectframe.Value

    @property
    def VarOleObject(self):
        return self.boundobjectframe.VarOleObject

    @property
    def Verb(self):
        return self.boundobjectframe.Verb

    @property
    def VerticalAnchor(self):
        return self.boundobjectframe.VerticalAnchor

    @property
    def Visible(self):
        return self.boundobjectframe.Visible

    @Visible.setter
    def Visible(self, value):
        self.boundobjectframe.Visible = value

    @property
    def Width(self):
        return self.boundobjectframe.Width

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

    @property
    def HasLegend(self):
        return self.chart.HasLegend

    @property
    def HasSubtitle(self):
        return self.chart.HasSubtitle

    @property
    def HasTitle(self):
        return self.chart.HasTitle

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

    @property
    def DisplayBoxWhiskerMeanMarker(self):
        return self.chartseries.DisplayBoxWhiskerMeanMarker

    @property
    def DisplayDataLabel(self):
        return self.chartseries.DisplayDataLabel

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

    @property
    def ShowWaterfallConnectorLines(self):
        return self.chartseries.ShowWaterfallConnectorLines

    @property
    def ShowWaterfallTotal(self):
        return self.chartseries.ShowWaterfallTotal

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

    @property
    def Application(self):
        return self.checkbox.Application

    @property
    def AutoLabel(self):
        return self.checkbox.AutoLabel

    @property
    def BorderColor(self):
        return self.checkbox.BorderColor

    @property
    def BorderShade(self):
        return self.checkbox.BorderShade

    @property
    def BorderStyle(self):
        return self.checkbox.BorderStyle

    @property
    def BorderThemeColorIndex(self):
        return self.checkbox.BorderThemeColorIndex

    @property
    def BorderTint(self):
        return self.checkbox.BorderTint

    @property
    def BorderWidth(self):
        return self.checkbox.BorderWidth

    @property
    def BottomPadding(self):
        return self.checkbox.BottomPadding

    @property
    def ColumnHidden(self):
        return self.checkbox.ColumnHidden

    @property
    def ColumnOrder(self):
        return self.checkbox.ColumnOrder

    @property
    def ColumnWidth(self):
        return self.checkbox.ColumnWidth

    @property
    def Controls(self):
        return Controls(self.checkbox.Controls)

    @property
    def ControlSource(self):
        return self.checkbox.ControlSource

    @property
    def ControlTipText(self):
        return self.checkbox.ControlTipText

    @property
    def ControlType(self):
        return self.checkbox.ControlType

    @property
    def DefaultValue(self):
        return self.checkbox.DefaultValue

    @property
    def DisplayWhen(self):
        return self.checkbox.DisplayWhen

    @property
    def Enabled(self):
        return self.checkbox.Enabled

    @property
    def EventProcPrefix(self):
        return self.checkbox.EventProcPrefix

    @property
    def GridlineColor(self):
        return self.checkbox.GridlineColor

    @property
    def GridlineShade(self):
        return self.checkbox.GridlineShade

    @property
    def GridlineStyleBottom(self):
        return self.checkbox.GridlineStyleBottom

    @property
    def GridlineStyleLeft(self):
        return self.checkbox.GridlineStyleLeft

    @property
    def GridlineStyleRight(self):
        return self.checkbox.GridlineStyleRight

    @property
    def GridlineStyleTop(self):
        return self.checkbox.GridlineStyleTop

    @property
    def GridlineThemeColorIndex(self):
        return self.checkbox.GridlineThemeColorIndex

    @property
    def GridlineTint(self):
        return self.checkbox.GridlineTint

    @property
    def GridlineWidthBottom(self):
        return self.checkbox.GridlineWidthBottom

    @property
    def GridlineWidthLeft(self):
        return self.checkbox.GridlineWidthLeft

    @property
    def GridlineWidthRight(self):
        return self.checkbox.GridlineWidthRight

    @property
    def GridlineWidthTop(self):
        return self.checkbox.GridlineWidthTop

    @property
    def Height(self):
        return self.checkbox.Height

    @property
    def HelpContextId(self):
        return self.checkbox.HelpContextId

    @property
    def HideDuplicates(self):
        return self.checkbox.HideDuplicates

    @property
    def HorizontalAnchor(self):
        return self.checkbox.HorizontalAnchor

    @property
    def InSelection(self):
        return self.checkbox.InSelection

    @property
    def IsVisible(self):
        return self.checkbox.IsVisible

    @property
    def LabelAlign(self):
        return self.checkbox.LabelAlign

    @property
    def LabelX(self):
        return self.checkbox.LabelX

    @property
    def LabelY(self):
        return self.checkbox.LabelY

    @property
    def Layout(self):
        return AcLayoutType(self.checkbox.Layout)

    @property
    def LayoutID(self):
        return self.checkbox.LayoutID

    @property
    def Left(self):
        return self.checkbox.Left

    @property
    def LeftPadding(self):
        return self.checkbox.LeftPadding

    @property
    def Locked(self):
        return self.checkbox.Locked

    @property
    def Name(self):
        return self.checkbox.Name

    @property
    def OldBorderStyle(self):
        return self.checkbox.OldBorderStyle

    @property
    def OldValue(self):
        return self.checkbox.OldValue

    @property
    def OnClick(self):
        return self.checkbox.OnClick

    @property
    def OnDblClick(self):
        return self.checkbox.OnDblClick

    @property
    def OnEnter(self):
        return self.checkbox.OnEnter

    @property
    def OnExit(self):
        return self.checkbox.OnExit

    @property
    def OnGotFocus(self):
        return self.checkbox.OnGotFocus

    @property
    def OnKeyDown(self):
        return self.checkbox.OnKeyDown

    @property
    def OnKeyPress(self):
        return self.checkbox.OnKeyPress

    @property
    def OnKeyUp(self):
        return self.checkbox.OnKeyUp

    @property
    def OnLostFocus(self):
        return self.checkbox.OnLostFocus

    @property
    def OnMouseDown(self):
        return self.checkbox.OnMouseDown

    @property
    def OnMouseMove(self):
        return self.checkbox.OnMouseMove

    @property
    def OnMouseUp(self):
        return self.checkbox.OnMouseUp

    @property
    def OptionValue(self):
        return self.checkbox.OptionValue

    @property
    def Parent(self):
        return self.checkbox.Parent

    @property
    def Properties(self):
        return Properties(self.checkbox.Properties)

    @property
    def ReadingOrder(self):
        return self.checkbox.ReadingOrder

    @property
    def RightPadding(self):
        return self.checkbox.RightPadding

    @property
    def Section(self):
        return self.checkbox.Section

    @property
    def ShortcutMenuBar(self):
        return self.checkbox.ShortcutMenuBar

    @property
    def SpecialEffect(self):
        return self.checkbox.SpecialEffect

    @property
    def StatusBarText(self):
        return self.checkbox.StatusBarText

    @property
    def TabIndex(self):
        return self.checkbox.TabIndex

    @property
    def TabStop(self):
        return self.checkbox.TabStop

    @property
    def Tag(self):
        return self.checkbox.Tag

    @property
    def Top(self):
        return self.checkbox.Top

    @property
    def TopPadding(self):
        return self.checkbox.TopPadding

    @property
    def TripleState(self):
        return self.checkbox.TripleState

    @property
    def ValidationRule(self):
        return self.checkbox.ValidationRule

    @property
    def ValidationText(self):
        return self.checkbox.ValidationText

    @property
    def Value(self):
        return self.checkbox.Value

    @property
    def VerticalAnchor(self):
        return self.checkbox.VerticalAnchor

    @property
    def Visible(self):
        return self.checkbox.Visible

    @Visible.setter
    def Visible(self, value):
        self.checkbox.Visible = value

    @property
    def Width(self):
        return self.checkbox.Width

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

    @property
    def AllowAutoCorrect(self):
        return self.combobox.AllowAutoCorrect

    @property
    def AllowValueListEdits(self):
        return self.combobox.AllowValueListEdits

    @property
    def Application(self):
        return self.combobox.Application

    @property
    def AutoExpand(self):
        return self.combobox.AutoExpand

    @property
    def AutoLabel(self):
        return self.combobox.AutoLabel

    @property
    def BackColor(self):
        return self.combobox.BackColor

    @property
    def BackShade(self):
        return self.combobox.BackShade

    @property
    def BackStyle(self):
        return self.combobox.BackStyle

    @property
    def BackThemeColorIndex(self):
        return self.combobox.BackThemeColorIndex

    @property
    def BackTint(self):
        return self.combobox.BackTint

    @property
    def BorderColor(self):
        return self.combobox.BorderColor

    @property
    def BorderShade(self):
        return self.combobox.BorderShade

    @property
    def BorderStyle(self):
        return self.combobox.BorderStyle

    @property
    def BorderThemeColorIndex(self):
        return self.combobox.BorderThemeColorIndex

    @property
    def BorderTint(self):
        return self.combobox.BorderTint

    @property
    def BorderWidth(self):
        return self.combobox.BorderWidth

    @property
    def BottomMargin(self):
        return self.combobox.BottomMargin

    @property
    def BottomPadding(self):
        return self.combobox.BottomPadding

    @property
    def BoundColumn(self):
        return self.combobox.BoundColumn

    @property
    def CanGrow(self):
        return self.combobox.CanGrow

    @property
    def CanShrink(self):
        return self.combobox.CanShrink

    def Column(self, *args, Index=None, Row=None):
        arguments = {"Index": Index, "Row": Row}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.combobox.Column(*args, **arguments)

    @property
    def ColumnCount(self):
        return self.combobox.ColumnCount

    @property
    def ColumnHeads(self):
        return self.combobox.ColumnHeads

    @property
    def ColumnHidden(self):
        return self.combobox.ColumnHidden

    @property
    def ColumnOrder(self):
        return self.combobox.ColumnOrder

    @property
    def ColumnWidth(self):
        return self.combobox.ColumnWidth

    @property
    def ColumnWidths(self):
        return self.combobox.ColumnWidths

    @property
    def Controls(self):
        return Controls(self.combobox.Controls)

    @property
    def ControlSource(self):
        return self.combobox.ControlSource

    @property
    def ControlTipText(self):
        return self.combobox.ControlTipText

    @property
    def ControlType(self):
        return self.combobox.ControlType

    @property
    def DecimalPlaces(self):
        return self.combobox.DecimalPlaces

    @property
    def DefaultValue(self):
        return self.combobox.DefaultValue

    @property
    def DisplayAsHyperlink(self):
        return self.combobox.DisplayAsHyperlink

    @property
    def DisplayWhen(self):
        return self.combobox.DisplayWhen

    @property
    def Enabled(self):
        return self.combobox.Enabled

    @property
    def EventProcPrefix(self):
        return self.combobox.EventProcPrefix

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

    @property
    def ForeColor(self):
        return self.combobox.ForeColor

    @property
    def ForeShade(self):
        return self.combobox.ForeShade

    @property
    def ForeThemeColorIndex(self):
        return self.combobox.ForeThemeColorIndex

    @property
    def ForeTint(self):
        return self.combobox.ForeTint

    @property
    def Format(self):
        return self.combobox.Format

    @property
    def FormatConditions(self):
        return self.combobox.FormatConditions

    @property
    def GridlineColor(self):
        return self.combobox.GridlineColor

    @property
    def GridlineShade(self):
        return self.combobox.GridlineShade

    @property
    def GridlineStyleBottom(self):
        return self.combobox.GridlineStyleBottom

    @property
    def GridlineStyleLeft(self):
        return self.combobox.GridlineStyleLeft

    @property
    def GridlineStyleRight(self):
        return self.combobox.GridlineStyleRight

    @property
    def GridlineStyleTop(self):
        return self.combobox.GridlineStyleTop

    @property
    def GridlineThemeColorIndex(self):
        return self.combobox.GridlineThemeColorIndex

    @property
    def GridlineTint(self):
        return self.combobox.GridlineTint

    @property
    def GridlineWidthBottom(self):
        return self.combobox.GridlineWidthBottom

    @property
    def GridlineWidthLeft(self):
        return self.combobox.GridlineWidthLeft

    @property
    def GridlineWidthRight(self):
        return self.combobox.GridlineWidthRight

    @property
    def GridlineWidthTop(self):
        return self.combobox.GridlineWidthTop

    @property
    def Height(self):
        return self.combobox.Height

    @property
    def HelpContextId(self):
        return self.combobox.HelpContextId

    @property
    def HideDuplicates(self):
        return self.combobox.HideDuplicates

    @property
    def HorizontalAnchor(self):
        return self.combobox.HorizontalAnchor

    @property
    def IMEHold(self):
        return self.combobox.IMEHold

    @property
    def InheritValueList(self):
        return self.combobox.InheritValueList

    @property
    def InputMask(self):
        return self.combobox.InputMask

    @property
    def InSelection(self):
        return self.combobox.InSelection

    @property
    def IsHyperlink(self):
        return self.combobox.IsHyperlink

    @property
    def IsVisible(self):
        return self.combobox.IsVisible

    def ItemData(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.combobox.ItemData(*args, **arguments)

    @property
    def ItemsSelected(self):
        return self.combobox.ItemsSelected

    @property
    def LabelAlign(self):
        return self.combobox.LabelAlign

    @property
    def LabelX(self):
        return self.combobox.LabelX

    @property
    def LabelY(self):
        return self.combobox.LabelY

    @property
    def Layout(self):
        return AcLayoutType(self.combobox.Layout)

    @property
    def LayoutID(self):
        return self.combobox.LayoutID

    @property
    def Left(self):
        return self.combobox.Left

    @property
    def LeftMargin(self):
        return self.combobox.LeftMargin

    @property
    def LeftPadding(self):
        return self.combobox.LeftPadding

    @property
    def LimitToList(self):
        return self.combobox.LimitToList

    @property
    def ListCount(self):
        return self.combobox.ListCount

    @property
    def ListIndex(self):
        return self.combobox.ListIndex

    @property
    def ListItemsEditForm(self):
        return self.combobox.ListItemsEditForm

    @property
    def ListRows(self):
        return self.combobox.ListRows

    @property
    def ListWidth(self):
        return self.combobox.ListWidth

    @property
    def Locked(self):
        return self.combobox.Locked

    @property
    def Name(self):
        return self.combobox.Name

    @property
    def OldBorderStyle(self):
        return self.combobox.OldBorderStyle

    @property
    def OldValue(self):
        return self.combobox.OldValue

    @property
    def OnChange(self):
        return self.combobox.OnChange

    @property
    def OnClick(self):
        return self.combobox.OnClick

    @property
    def OnDblClick(self):
        return self.combobox.OnDblClick

    @property
    def OnDirty(self):
        return self.combobox.OnDirty

    @property
    def OnEnter(self):
        return self.combobox.OnEnter

    @property
    def OnExit(self):
        return self.combobox.OnExit

    @property
    def OnGotFocus(self):
        return self.combobox.OnGotFocus

    @property
    def OnKeyDown(self):
        return self.combobox.OnKeyDown

    @property
    def OnKeyPress(self):
        return self.combobox.OnKeyPress

    @property
    def OnKeyUp(self):
        return self.combobox.OnKeyUp

    @property
    def OnLostFocus(self):
        return self.combobox.OnLostFocus

    @property
    def OnMouseDown(self):
        return self.combobox.OnMouseDown

    @property
    def OnMouseMove(self):
        return self.combobox.OnMouseMove

    @property
    def OnMouseUp(self):
        return self.combobox.OnMouseUp

    @property
    def OnNotInList(self):
        return self.combobox.OnNotInList

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

    @property
    def Recordset(self):
        return self.combobox.Recordset

    @Recordset.setter
    def Recordset(self, value):
        self.combobox.Recordset = value

    @property
    def RightMargin(self):
        return self.combobox.RightMargin

    @property
    def RightPadding(self):
        return self.combobox.RightPadding

    @property
    def RowSource(self):
        return self.combobox.RowSource

    @property
    def RowSourceType(self):
        return self.combobox.RowSourceType

    @property
    def ScrollBarAlign(self):
        return self.combobox.ScrollBarAlign

    @property
    def Section(self):
        return self.combobox.Section

    def Selected(self, *args, lRow=None):
        arguments = {"lRow": lRow}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.combobox.Selected(*args, **arguments)

    @property
    def SelLength(self):
        return self.combobox.SelLength

    @property
    def SelStart(self):
        return self.combobox.SelStart

    @property
    def SelText(self):
        return self.combobox.SelText

    @property
    def SeparatorCharacters(self):
        return self.combobox.SeparatorCharacters

    @property
    def ShortcutMenuBar(self):
        return self.combobox.ShortcutMenuBar

    @property
    def ShowOnlyRowSourceValues(self):
        return self.combobox.ShowOnlyRowSourceValues

    @property
    def SmartTags(self):
        return SmartTags(self.combobox.SmartTags)

    @property
    def SpecialEffect(self):
        return self.combobox.SpecialEffect

    @property
    def StatusBarText(self):
        return self.combobox.StatusBarText

    @property
    def TabIndex(self):
        return self.combobox.TabIndex

    @property
    def TabStop(self):
        return self.combobox.TabStop

    @property
    def Tag(self):
        return self.combobox.Tag

    @property
    def Text(self):
        return self.combobox.Text

    @property
    def TextAlign(self):
        return self.combobox.TextAlign

    @property
    def ThemeFontIndex(self):
        return self.combobox.ThemeFontIndex

    @property
    def Top(self):
        return self.combobox.Top

    @property
    def TopMargin(self):
        return self.combobox.TopMargin

    @property
    def TopPadding(self):
        return self.combobox.TopPadding

    @property
    def ValidationRule(self):
        return self.combobox.ValidationRule

    @property
    def ValidationText(self):
        return self.combobox.ValidationText

    @property
    def Value(self):
        return self.combobox.Value

    @property
    def VerticalAnchor(self):
        return self.combobox.VerticalAnchor

    @property
    def Visible(self):
        return self.combobox.Visible

    @Visible.setter
    def Visible(self, value):
        self.combobox.Visible = value

    @property
    def Width(self):
        return self.combobox.Width

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

    @property
    def Alignment(self):
        return self.commandbutton.Alignment

    @property
    def Application(self):
        return self.commandbutton.Application

    @property
    def AutoLabel(self):
        return self.commandbutton.AutoLabel

    @property
    def AutoRepeat(self):
        return self.commandbutton.AutoRepeat

    @property
    def BackColor(self):
        return self.commandbutton.BackColor

    @property
    def BackShade(self):
        return self.commandbutton.BackShade

    @property
    def BackStyle(self):
        return self.commandbutton.BackStyle

    @property
    def BackThemeColorIndex(self):
        return self.commandbutton.BackThemeColorIndex

    @property
    def BackTint(self):
        return self.commandbutton.BackTint

    @property
    def Bevel(self):
        return self.commandbutton.Bevel

    @property
    def BorderColor(self):
        return self.commandbutton.BorderColor

    @property
    def BorderShade(self):
        return self.commandbutton.BorderShade

    @property
    def BorderStyle(self):
        return self.commandbutton.BorderStyle

    @property
    def BorderThemeColorIndex(self):
        return self.commandbutton.BorderThemeColorIndex

    @property
    def BorderTint(self):
        return self.commandbutton.BorderTint

    @property
    def BorderWidth(self):
        return self.commandbutton.BorderWidth

    @property
    def BottomPadding(self):
        return self.commandbutton.BottomPadding

    @property
    def Cancel(self):
        return self.commandbutton.Cancel

    @property
    def Caption(self):
        return self.commandbutton.Caption

    @property
    def Controls(self):
        return Controls(self.commandbutton.Controls)

    @property
    def ControlTipText(self):
        return self.commandbutton.ControlTipText

    @property
    def ControlType(self):
        return self.commandbutton.ControlType

    @property
    def CursorOnHover(self):
        return self.commandbutton.CursorOnHover

    @property
    def Default(self):
        return self.commandbutton.Default

    @property
    def DisplayWhen(self):
        return self.commandbutton.DisplayWhen

    @property
    def Enabled(self):
        return self.commandbutton.Enabled

    @property
    def EventProcPrefix(self):
        return self.commandbutton.EventProcPrefix

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

    @property
    def ForeColor(self):
        return self.commandbutton.ForeColor

    @property
    def ForeShade(self):
        return self.commandbutton.ForeShade

    @property
    def ForeThemeColorIndex(self):
        return self.commandbutton.ForeThemeColorIndex

    @property
    def ForeTint(self):
        return self.commandbutton.ForeTint

    @property
    def Glow(self):
        return self.commandbutton.Glow

    @property
    def Gradient(self):
        return self.commandbutton.Gradient

    @property
    def GridlineColor(self):
        return self.commandbutton.GridlineColor

    @property
    def GridlineShade(self):
        return self.commandbutton.GridlineShade

    @property
    def GridlineStyleBottom(self):
        return self.commandbutton.GridlineStyleBottom

    @property
    def GridlineStyleLeft(self):
        return self.commandbutton.GridlineStyleLeft

    @property
    def GridlineStyleRight(self):
        return self.commandbutton.GridlineStyleRight

    @property
    def GridlineStyleTop(self):
        return self.commandbutton.GridlineStyleTop

    @property
    def GridlineThemeColorIndex(self):
        return self.commandbutton.GridlineThemeColorIndex

    @property
    def GridlineTint(self):
        return self.commandbutton.GridlineTint

    @property
    def GridlineWidthBottom(self):
        return self.commandbutton.GridlineWidthBottom

    @property
    def GridlineWidthLeft(self):
        return self.commandbutton.GridlineWidthLeft

    @property
    def GridlineWidthRight(self):
        return self.commandbutton.GridlineWidthRight

    @property
    def GridlineWidthTop(self):
        return self.commandbutton.GridlineWidthTop

    @property
    def Height(self):
        return self.commandbutton.Height

    @property
    def HelpContextId(self):
        return self.commandbutton.HelpContextId

    @property
    def HorizontalAnchor(self):
        return self.commandbutton.HorizontalAnchor

    @property
    def HoverColor(self):
        return self.commandbutton.HoverColor

    @property
    def HoverForeColor(self):
        return self.commandbutton.HoverForeColor

    @property
    def HoverForeShade(self):
        return self.commandbutton.HoverForeShade

    @property
    def HoverForeThemeColorIndex(self):
        return self.commandbutton.HoverForeThemeColorIndex

    @property
    def HoverForeTint(self):
        return self.commandbutton.HoverForeTint

    @property
    def HoverShade(self):
        return self.commandbutton.HoverShade

    @property
    def HoverThemeColorIndex(self):
        return self.commandbutton.HoverThemeColorIndex

    @property
    def HoverTint(self):
        return self.commandbutton.HoverTint

    @property
    def Hyperlink(self):
        return self.commandbutton.Hyperlink

    @property
    def HyperlinkAddress(self):
        return self.commandbutton.HyperlinkAddress

    @property
    def HyperlinkSubAddress(self):
        return self.commandbutton.HyperlinkSubAddress

    @property
    def InSelection(self):
        return self.commandbutton.InSelection

    @property
    def IsVisible(self):
        return self.commandbutton.IsVisible

    @property
    def LabelAlign(self):
        return self.commandbutton.LabelAlign

    @property
    def LabelX(self):
        return self.commandbutton.LabelX

    @property
    def LabelY(self):
        return self.commandbutton.LabelY

    @property
    def Layout(self):
        return AcLayoutType(self.commandbutton.Layout)

    @property
    def LayoutID(self):
        return self.commandbutton.LayoutID

    @property
    def Left(self):
        return self.commandbutton.Left

    @property
    def LeftPadding(self):
        return self.commandbutton.LeftPadding

    @property
    def Name(self):
        return self.commandbutton.Name

    @property
    def ObjectPalette(self):
        return self.commandbutton.ObjectPalette

    @property
    def OldValue(self):
        return self.commandbutton.OldValue

    @property
    def OnClick(self):
        return self.commandbutton.OnClick

    @property
    def OnDblClick(self):
        return self.commandbutton.OnDblClick

    @property
    def OnEnter(self):
        return self.commandbutton.OnEnter

    @property
    def OnExit(self):
        return self.commandbutton.OnExit

    @property
    def OnGotFocus(self):
        return self.commandbutton.OnGotFocus

    @property
    def OnKeyDown(self):
        return self.commandbutton.OnKeyDown

    @property
    def OnKeyPress(self):
        return self.commandbutton.OnKeyPress

    @property
    def OnKeyUp(self):
        return self.commandbutton.OnKeyUp

    @property
    def OnLostFocus(self):
        return self.commandbutton.OnLostFocus

    @property
    def OnMouseDown(self):
        return self.commandbutton.OnMouseDown

    @property
    def OnMouseMove(self):
        return self.commandbutton.OnMouseMove

    @property
    def OnMouseUp(self):
        return self.commandbutton.OnMouseUp

    @property
    def OnPush(self):
        return self.commandbutton.OnPush

    @property
    def Parent(self):
        return self.commandbutton.Parent

    @property
    def Picture(self):
        return self.commandbutton.Picture

    @property
    def PictureCaptionArrangement(self):
        return self.commandbutton.PictureCaptionArrangement

    @property
    def PictureData(self):
        return self.commandbutton.PictureData

    @property
    def PictureType(self):
        return self.commandbutton.PictureType

    @property
    def PressedColor(self):
        return self.commandbutton.PressedColor

    @property
    def PressedForeColor(self):
        return self.commandbutton.PressedForeColor

    @property
    def PressedForeShade(self):
        return self.commandbutton.PressedForeShade

    @property
    def PressedForeThemeColorIndex(self):
        return self.commandbutton.PressedForeThemeColorIndex

    @property
    def PressedForeTint(self):
        return self.commandbutton.PressedForeTint

    @property
    def PressedShade(self):
        return self.commandbutton.PressedShade

    @property
    def PressedThemeColorIndex(self):
        return self.commandbutton.PressedThemeColorIndex

    @property
    def PressedTint(self):
        return self.commandbutton.PressedTint

    @property
    def Properties(self):
        return Properties(self.commandbutton.Properties)

    @property
    def QuickStyle(self):
        return self.commandbutton.QuickStyle

    @property
    def ReadingOrder(self):
        return self.commandbutton.ReadingOrder

    @property
    def RightPadding(self):
        return self.commandbutton.RightPadding

    @property
    def Section(self):
        return self.commandbutton.Section

    @property
    def Shadow(self):
        return self.commandbutton.Shadow

    @property
    def Shape(self):
        return self.commandbutton.Shape

    @Shape.setter
    def Shape(self, value):
        self.commandbutton.Shape = value

    @property
    def ShortcutMenuBar(self):
        return self.commandbutton.ShortcutMenuBar

    @property
    def SoftEdges(self):
        return self.commandbutton.SoftEdges

    @property
    def StatusBarText(self):
        return self.commandbutton.StatusBarText

    @property
    def TabIndex(self):
        return self.commandbutton.TabIndex

    @property
    def TabStop(self):
        return self.commandbutton.TabStop

    @property
    def Tag(self):
        return self.commandbutton.Tag

    @property
    def ThemeFontIndex(self):
        return self.commandbutton.ThemeFontIndex

    @property
    def Top(self):
        return self.commandbutton.Top

    @property
    def TopPadding(self):
        return self.commandbutton.TopPadding

    @property
    def Transparent(self):
        return self.commandbutton.Transparent

    @property
    def UseTheme(self):
        return self.commandbutton.UseTheme

    @property
    def VerticalAnchor(self):
        return self.commandbutton.VerticalAnchor

    @property
    def Visible(self):
        return self.commandbutton.Visible

    @Visible.setter
    def Visible(self, value):
        self.commandbutton.Visible = value

    @property
    def Width(self):
        return self.commandbutton.Width

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

    def Column(self, *args, Index=None, Row=None):
        arguments = {"Index": Index, "Row": Row}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.control.Column(*args, **arguments)

    @property
    def Controls(self):
        return Controls(self.control.Controls)

    @property
    def Form(self):
        return self.control.Form

    @property
    def GridlineColor(self):
        return self.control.GridlineColor

    @property
    def GridlineStyleBottom(self):
        return self.control.GridlineStyleBottom

    @property
    def GridlineStyleLeft(self):
        return self.control.GridlineStyleLeft

    @property
    def GridlineStyleRight(self):
        return self.control.GridlineStyleRight

    @property
    def GridlineStyleTop(self):
        return self.control.GridlineStyleTop

    @property
    def GridlineWidthBottom(self):
        return self.control.GridlineWidthBottom

    @property
    def GridlineWidthLeft(self):
        return self.control.GridlineWidthLeft

    @property
    def GridlineWidthRight(self):
        return self.control.GridlineWidthRight

    @property
    def GridlineWidthTop(self):
        return self.control.GridlineWidthTop

    @property
    def HorizontalAnchor(self):
        return self.control.HorizontalAnchor

    @property
    def Hyperlink(self):
        return self.control.Hyperlink

    def ItemData(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.control.ItemData(*args, **arguments)

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

    @property
    def Name(self):
        return self.control.Name

    @property
    def Object(self):
        return self.control.Object

    def ObjectVerbs(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.control.ObjectVerbs(*args, **arguments)

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

    def Selected(self, *args, lRow=None):
        arguments = {"lRow": lRow}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.control.Selected(*args, **arguments)

    @property
    def SmartTags(self):
        return SmartTags(self.control.SmartTags)

    @property
    def TopPadding(self):
        return self.control.TopPadding

    @property
    def VerticalAnchor(self):
        return self.control.VerticalAnchor

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
        return self.controls.Item(*args, **arguments)

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

    @property
    def BorderShade(self):
        return self.customcontrol.BorderShade

    @property
    def BorderStyle(self):
        return self.customcontrol.BorderStyle

    @property
    def BorderThemeColorIndex(self):
        return self.customcontrol.BorderThemeColorIndex

    @property
    def BorderTint(self):
        return self.customcontrol.BorderTint

    @property
    def BorderWidth(self):
        return self.customcontrol.BorderWidth

    @property
    def BottomPadding(self):
        return self.customcontrol.BottomPadding

    @property
    def Cancel(self):
        return self.customcontrol.Cancel

    @property
    def Class(self):
        return self.customcontrol.Class

    @property
    def Controls(self):
        return Controls(self.customcontrol.Controls)

    @property
    def ControlSource(self):
        return self.customcontrol.ControlSource

    @property
    def ControlTipText(self):
        return self.customcontrol.ControlTipText

    @property
    def ControlType(self):
        return self.customcontrol.ControlType

    @property
    def Custom(self):
        return self.customcontrol.Custom

    @Custom.setter
    def Custom(self, value):
        self.customcontrol.Custom = value

    @property
    def Default(self):
        return self.customcontrol.Default

    @property
    def DisplayWhen(self):
        return self.customcontrol.DisplayWhen

    @property
    def Enabled(self):
        return self.customcontrol.Enabled

    @property
    def EventProcPrefix(self):
        return self.customcontrol.EventProcPrefix

    @property
    def GridlineColor(self):
        return self.customcontrol.GridlineColor

    @property
    def GridlineStyleBottom(self):
        return self.customcontrol.GridlineStyleBottom

    @property
    def GridlineStyleLeft(self):
        return self.customcontrol.GridlineStyleLeft

    @property
    def GridlineStyleRight(self):
        return self.customcontrol.GridlineStyleRight

    @property
    def GridlineStyleTop(self):
        return self.customcontrol.GridlineStyleTop

    @property
    def GridlineWidthBottom(self):
        return self.customcontrol.GridlineWidthBottom

    @property
    def GridlineWidthLeft(self):
        return self.customcontrol.GridlineWidthLeft

    @property
    def GridlineWidthRight(self):
        return self.customcontrol.GridlineWidthRight

    @property
    def GridlineWidthTop(self):
        return self.customcontrol.GridlineWidthTop

    @property
    def Height(self):
        return self.customcontrol.Height

    @property
    def HelpContextId(self):
        return self.customcontrol.HelpContextId

    @property
    def HorizontalAnchor(self):
        return self.customcontrol.HorizontalAnchor

    @property
    def InSelection(self):
        return self.customcontrol.InSelection

    @property
    def IsVisible(self):
        return self.customcontrol.IsVisible

    @property
    def Layout(self):
        return AcLayoutType(self.customcontrol.Layout)

    @property
    def LayoutID(self):
        return self.customcontrol.LayoutID

    @property
    def Left(self):
        return self.customcontrol.Left

    @property
    def LeftPadding(self):
        return self.customcontrol.LeftPadding

    @property
    def Locked(self):
        return self.customcontrol.Locked

    @property
    def Name(self):
        return self.customcontrol.Name

    @property
    def Object(self):
        return self.customcontrol.Object

    @property
    def ObjectPalette(self):
        return self.customcontrol.ObjectPalette

    def ObjectVerbs(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.customcontrol.ObjectVerbs(*args, **arguments)

    @property
    def ObjectVerbsCount(self):
        return self.customcontrol.ObjectVerbsCount

    @property
    def OldBorderStyle(self):
        return self.customcontrol.OldBorderStyle

    @property
    def OldValue(self):
        return self.customcontrol.OldValue

    @property
    def OLEClass(self):
        return self.customcontrol.OLEClass

    @property
    def OnEnter(self):
        return self.customcontrol.OnEnter

    @property
    def OnExit(self):
        return self.customcontrol.OnExit

    @property
    def OnGotFocus(self):
        return self.customcontrol.OnGotFocus

    @property
    def OnLostFocus(self):
        return self.customcontrol.OnLostFocus

    @property
    def OnUpdated(self):
        return self.customcontrol.OnUpdated

    @property
    def Parent(self):
        return self.customcontrol.Parent

    @property
    def Properties(self):
        return Properties(self.customcontrol.Properties)

    @property
    def RightPadding(self):
        return self.customcontrol.RightPadding

    @property
    def Section(self):
        return self.customcontrol.Section

    @property
    def SpecialEffect(self):
        return self.customcontrol.SpecialEffect

    @property
    def TabIndex(self):
        return self.customcontrol.TabIndex

    @property
    def TabStop(self):
        return self.customcontrol.TabStop

    @property
    def Tag(self):
        return self.customcontrol.Tag

    @property
    def Top(self):
        return self.customcontrol.Top

    @property
    def TopPadding(self):
        return self.customcontrol.TopPadding

    @property
    def Value(self):
        return self.customcontrol.Value

    @property
    def VarOleObject(self):
        return self.customcontrol.VarOleObject

    @property
    def Verb(self):
        return self.customcontrol.Verb

    @property
    def VerticalAnchor(self):
        return self.customcontrol.VerticalAnchor

    @property
    def Visible(self):
        return self.customcontrol.Visible

    @Visible.setter
    def Visible(self, value):
        self.customcontrol.Visible = value

    @property
    def Width(self):
        return self.customcontrol.Width

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
        return self.dependencyobjects.Item(*args, **arguments)

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

    @property
    def BackShade(self):
        return self.emptycell.BackShade

    @property
    def BackStyle(self):
        return self.emptycell.BackStyle

    @property
    def BackThemeColorIndex(self):
        return self.emptycell.BackThemeColorIndex

    @property
    def BackTint(self):
        return self.emptycell.BackTint

    @property
    def BottomPadding(self):
        return self.emptycell.BottomPadding

    @property
    def ControlType(self):
        return self.emptycell.ControlType

    @property
    def DisplayWhen(self):
        return self.emptycell.DisplayWhen

    @property
    def EventProcPrefix(self):
        return self.emptycell.EventProcPrefix

    @property
    def GridlineColor(self):
        return self.emptycell.GridlineColor

    @property
    def GridlineShade(self):
        return self.emptycell.GridlineShade

    @property
    def GridlineStyleBottom(self):
        return self.emptycell.GridlineStyleBottom

    @property
    def GridlineStyleLeft(self):
        return self.emptycell.GridlineStyleLeft

    @property
    def GridlineStyleRight(self):
        return self.emptycell.GridlineStyleRight

    @property
    def GridlineStyleTop(self):
        return self.emptycell.GridlineStyleTop

    @property
    def GridlineThemeColorIndex(self):
        return self.emptycell.GridlineThemeColorIndex

    @property
    def GridlineTint(self):
        return self.emptycell.GridlineTint

    @property
    def GridlineWidthBottom(self):
        return self.emptycell.GridlineWidthBottom

    @property
    def GridlineWidthLeft(self):
        return self.emptycell.GridlineWidthLeft

    @property
    def GridlineWidthRight(self):
        return self.emptycell.GridlineWidthRight

    @property
    def GridlineWidthTop(self):
        return self.emptycell.GridlineWidthTop

    @property
    def Height(self):
        return self.emptycell.Height

    @property
    def HelpContextId(self):
        return self.emptycell.HelpContextId

    @property
    def HorizontalAnchor(self):
        return self.emptycell.HorizontalAnchor

    @property
    def InSelection(self):
        return self.emptycell.InSelection

    @property
    def IsVisible(self):
        return self.emptycell.IsVisible

    @property
    def Layout(self):
        return AcLayoutType(self.emptycell.Layout)

    @property
    def LayoutID(self):
        return self.emptycell.LayoutID

    @property
    def Left(self):
        return self.emptycell.Left

    @property
    def LeftPadding(self):
        return self.emptycell.LeftPadding

    @property
    def Name(self):
        return self.emptycell.Name

    @property
    def Parent(self):
        return self.emptycell.Parent

    @property
    def Properties(self):
        return Properties(self.emptycell.Properties)

    @property
    def RightPadding(self):
        return self.emptycell.RightPadding

    @property
    def Section(self):
        return self.emptycell.Section

    @property
    def ShortcutMenuBar(self):
        return self.emptycell.ShortcutMenuBar

    @property
    def SpecialEffect(self):
        return self.emptycell.SpecialEffect

    @property
    def StatusBarText(self):
        return self.emptycell.StatusBarText

    @property
    def Tag(self):
        return self.emptycell.Tag

    @property
    def Top(self):
        return self.emptycell.Top

    @property
    def TopPadding(self):
        return self.emptycell.TopPadding

    @property
    def VerticalAnchor(self):
        return self.emptycell.VerticalAnchor

    @property
    def Visible(self):
        return self.emptycell.Visible

    @Visible.setter
    def Visible(self, value):
        self.emptycell.Visible = value

    @property
    def Width(self):
        return self.emptycell.Width

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
        return self.entities.Item(*args, **arguments)

    @property
    def Parent(self):
        return self.entities.Parent

class Entity:

    def __init__(self, entity=None):
        self.entity = entity

    @property
    def Name(self):
        return self.entity.Name

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

    @property
    def AllowDatasheetView(self):
        return self.form.AllowDatasheetView

    @AllowDatasheetView.setter
    def AllowDatasheetView(self, value):
        self.form.AllowDatasheetView = value

    @property
    def AllowDeletions(self):
        return self.form.AllowDeletions

    @property
    def AllowEdits(self):
        return self.form.AllowEdits

    @property
    def AllowFilters(self):
        return self.form.AllowFilters

    @property
    def AllowFormView(self):
        return self.form.AllowFormView

    @AllowFormView.setter
    def AllowFormView(self, value):
        self.form.AllowFormView = value

    @property
    def AllowLayoutView(self):
        return self.form.AllowLayoutView

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

    @property
    def BorderStyle(self):
        return self.form.BorderStyle

    @property
    def Caption(self):
        return self.form.Caption

    @property
    def ChartSpace(self):
        return self.form.ChartSpace

    @property
    def CloseButton(self):
        return self.form.CloseButton

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

    @property
    def Controls(self):
        return Controls(self.form.Controls)

    @property
    def Count(self):
        return self.form.Count

    @property
    def CurrentRecord(self):
        return self.form.CurrentRecord

    @property
    def CurrentSectionLeft(self):
        return self.form.CurrentSectionLeft

    @property
    def CurrentSectionTop(self):
        return self.form.CurrentSectionTop

    @property
    def CurrentView(self):
        return self.form.CurrentView

    @property
    def Cycle(self):
        return self.form.Cycle

    @property
    def DataChange(self):
        return self.form.DataChange

    @DataChange.setter
    def DataChange(self, value):
        self.form.DataChange = value

    @property
    def DataEntry(self):
        return self.form.DataEntry

    @property
    def DataSetChange(self):
        return self.form.DataSetChange

    @DataSetChange.setter
    def DataSetChange(self, value):
        self.form.DataSetChange = value

    @property
    def DatasheetAlternateBackColor(self):
        return self.form.DatasheetAlternateBackColor

    @property
    def DatasheetBackColor(self):
        return self.form.DatasheetBackColor

    @property
    def DatasheetBorderLineStyle(self):
        return self.form.DatasheetBorderLineStyle

    @DatasheetBorderLineStyle.setter
    def DatasheetBorderLineStyle(self, value):
        self.form.DatasheetBorderLineStyle = value

    @property
    def DatasheetCellsEffect(self):
        return self.form.DatasheetCellsEffect

    @property
    def DatasheetColumnHeaderUnderlineStyle(self):
        return self.form.DatasheetColumnHeaderUnderlineStyle

    @DatasheetColumnHeaderUnderlineStyle.setter
    def DatasheetColumnHeaderUnderlineStyle(self, value):
        self.form.DatasheetColumnHeaderUnderlineStyle = value

    @property
    def DatasheetFontHeight(self):
        return self.form.DatasheetFontHeight

    @property
    def DatasheetFontItalic(self):
        return self.form.DatasheetFontItalic

    @property
    def DatasheetFontName(self):
        return self.form.DatasheetFontName

    @property
    def DatasheetFontUnderline(self):
        return self.form.DatasheetFontUnderline

    @property
    def DatasheetFontWeight(self):
        return self.form.DatasheetFontWeight

    @property
    def DatasheetForeColor(self):
        return self.form.DatasheetForeColor

    @property
    def DatasheetGridlinesBehavior(self):
        return self.form.DatasheetGridlinesBehavior

    @property
    def DatasheetGridlinesColor(self):
        return self.form.DatasheetGridlinesColor

    def DefaultControl(self, *args, ControlType=None):
        arguments = {"ControlType": ControlType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.form.DefaultControl(*args, **arguments)

    @property
    def DefaultView(self):
        return self.form.DefaultView

    @property
    def Dirty(self):
        return self.form.Dirty

    @property
    def DisplayOnSharePointSite(self):
        return self.form.DisplayOnSharePointSite

    @property
    def DividingLines(self):
        return self.form.DividingLines

    @property
    def FastLaserPrinting(self):
        return self.form.FastLaserPrinting

    @property
    def FetchDefaults(self):
        return self.form.FetchDefaults

    @FetchDefaults.setter
    def FetchDefaults(self, value):
        self.form.FetchDefaults = value

    @property
    def Filter(self):
        return self.form.Filter

    @property
    def FilterOn(self):
        return self.form.FilterOn

    @property
    def FilterOnLoad(self):
        return self.form.FilterOnLoad

    @property
    def FitToScreen(self):
        return self.form.FitToScreen

    @property
    def Form(self):
        return self.form.Form

    @property
    def FrozenColumns(self):
        return self.form.FrozenColumns

    @property
    def GridX(self):
        return self.form.GridX

    @property
    def GridY(self):
        return self.form.GridY

    @property
    def HasModule(self):
        return self.form.HasModule

    @property
    def HelpContextId(self):
        return self.form.HelpContextId

    @property
    def HelpFile(self):
        return self.form.HelpFile

    @property
    def HorizontalDatasheetGridlineStyle(self):
        return self.form.HorizontalDatasheetGridlineStyle

    @HorizontalDatasheetGridlineStyle.setter
    def HorizontalDatasheetGridlineStyle(self, value):
        self.form.HorizontalDatasheetGridlineStyle = value

    @property
    def Hwnd(self):
        return self.form.Hwnd

    @property
    def InsideHeight(self):
        return self.form.InsideHeight

    @property
    def InsideWidth(self):
        return self.form.InsideWidth

    @property
    def KeyPreview(self):
        return self.form.KeyPreview

    @property
    def LayoutForPrint(self):
        return self.form.LayoutForPrint

    @property
    def MaxRecButton(self):
        return self.form.MaxRecButton

    @property
    def MaxRecords(self):
        return self.form.MaxRecords

    @property
    def MenuBar(self):
        return self.form.MenuBar

    @property
    def MinMaxButtons(self):
        return self.form.MinMaxButtons

    @property
    def Modal(self):
        return self.form.Modal

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

    @property
    def NavigationButtons(self):
        return self.form.NavigationButtons

    @property
    def NavigationCaption(self):
        return self.form.NavigationCaption

    @property
    def NewRecord(self):
        return self.form.NewRecord

    @property
    def OnActivate(self):
        return self.form.OnActivate

    @property
    def OnApplyFilter(self):
        return self.form.OnApplyFilter

    @property
    def OnClick(self):
        return self.form.OnClick

    @property
    def OnClose(self):
        return self.form.OnClose

    @property
    def OnConnect(self):
        return self.form.OnConnect

    @OnConnect.setter
    def OnConnect(self, value):
        self.form.OnConnect = value

    @property
    def OnCurrent(self):
        return self.form.OnCurrent

    @property
    def OnDblClick(self):
        return self.form.OnDblClick

    @property
    def OnDeactivate(self):
        return self.form.OnDeactivate

    @property
    def OnDelete(self):
        return self.form.OnDelete

    @property
    def OnDirty(self):
        return self.form.OnDirty

    @property
    def OnDisconnect(self):
        return self.form.OnDisconnect

    @OnDisconnect.setter
    def OnDisconnect(self, value):
        self.form.OnDisconnect = value

    @property
    def OnError(self):
        return self.form.OnError

    @property
    def OnFilter(self):
        return self.form.OnFilter

    @property
    def OnGotFocus(self):
        return self.form.OnGotFocus

    @property
    def OnInsert(self):
        return self.form.OnInsert

    @property
    def OnKeyDown(self):
        return self.form.OnKeyDown

    @property
    def OnKeyPress(self):
        return self.form.OnKeyPress

    @property
    def OnKeyUp(self):
        return self.form.OnKeyUp

    @property
    def OnLoad(self):
        return self.form.OnLoad

    @property
    def OnLostFocus(self):
        return self.form.OnLostFocus

    @property
    def OnMouseDown(self):
        return self.form.OnMouseDown

    @property
    def OnMouseMove(self):
        return self.form.OnMouseMove

    @property
    def OnMouseUp(self):
        return self.form.OnMouseUp

    @property
    def OnOpen(self):
        return self.form.OnOpen

    @property
    def OnResize(self):
        return self.form.OnResize

    @property
    def OnTimer(self):
        return self.form.OnTimer

    @property
    def OnUndo(self):
        return self.form.OnUndo

    @OnUndo.setter
    def OnUndo(self, value):
        self.form.OnUndo = value

    @property
    def OnUnload(self):
        return self.form.OnUnload

    @property
    def OpenArgs(self):
        return self.form.OpenArgs

    @property
    def OrderBy(self):
        return self.form.OrderBy

    @property
    def OrderByOn(self):
        return self.form.OrderByOn

    @property
    def OrderByOnLoad(self):
        return self.form.OrderByOnLoad

    @property
    def Orientation(self):
        return self.form.Orientation

    @property
    def Page(self):
        return self.form.Page

    @property
    def Pages(self):
        return self.form.Pages

    @property
    def Painting(self):
        return self.form.Painting

    @property
    def PaintPalette(self):
        return self.form.PaintPalette

    @property
    def PaletteSource(self):
        return self.form.PaletteSource

    @property
    def Parent(self):
        return self.form.Parent

    @property
    def Picture(self):
        return self.form.Picture

    @property
    def PictureAlignment(self):
        return self.form.PictureAlignment

    @property
    def PictureData(self):
        return self.form.PictureData

    @property
    def PicturePalette(self):
        return self.form.PicturePalette

    @property
    def PictureSizeMode(self):
        return self.form.PictureSizeMode

    @property
    def PictureTiling(self):
        return self.form.PictureTiling

    @property
    def PictureType(self):
        return self.form.PictureType

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

    @property
    def PrtDevNames(self):
        return self.form.PrtDevNames

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

    @property
    def RecordSelectors(self):
        return self.form.RecordSelectors

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

    @property
    def RecordSourceQualifier(self):
        return self.form.RecordSourceQualifier

    @RecordSourceQualifier.setter
    def RecordSourceQualifier(self, value):
        self.form.RecordSourceQualifier = value

    @property
    def ResyncCommand(self):
        return self.form.ResyncCommand

    @property
    def RibbonName(self):
        return self.form.RibbonName

    @property
    def RowHeight(self):
        return self.form.RowHeight

    @property
    def ScrollBars(self):
        return self.form.ScrollBars

    def Section(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.form.Section(*args, **arguments)

    @property
    def SelectionChange(self):
        return self.form.SelectionChange

    @SelectionChange.setter
    def SelectionChange(self, value):
        self.form.SelectionChange = value

    @property
    def SelHeight(self):
        return self.form.SelHeight

    @property
    def SelLeft(self):
        return self.form.SelLeft

    @property
    def SelTop(self):
        return self.form.SelTop

    @property
    def SelWidth(self):
        return self.form.SelWidth

    @property
    def ServerFilter(self):
        return self.form.ServerFilter

    @property
    def ServerFilterByForm(self):
        return self.form.ServerFilterByForm

    @property
    def ShortcutMenu(self):
        return self.form.ShortcutMenu

    @property
    def ShortcutMenuBar(self):
        return self.form.ShortcutMenuBar

    @property
    def SplitFormDatasheet(self):
        return self.form.SplitFormDatasheet

    @property
    def SplitFormOrientation(self):
        return self.form.SplitFormOrientation

    @property
    def SplitFormPrinting(self):
        return self.form.SplitFormPrinting

    @property
    def SplitFormSize(self):
        return self.form.SplitFormSize

    @property
    def SplitFormSplitterBar(self):
        return self.form.SplitFormSplitterBar

    @property
    def SplitFormSplitterBarSave(self):
        return self.form.SplitFormSplitterBarSave

    @property
    def SubdatasheetExpanded(self):
        return self.form.SubdatasheetExpanded

    @property
    def SubdatasheetHeight(self):
        return self.form.SubdatasheetHeight

    @property
    def Tag(self):
        return self.form.Tag

    @property
    def TimerInterval(self):
        return self.form.TimerInterval

    @property
    def Toolbar(self):
        return self.form.Toolbar

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

    @property
    def Visible(self):
        return self.form.Visible

    @Visible.setter
    def Visible(self, value):
        self.form.Visible = value

    @property
    def Width(self):
        return self.form.Width

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

    @property
    def Enabled(self):
        return self.formatcondition.Enabled

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

    @property
    def LongestBarLimit(self):
        return self.formatcondition.LongestBarLimit

    @property
    def LongestBarValue(self):
        return self.formatcondition.LongestBarValue

    @property
    def Operator(self):
        return self.formatcondition.Operator

    @property
    def ShortestBarLimit(self):
        return self.formatcondition.ShortestBarLimit

    @property
    def ShortestBarValue(self):
        return self.formatcondition.ShortestBarValue

    @property
    def ShowBarOnly(self):
        return self.formatcondition.ShowBarOnly

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
        return self.formatconditions.Item(*args, **arguments)

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
        return self.forms.Item(*args, **arguments)

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

    @property
    def GroupFooter(self):
        return self.grouplevel.GroupFooter

    @property
    def GroupHeader(self):
        return self.grouplevel.GroupHeader

    @property
    def GroupInterval(self):
        return self.grouplevel.GroupInterval

    @property
    def GroupOn(self):
        return self.grouplevel.GroupOn

    @property
    def KeepTogether(self):
        return self.grouplevel.KeepTogether

    @property
    def Parent(self):
        return self.grouplevel.Parent

    @property
    def Properties(self):
        return Properties(self.grouplevel.Properties)

    @property
    def SortOrder(self):
        return self.grouplevel.SortOrder

class Hyperlink:

    def __init__(self, hyperlink=None):
        self.hyperlink = hyperlink

    @property
    def Address(self):
        return self.hyperlink.Address

    @property
    def EmailSubject(self):
        return self.hyperlink.EmailSubject

    @property
    def ScreenTip(self):
        return self.hyperlink.ScreenTip

    @property
    def SubAddress(self):
        return self.hyperlink.SubAddress

    @property
    def TextToDisplay(self):
        return self.hyperlink.TextToDisplay

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

    @property
    def BackShade(self):
        return self.image.BackShade

    @property
    def BackStyle(self):
        return self.image.BackStyle

    @property
    def BackThemeColorIndex(self):
        return self.image.BackThemeColorIndex

    @property
    def BackTint(self):
        return self.image.BackTint

    @property
    def BorderColor(self):
        return self.image.BorderColor

    @property
    def BorderShade(self):
        return self.image.BorderShade

    @property
    def BorderStyle(self):
        return self.image.BorderStyle

    @property
    def BorderThemeColorIndex(self):
        return self.image.BorderThemeColorIndex

    @property
    def BorderTint(self):
        return self.image.BorderTint

    @property
    def BorderWidth(self):
        return self.image.BorderWidth

    @property
    def BottomPadding(self):
        return self.image.BottomPadding

    @property
    def Controls(self):
        return Controls(self.image.Controls)

    @property
    def ControlTipText(self):
        return self.image.ControlTipText

    @property
    def ControlType(self):
        return self.image.ControlType

    @property
    def DisplayWhen(self):
        return self.image.DisplayWhen

    @property
    def EventProcPrefix(self):
        return self.image.EventProcPrefix

    @property
    def GridlineColor(self):
        return self.image.GridlineColor

    @property
    def GridlineShade(self):
        return self.image.GridlineShade

    @property
    def GridlineStyleBottom(self):
        return self.image.GridlineStyleBottom

    @property
    def GridlineStyleLeft(self):
        return self.image.GridlineStyleLeft

    @property
    def GridlineStyleRight(self):
        return self.image.GridlineStyleRight

    @property
    def GridlineStyleTop(self):
        return self.image.GridlineStyleTop

    @property
    def GridlineThemeColorIndex(self):
        return self.image.GridlineThemeColorIndex

    @property
    def GridlineTint(self):
        return self.image.GridlineTint

    @property
    def GridlineWidthBottom(self):
        return self.image.GridlineWidthBottom

    @property
    def GridlineWidthLeft(self):
        return self.image.GridlineWidthLeft

    @property
    def GridlineWidthRight(self):
        return self.image.GridlineWidthRight

    @property
    def GridlineWidthTop(self):
        return self.image.GridlineWidthTop

    @property
    def Height(self):
        return self.image.Height

    @property
    def HelpContextId(self):
        return self.image.HelpContextId

    @property
    def HorizontalAnchor(self):
        return self.image.HorizontalAnchor

    @property
    def Hyperlink(self):
        return self.image.Hyperlink

    @property
    def HyperlinkAddress(self):
        return self.image.HyperlinkAddress

    @property
    def HyperlinkSubAddress(self):
        return self.image.HyperlinkSubAddress

    @property
    def ImageHeight(self):
        return self.image.ImageHeight

    @property
    def ImageWidth(self):
        return self.image.ImageWidth

    @property
    def InSelection(self):
        return self.image.InSelection

    @property
    def IsVisible(self):
        return self.image.IsVisible

    @property
    def Layout(self):
        return AcLayoutType(self.image.Layout)

    @property
    def LayoutID(self):
        return self.image.LayoutID

    @property
    def Left(self):
        return self.image.Left

    @property
    def LeftPadding(self):
        return self.image.LeftPadding

    @property
    def Name(self):
        return self.image.Name

    @property
    def ObjectPalette(self):
        return self.image.ObjectPalette

    @property
    def OldBorderStyle(self):
        return self.image.OldBorderStyle

    @property
    def OldValue(self):
        return self.image.OldValue

    @property
    def OnClick(self):
        return self.image.OnClick

    @property
    def OnDblClick(self):
        return self.image.OnDblClick

    @property
    def OnMouseDown(self):
        return self.image.OnMouseDown

    @property
    def OnMouseMove(self):
        return self.image.OnMouseMove

    @property
    def OnMouseUp(self):
        return self.image.OnMouseUp

    @property
    def Parent(self):
        return self.image.Parent

    @property
    def Picture(self):
        return self.image.Picture

    @property
    def PictureAlignment(self):
        return self.image.PictureAlignment

    @property
    def PictureData(self):
        return self.image.PictureData

    @property
    def PictureTiling(self):
        return self.image.PictureTiling

    @property
    def PictureType(self):
        return self.image.PictureType

    @property
    def Properties(self):
        return Properties(self.image.Properties)

    @property
    def RightPadding(self):
        return self.image.RightPadding

    @property
    def Section(self):
        return self.image.Section

    @property
    def ShortcutMenuBar(self):
        return self.image.ShortcutMenuBar

    @property
    def SizeMode(self):
        return self.image.SizeMode

    @property
    def SpecialEffect(self):
        return self.image.SpecialEffect

    @property
    def Tag(self):
        return self.image.Tag

    @property
    def Top(self):
        return self.image.Top

    @property
    def TopPadding(self):
        return self.image.TopPadding

    @property
    def VerticalAnchor(self):
        return self.image.VerticalAnchor

    @property
    def Visible(self):
        return self.image.Visible

    @Visible.setter
    def Visible(self, value):
        self.image.Visible = value

    @property
    def Width(self):
        return self.image.Width

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

    @property
    def Name(self):
        return self.importexportspecification.Name

    @property
    def Parent(self):
        return self.importexportspecification.Parent

    @property
    def XML(self):
        return self.importexportspecification.XML

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
        return self.importexportspecifications.Item(*args, **arguments)

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

    @property
    def BackShade(self):
        return self.label.BackShade

    @property
    def BackStyle(self):
        return self.label.BackStyle

    @property
    def BackThemeColorIndex(self):
        return self.label.BackThemeColorIndex

    @property
    def BackTint(self):
        return self.label.BackTint

    @property
    def BorderColor(self):
        return self.label.BorderColor

    @property
    def BorderShade(self):
        return self.label.BorderShade

    @property
    def BorderStyle(self):
        return self.label.BorderStyle

    @property
    def BorderThemeColorIndex(self):
        return self.label.BorderThemeColorIndex

    @property
    def BorderTint(self):
        return self.label.BorderTint

    @property
    def BorderWidth(self):
        return self.label.BorderWidth

    @property
    def BottomMargin(self):
        return self.label.BottomMargin

    @property
    def BottomPadding(self):
        return self.label.BottomPadding

    @property
    def Caption(self):
        return self.label.Caption

    @property
    def ControlTipText(self):
        return self.label.ControlTipText

    @property
    def ControlType(self):
        return self.label.ControlType

    @property
    def DisplayWhen(self):
        return self.label.DisplayWhen

    @property
    def EventProcPrefix(self):
        return self.label.EventProcPrefix

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

    @property
    def ForeColor(self):
        return self.label.ForeColor

    @property
    def ForeShade(self):
        return self.label.ForeShade

    @property
    def ForeThemeColorIndex(self):
        return self.label.ForeThemeColorIndex

    @property
    def ForeTint(self):
        return self.label.ForeTint

    @property
    def GridlineColor(self):
        return self.label.GridlineColor

    @property
    def GridlineShade(self):
        return self.label.GridlineShade

    @property
    def GridlineStyleBottom(self):
        return self.label.GridlineStyleBottom

    @property
    def GridlineStyleLeft(self):
        return self.label.GridlineStyleLeft

    @property
    def GridlineStyleRight(self):
        return self.label.GridlineStyleRight

    @property
    def GridlineStyleTop(self):
        return self.label.GridlineStyleTop

    @property
    def GridlineThemeColorIndex(self):
        return self.label.GridlineThemeColorIndex

    @property
    def GridlineTint(self):
        return self.label.GridlineTint

    @property
    def GridlineWidthBottom(self):
        return self.label.GridlineWidthBottom

    @property
    def GridlineWidthLeft(self):
        return self.label.GridlineWidthLeft

    @property
    def GridlineWidthRight(self):
        return self.label.GridlineWidthRight

    @property
    def GridlineWidthTop(self):
        return self.label.GridlineWidthTop

    @property
    def Height(self):
        return self.label.Height

    @property
    def HelpContextId(self):
        return self.label.HelpContextId

    @property
    def HorizontalAnchor(self):
        return self.label.HorizontalAnchor

    @property
    def Hyperlink(self):
        return self.label.Hyperlink

    @property
    def HyperlinkAddress(self):
        return self.label.HyperlinkAddress

    @property
    def HyperlinkSubAddress(self):
        return self.label.HyperlinkSubAddress

    @property
    def InSelection(self):
        return self.label.InSelection

    @property
    def IsVisible(self):
        return self.label.IsVisible

    @property
    def Layout(self):
        return AcLayoutType(self.label.Layout)

    @property
    def LayoutID(self):
        return self.label.LayoutID

    @property
    def Left(self):
        return self.label.Left

    @property
    def LeftMargin(self):
        return self.label.LeftMargin

    @property
    def LeftPadding(self):
        return self.label.LeftPadding

    @property
    def LineSpacing(self):
        return self.label.LineSpacing

    @property
    def Name(self):
        return self.label.Name

    @property
    def OldBorderStyle(self):
        return self.label.OldBorderStyle

    @property
    def OnClick(self):
        return self.label.OnClick

    @property
    def OnDblClick(self):
        return self.label.OnDblClick

    @property
    def OnMouseDown(self):
        return self.label.OnMouseDown

    @property
    def OnMouseMove(self):
        return self.label.OnMouseMove

    @property
    def OnMouseUp(self):
        return self.label.OnMouseUp

    @property
    def Parent(self):
        return self.label.Parent

    @property
    def Properties(self):
        return Properties(self.label.Properties)

    @property
    def ReadingOrder(self):
        return self.label.ReadingOrder

    @property
    def RightMargin(self):
        return self.label.RightMargin

    @property
    def RightPadding(self):
        return self.label.RightPadding

    @property
    def Section(self):
        return self.label.Section

    @property
    def ShortcutMenuBar(self):
        return self.label.ShortcutMenuBar

    @property
    def SmartTags(self):
        return SmartTags(self.label.SmartTags)

    @property
    def SpecialEffect(self):
        return self.label.SpecialEffect

    @property
    def Tag(self):
        return self.label.Tag

    @property
    def TextAlign(self):
        return self.label.TextAlign

    @property
    def ThemeFontIndex(self):
        return self.label.ThemeFontIndex

    @property
    def Top(self):
        return self.label.Top

    @property
    def TopMargin(self):
        return self.label.TopMargin

    @property
    def TopPadding(self):
        return self.label.TopPadding

    @property
    def Vertical(self):
        return self.label.Vertical

    @property
    def VerticalAnchor(self):
        return self.label.VerticalAnchor

    @property
    def Visible(self):
        return self.label.Visible

    @Visible.setter
    def Visible(self, value):
        self.label.Visible = value

    @property
    def Width(self):
        return self.label.Width

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

    @property
    def BorderShade(self):
        return self.line.BorderShade

    @property
    def BorderStyle(self):
        return self.line.BorderStyle

    @property
    def BorderThemeColorIndex(self):
        return self.line.BorderThemeColorIndex

    @property
    def BorderTint(self):
        return self.line.BorderTint

    @property
    def BorderWidth(self):
        return self.line.BorderWidth

    @property
    def ControlType(self):
        return self.line.ControlType

    @property
    def DisplayWhen(self):
        return self.line.DisplayWhen

    @property
    def EventProcPrefix(self):
        return self.line.EventProcPrefix

    @property
    def Height(self):
        return self.line.Height

    @property
    def HorizontalAnchor(self):
        return self.line.HorizontalAnchor

    @property
    def InSelection(self):
        return self.line.InSelection

    @property
    def IsVisible(self):
        return self.line.IsVisible

    @property
    def Left(self):
        return self.line.Left

    @property
    def LineSlant(self):
        return self.line.LineSlant

    @property
    def Name(self):
        return self.line.Name

    @property
    def OldBorderStyle(self):
        return self.line.OldBorderStyle

    @property
    def Parent(self):
        return self.line.Parent

    @property
    def Properties(self):
        return Properties(self.line.Properties)

    @property
    def Section(self):
        return self.line.Section

    @property
    def SpecialEffect(self):
        return self.line.SpecialEffect

    @property
    def Tag(self):
        return self.line.Tag

    @property
    def Top(self):
        return self.line.Top

    @property
    def VerticalAnchor(self):
        return self.line.VerticalAnchor

    @property
    def Visible(self):
        return self.line.Visible

    @Visible.setter
    def Visible(self, value):
        self.line.Visible = value

    @property
    def Width(self):
        return self.line.Width

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

    @property
    def AllowValueListEdits(self):
        return self.listbox.AllowValueListEdits

    @property
    def Application(self):
        return self.listbox.Application

    @property
    def AutoLabel(self):
        return self.listbox.AutoLabel

    @property
    def BackColor(self):
        return self.listbox.BackColor

    @property
    def BackShade(self):
        return self.listbox.BackShade

    @property
    def BackThemeColorIndex(self):
        return self.listbox.BackThemeColorIndex

    @property
    def BackTint(self):
        return self.listbox.BackTint

    @property
    def BorderColor(self):
        return self.listbox.BorderColor

    @property
    def BorderShade(self):
        return self.listbox.BorderShade

    @property
    def BorderStyle(self):
        return self.listbox.BorderStyle

    @property
    def BorderThemeColorIndex(self):
        return self.listbox.BorderThemeColorIndex

    @property
    def BorderTint(self):
        return self.listbox.BorderTint

    @property
    def BorderWidth(self):
        return self.listbox.BorderWidth

    @property
    def BottomPadding(self):
        return self.listbox.BottomPadding

    @property
    def BoundColumn(self):
        return self.listbox.BoundColumn

    def Column(self, *args, Index=None, Row=None):
        arguments = {"Index": Index, "Row": Row}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.listbox.Column(*args, **arguments)

    @property
    def ColumnCount(self):
        return self.listbox.ColumnCount

    @property
    def ColumnHeads(self):
        return self.listbox.ColumnHeads

    @property
    def ColumnHidden(self):
        return self.listbox.ColumnHidden

    @property
    def ColumnOrder(self):
        return self.listbox.ColumnOrder

    @property
    def ColumnWidth(self):
        return self.listbox.ColumnWidth

    @property
    def ColumnWidths(self):
        return self.listbox.ColumnWidths

    @property
    def Controls(self):
        return Controls(self.listbox.Controls)

    @property
    def ControlSource(self):
        return self.listbox.ControlSource

    @property
    def ControlTipText(self):
        return self.listbox.ControlTipText

    @property
    def ControlType(self):
        return self.listbox.ControlType

    @property
    def DefaultValue(self):
        return self.listbox.DefaultValue

    @property
    def DisplayWhen(self):
        return self.listbox.DisplayWhen

    @property
    def Enabled(self):
        return self.listbox.Enabled

    @property
    def EventProcPrefix(self):
        return self.listbox.EventProcPrefix

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

    @property
    def ForeColor(self):
        return self.listbox.ForeColor

    @property
    def ForeShade(self):
        return self.listbox.ForeShade

    @property
    def ForeThemeColorIndex(self):
        return self.listbox.ForeThemeColorIndex

    @property
    def ForeTint(self):
        return self.listbox.ForeTint

    @property
    def GridlineColor(self):
        return self.listbox.GridlineColor

    @property
    def GridlineShade(self):
        return self.listbox.GridlineShade

    @property
    def GridlineStyleBottom(self):
        return self.listbox.GridlineStyleBottom

    @property
    def GridlineStyleLeft(self):
        return self.listbox.GridlineStyleLeft

    @property
    def GridlineStyleRight(self):
        return self.listbox.GridlineStyleRight

    @property
    def GridlineStyleTop(self):
        return self.listbox.GridlineStyleTop

    @property
    def GridlineThemeColorIndex(self):
        return self.listbox.GridlineThemeColorIndex

    @property
    def GridlineTint(self):
        return self.listbox.GridlineTint

    @property
    def GridlineWidthBottom(self):
        return self.listbox.GridlineWidthBottom

    @property
    def GridlineWidthLeft(self):
        return self.listbox.GridlineWidthLeft

    @property
    def GridlineWidthRight(self):
        return self.listbox.GridlineWidthRight

    @property
    def GridlineWidthTop(self):
        return self.listbox.GridlineWidthTop

    @property
    def Height(self):
        return self.listbox.Height

    @property
    def HelpContextId(self):
        return self.listbox.HelpContextId

    @property
    def HideDuplicates(self):
        return self.listbox.HideDuplicates

    @property
    def HorizontalAnchor(self):
        return self.listbox.HorizontalAnchor

    @property
    def Hyperlink(self):
        return self.listbox.Hyperlink

    @property
    def IMEHold(self):
        return self.listbox.IMEHold

    @property
    def InheritValueList(self):
        return self.listbox.InheritValueList

    @property
    def InSelection(self):
        return self.listbox.InSelection

    @property
    def IsVisible(self):
        return self.listbox.IsVisible

    def ItemData(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.listbox.ItemData(*args, **arguments)

    @property
    def ItemsSelected(self):
        return self.listbox.ItemsSelected

    @property
    def LabelAlign(self):
        return self.listbox.LabelAlign

    @property
    def LabelX(self):
        return self.listbox.LabelX

    @property
    def LabelY(self):
        return self.listbox.LabelY

    @property
    def Layout(self):
        return AcLayoutType(self.listbox.Layout)

    @property
    def LayoutID(self):
        return self.listbox.LayoutID

    @property
    def Left(self):
        return self.listbox.Left

    @property
    def LeftPadding(self):
        return self.listbox.LeftPadding

    @property
    def ListCount(self):
        return self.listbox.ListCount

    @property
    def ListIndex(self):
        return self.listbox.ListIndex

    @property
    def ListItemsEditForm(self):
        return self.listbox.ListItemsEditForm

    @property
    def Locked(self):
        return self.listbox.Locked

    @property
    def MultiSelect(self):
        return self.listbox.MultiSelect

    @property
    def Name(self):
        return self.listbox.Name

    @property
    def OldBorderStyle(self):
        return self.listbox.OldBorderStyle

    @property
    def OldValue(self):
        return self.listbox.OldValue

    @property
    def OnClick(self):
        return self.listbox.OnClick

    @property
    def OnDblClick(self):
        return self.listbox.OnDblClick

    @property
    def OnEnter(self):
        return self.listbox.OnEnter

    @property
    def OnExit(self):
        return self.listbox.OnExit

    @property
    def OnGotFocus(self):
        return self.listbox.OnGotFocus

    @property
    def OnKeyDown(self):
        return self.listbox.OnKeyDown

    @property
    def OnKeyPress(self):
        return self.listbox.OnKeyPress

    @property
    def OnKeyUp(self):
        return self.listbox.OnKeyUp

    @property
    def OnLostFocus(self):
        return self.listbox.OnLostFocus

    @property
    def OnMouseDown(self):
        return self.listbox.OnMouseDown

    @property
    def OnMouseMove(self):
        return self.listbox.OnMouseMove

    @property
    def OnMouseUp(self):
        return self.listbox.OnMouseUp

    @property
    def Parent(self):
        return self.listbox.Parent

    @property
    def Properties(self):
        return Properties(self.listbox.Properties)

    @property
    def ReadingOrder(self):
        return self.listbox.ReadingOrder

    @property
    def Recordset(self):
        return self.listbox.Recordset

    @Recordset.setter
    def Recordset(self, value):
        self.listbox.Recordset = value

    @property
    def RightPadding(self):
        return self.listbox.RightPadding

    @property
    def RowSource(self):
        return self.listbox.RowSource

    @property
    def RowSourceType(self):
        return self.listbox.RowSourceType

    @property
    def ScrollBarAlign(self):
        return self.listbox.ScrollBarAlign

    @property
    def Section(self):
        return self.listbox.Section

    def Selected(self, *args, lRow=None):
        arguments = {"lRow": lRow}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.listbox.Selected(*args, **arguments)

    @property
    def ShortcutMenuBar(self):
        return self.listbox.ShortcutMenuBar

    @property
    def ShowOnlyRowSourceValues(self):
        return self.listbox.ShowOnlyRowSourceValues

    @property
    def SmartTags(self):
        return SmartTags(self.listbox.SmartTags)

    @property
    def SpecialEffect(self):
        return self.listbox.SpecialEffect

    @property
    def StatusBarText(self):
        return self.listbox.StatusBarText

    @property
    def TabIndex(self):
        return self.listbox.TabIndex

    @property
    def TabStop(self):
        return self.listbox.TabStop

    @property
    def Tag(self):
        return self.listbox.Tag

    @property
    def ThemeFontIndex(self):
        return self.listbox.ThemeFontIndex

    @property
    def Top(self):
        return self.listbox.Top

    @property
    def TopPadding(self):
        return self.listbox.TopPadding

    @property
    def ValidationRule(self):
        return self.listbox.ValidationRule

    @property
    def ValidationText(self):
        return self.listbox.ValidationText

    @property
    def Value(self):
        return self.listbox.Value

    @property
    def VerticalAnchor(self):
        return self.listbox.VerticalAnchor

    @property
    def Visible(self):
        return self.listbox.Visible

    @Visible.setter
    def Visible(self, value):
        self.listbox.Visible = value

    @property
    def Width(self):
        return self.listbox.Width

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
        return self.module.Lines(*args, **arguments)

    @property
    def Name(self):
        return self.module.Name

    @property
    def Parent(self):
        return self.module.Parent

    def ProcBodyLine(self, *args, ProcName=None, ProcKind=None):
        arguments = {"ProcName": ProcName, "ProcKind": ProcKind}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.module.ProcBodyLine(*args, **arguments)

    def ProcCountLines(self, *args, ProcName=None, ProcKind=None):
        arguments = {"ProcName": ProcName, "ProcKind": ProcKind}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.module.ProcCountLines(*args, **arguments)

    def ProcOfLine(self, *args, Line=None, ProcKind=None):
        arguments = {"Line": Line, "ProcKind": ProcKind}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.module.ProcOfLine(*args, **arguments)

    def ProcStartLine(self, *args, ProcName=None, ProcKind=None):
        arguments = {"ProcName": ProcName, "ProcKind": ProcKind}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.module.ProcStartLine(*args, **arguments)

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
        return self.modules.Item(*args, **arguments)

    @property
    def Parent(self):
        return self.modules.Parent

class NavigationButton:

    def __init__(self, navigationbutton=None):
        self.navigationbutton = navigationbutton

    @property
    def AddColon(self):
        return self.navigationbutton.AddColon

    @property
    def Alignment(self):
        return self.navigationbutton.Alignment

    @property
    def Application(self):
        return self.navigationbutton.Application

    @property
    def AutoLabel(self):
        return self.navigationbutton.AutoLabel

    @property
    def AutoRepeat(self):
        return self.navigationbutton.AutoRepeat

    @property
    def BackColor(self):
        return self.navigationbutton.BackColor

    @property
    def BackShade(self):
        return self.navigationbutton.BackShade

    @property
    def BackStyle(self):
        return self.navigationbutton.BackStyle

    @property
    def BackThemeColorIndex(self):
        return self.navigationbutton.BackThemeColorIndex

    @property
    def BackTint(self):
        return self.navigationbutton.BackTint

    @property
    def Bevel(self):
        return self.navigationbutton.Bevel

    @property
    def BorderColor(self):
        return self.navigationbutton.BorderColor

    @property
    def BorderShade(self):
        return self.navigationbutton.BorderShade

    @property
    def BorderStyle(self):
        return self.navigationbutton.BorderStyle

    @property
    def BorderThemeColorIndex(self):
        return self.navigationbutton.BorderThemeColorIndex

    @property
    def BorderTint(self):
        return self.navigationbutton.BorderTint

    @property
    def BorderWidth(self):
        return self.navigationbutton.BorderWidth

    @property
    def BottomPadding(self):
        return self.navigationbutton.BottomPadding

    @property
    def Caption(self):
        return self.navigationbutton.Caption

    @property
    def Controls(self):
        return Controls(self.navigationbutton.Controls)

    @property
    def ControlTipText(self):
        return self.navigationbutton.ControlTipText

    @property
    def ControlType(self):
        return self.navigationbutton.ControlType

    @property
    def CursorOnHover(self):
        return self.navigationbutton.CursorOnHover

    @property
    def DisplayWhen(self):
        return self.navigationbutton.DisplayWhen

    @property
    def Enabled(self):
        return self.navigationbutton.Enabled

    @property
    def EventProcPrefix(self):
        return self.navigationbutton.EventProcPrefix

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

    @property
    def ForeColor(self):
        return self.navigationbutton.ForeColor

    @property
    def ForeShade(self):
        return self.navigationbutton.ForeShade

    @property
    def ForeThemeColorIndex(self):
        return self.navigationbutton.ForeThemeColorIndex

    @property
    def ForeTint(self):
        return self.navigationbutton.ForeTint

    @property
    def Glow(self):
        return self.navigationbutton.Glow

    @property
    def Gradient(self):
        return self.navigationbutton.Gradient

    @property
    def GridlineColor(self):
        return self.navigationbutton.GridlineColor

    @property
    def GridlineShade(self):
        return self.navigationbutton.GridlineShade

    @property
    def GridlineStyleBottom(self):
        return self.navigationbutton.GridlineStyleBottom

    @property
    def GridlineStyleLeft(self):
        return self.navigationbutton.GridlineStyleLeft

    @property
    def GridlineStyleRight(self):
        return self.navigationbutton.GridlineStyleRight

    @property
    def GridlineStyleTop(self):
        return self.navigationbutton.GridlineStyleTop

    @property
    def GridlineThemeColorIndex(self):
        return self.navigationbutton.GridlineThemeColorIndex

    @property
    def GridlineTint(self):
        return self.navigationbutton.GridlineTint

    @property
    def GridlineWidthBottom(self):
        return self.navigationbutton.GridlineWidthBottom

    @property
    def GridlineWidthLeft(self):
        return self.navigationbutton.GridlineWidthLeft

    @property
    def GridlineWidthRight(self):
        return self.navigationbutton.GridlineWidthRight

    @property
    def GridlineWidthTop(self):
        return self.navigationbutton.GridlineWidthTop

    @property
    def Height(self):
        return self.navigationbutton.Height

    @property
    def HelpContextId(self):
        return self.navigationbutton.HelpContextId

    @property
    def HorizontalAnchor(self):
        return self.navigationbutton.HorizontalAnchor

    @property
    def HoverColor(self):
        return self.navigationbutton.HoverColor

    @property
    def HoverForeColor(self):
        return self.navigationbutton.HoverForeColor

    @property
    def HoverForeShade(self):
        return self.navigationbutton.HoverForeShade

    @property
    def HoverForeThemeColorIndex(self):
        return self.navigationbutton.HoverForeThemeColorIndex

    @property
    def HoverForeTint(self):
        return self.navigationbutton.HoverForeTint

    @property
    def HoverShade(self):
        return self.navigationbutton.HoverShade

    @property
    def HoverThemeColorIndex(self):
        return self.navigationbutton.HoverThemeColorIndex

    @property
    def HoverTint(self):
        return self.navigationbutton.HoverTint

    @property
    def Hyperlink(self):
        return self.navigationbutton.Hyperlink

    @property
    def HyperlinkAddress(self):
        return self.navigationbutton.HyperlinkAddress

    @property
    def HyperlinkSubAddress(self):
        return self.navigationbutton.HyperlinkSubAddress

    @property
    def InSelection(self):
        return self.navigationbutton.InSelection

    @property
    def IsVisible(self):
        return self.navigationbutton.IsVisible

    @property
    def LabelAlign(self):
        return self.navigationbutton.LabelAlign

    @property
    def LabelX(self):
        return self.navigationbutton.LabelX

    @property
    def LabelY(self):
        return self.navigationbutton.LabelY

    @property
    def Layout(self):
        return AcLayoutType(self.navigationbutton.Layout)

    @property
    def LayoutID(self):
        return self.navigationbutton.LayoutID

    @property
    def Left(self):
        return self.navigationbutton.Left

    @property
    def LeftPadding(self):
        return self.navigationbutton.LeftPadding

    @property
    def Name(self):
        return self.navigationbutton.Name

    @property
    def NavigationTargetName(self):
        return self.navigationbutton.NavigationTargetName

    @property
    def NavigationWhereClause(self):
        return self.navigationbutton.NavigationWhereClause

    @property
    def ObjectPalette(self):
        return self.navigationbutton.ObjectPalette

    @property
    def OldValue(self):
        return self.navigationbutton.OldValue

    @property
    def OnClick(self):
        return self.navigationbutton.OnClick

    @property
    def OnDblClick(self):
        return self.navigationbutton.OnDblClick

    @property
    def OnEnter(self):
        return self.navigationbutton.OnEnter

    @property
    def OnExit(self):
        return self.navigationbutton.OnExit

    @property
    def OnGotFocus(self):
        return self.navigationbutton.OnGotFocus

    @property
    def OnKeyDown(self):
        return self.navigationbutton.OnKeyDown

    @property
    def OnKeyPress(self):
        return self.navigationbutton.OnKeyPress

    @property
    def OnKeyUp(self):
        return self.navigationbutton.OnKeyUp

    @property
    def OnLostFocus(self):
        return self.navigationbutton.OnLostFocus

    @property
    def OnMouseDown(self):
        return self.navigationbutton.OnMouseDown

    @property
    def OnMouseMove(self):
        return self.navigationbutton.OnMouseMove

    @property
    def OnMouseUp(self):
        return self.navigationbutton.OnMouseUp

    @property
    def OnPush(self):
        return self.navigationbutton.OnPush

    @property
    def Parent(self):
        return self.navigationbutton.Parent

    @property
    def ParentTab(self):
        return self.navigationbutton.ParentTab

    @property
    def Picture(self):
        return self.navigationbutton.Picture

    @property
    def PictureCaptionArrangement(self):
        return self.navigationbutton.PictureCaptionArrangement

    @property
    def PictureData(self):
        return self.navigationbutton.PictureData

    @property
    def PictureType(self):
        return self.navigationbutton.PictureType

    @property
    def PressedColor(self):
        return self.navigationbutton.PressedColor

    @property
    def PressedForeColor(self):
        return self.navigationbutton.PressedForeColor

    @property
    def PressedForeShade(self):
        return self.navigationbutton.PressedForeShade

    @property
    def PressedForeThemeColorIndex(self):
        return self.navigationbutton.PressedForeThemeColorIndex

    @property
    def PressedForeTint(self):
        return self.navigationbutton.PressedForeTint

    @property
    def PressedShade(self):
        return self.navigationbutton.PressedShade

    @property
    def PressedThemeColorIndex(self):
        return self.navigationbutton.PressedThemeColorIndex

    @property
    def PressedTint(self):
        return self.navigationbutton.PressedTint

    @property
    def Properties(self):
        return Properties(self.navigationbutton.Properties)

    @property
    def QuickStyle(self):
        return self.navigationbutton.QuickStyle

    @property
    def ReadingOrder(self):
        return self.navigationbutton.ReadingOrder

    @property
    def RightPadding(self):
        return self.navigationbutton.RightPadding

    @property
    def Section(self):
        return self.navigationbutton.Section

    @property
    def Shadow(self):
        return self.navigationbutton.Shadow

    @property
    def Shape(self):
        return self.navigationbutton.Shape

    @Shape.setter
    def Shape(self, value):
        self.navigationbutton.Shape = value

    @property
    def ShortcutMenuBar(self):
        return self.navigationbutton.ShortcutMenuBar

    @property
    def SoftEdges(self):
        return self.navigationbutton.SoftEdges

    @property
    def StatusBarText(self):
        return self.navigationbutton.StatusBarText

    @property
    def TabIndex(self):
        return self.navigationbutton.TabIndex

    @property
    def TabStop(self):
        return self.navigationbutton.TabStop

    @property
    def Tag(self):
        return self.navigationbutton.Tag

    @property
    def ThemeFontIndex(self):
        return self.navigationbutton.ThemeFontIndex

    @property
    def Top(self):
        return self.navigationbutton.Top

    @property
    def TopPadding(self):
        return self.navigationbutton.TopPadding

    @property
    def Transparent(self):
        return self.navigationbutton.Transparent

    @property
    def VerticalAnchor(self):
        return self.navigationbutton.VerticalAnchor

    @property
    def Visible(self):
        return self.navigationbutton.Visible

    @Visible.setter
    def Visible(self, value):
        self.navigationbutton.Visible = value

    @property
    def Width(self):
        return self.navigationbutton.Width

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

    @property
    def BackColor(self):
        return self.navigationcontrol.BackColor

    @property
    def BackShade(self):
        return self.navigationcontrol.BackShade

    @property
    def BackStyle(self):
        return self.navigationcontrol.BackStyle

    @property
    def BackThemeColorIndex(self):
        return self.navigationcontrol.BackThemeColorIndex

    @property
    def BackTint(self):
        return self.navigationcontrol.BackTint

    @property
    def BorderColor(self):
        return self.navigationcontrol.BorderColor

    @property
    def BorderShade(self):
        return self.navigationcontrol.BorderShade

    @property
    def BorderStyle(self):
        return self.navigationcontrol.BorderStyle

    @property
    def BorderThemeColorIndex(self):
        return self.navigationcontrol.BorderThemeColorIndex

    @property
    def BorderTint(self):
        return self.navigationcontrol.BorderTint

    @property
    def BorderWidth(self):
        return self.navigationcontrol.BorderWidth

    @property
    def BottomPadding(self):
        return self.navigationcontrol.BottomPadding

    @property
    def Controls(self):
        return Controls(self.navigationcontrol.Controls)

    @property
    def ControlTipText(self):
        return self.navigationcontrol.ControlTipText

    @property
    def ControlType(self):
        return self.navigationcontrol.ControlType

    @property
    def DisplayWhen(self):
        return self.navigationcontrol.DisplayWhen

    @property
    def Enabled(self):
        return self.navigationcontrol.Enabled

    @property
    def EventProcPrefix(self):
        return self.navigationcontrol.EventProcPrefix

    @property
    def FilterLookup(self):
        return self.navigationcontrol.FilterLookup

    @property
    def FormatConditions(self):
        return self.navigationcontrol.FormatConditions

    @property
    def GridlineColor(self):
        return self.navigationcontrol.GridlineColor

    @property
    def GridlineShade(self):
        return self.navigationcontrol.GridlineShade

    @property
    def GridlineStyleBottom(self):
        return self.navigationcontrol.GridlineStyleBottom

    @property
    def GridlineStyleLeft(self):
        return self.navigationcontrol.GridlineStyleLeft

    @property
    def GridlineStyleRight(self):
        return self.navigationcontrol.GridlineStyleRight

    @property
    def GridlineStyleTop(self):
        return self.navigationcontrol.GridlineStyleTop

    @property
    def GridlineThemeColorIndex(self):
        return self.navigationcontrol.GridlineThemeColorIndex

    @property
    def GridlineTint(self):
        return self.navigationcontrol.GridlineTint

    @property
    def GridlineWidthBottom(self):
        return self.navigationcontrol.GridlineWidthBottom

    @property
    def GridlineWidthLeft(self):
        return self.navigationcontrol.GridlineWidthLeft

    @property
    def GridlineWidthRight(self):
        return self.navigationcontrol.GridlineWidthRight

    @property
    def GridlineWidthTop(self):
        return self.navigationcontrol.GridlineWidthTop

    @property
    def Height(self):
        return self.navigationcontrol.Height

    @property
    def HelpContextId(self):
        return self.navigationcontrol.HelpContextId

    @property
    def HorizontalAnchor(self):
        return self.navigationcontrol.HorizontalAnchor

    @property
    def Hyperlink(self):
        return self.navigationcontrol.Hyperlink

    @property
    def InSelection(self):
        return self.navigationcontrol.InSelection

    @property
    def IsVisible(self):
        return self.navigationcontrol.IsVisible

    @property
    def Layout(self):
        return AcLayoutType(self.navigationcontrol.Layout)

    @property
    def LayoutID(self):
        return self.navigationcontrol.LayoutID

    @property
    def Left(self):
        return self.navigationcontrol.Left

    @property
    def LeftPadding(self):
        return self.navigationcontrol.LeftPadding

    @property
    def LineSpacing(self):
        return self.navigationcontrol.LineSpacing

    @property
    def Name(self):
        return self.navigationcontrol.Name

    @property
    def OldBorderStyle(self):
        return self.navigationcontrol.OldBorderStyle

    @property
    def OldValue(self):
        return self.navigationcontrol.OldValue

    @property
    def OnClick(self):
        return self.navigationcontrol.OnClick

    @property
    def OnDblClick(self):
        return self.navigationcontrol.OnDblClick

    @property
    def OnGotFocus(self):
        return self.navigationcontrol.OnGotFocus

    @property
    def OnKeyDown(self):
        return self.navigationcontrol.OnKeyDown

    @property
    def OnKeyPress(self):
        return self.navigationcontrol.OnKeyPress

    @property
    def OnKeyUp(self):
        return self.navigationcontrol.OnKeyUp

    @property
    def OnLostFocus(self):
        return self.navigationcontrol.OnLostFocus

    @property
    def OnMouseDown(self):
        return self.navigationcontrol.OnMouseDown

    @property
    def OnMouseMove(self):
        return self.navigationcontrol.OnMouseMove

    @property
    def OnMouseUp(self):
        return self.navigationcontrol.OnMouseUp

    @property
    def Parent(self):
        return self.navigationcontrol.Parent

    @property
    def Properties(self):
        return Properties(self.navigationcontrol.Properties)

    @property
    def ReadingOrder(self):
        return self.navigationcontrol.ReadingOrder

    @property
    def RightPadding(self):
        return self.navigationcontrol.RightPadding

    @property
    def ScrollBarAlign(self):
        return self.navigationcontrol.ScrollBarAlign

    @property
    def Section(self):
        return self.navigationcontrol.Section

    @property
    def SelectedTab(self):
        return self.navigationcontrol.SelectedTab

    @property
    def ShortcutMenuBar(self):
        return self.navigationcontrol.ShortcutMenuBar

    @property
    def SmartTags(self):
        return SmartTags(self.navigationcontrol.SmartTags)

    @property
    def Span(self):
        return self.navigationcontrol.Span

    @property
    def SpecialEffect(self):
        return self.navigationcontrol.SpecialEffect

    @property
    def StatusBarText(self):
        return self.navigationcontrol.StatusBarText

    @property
    def SubForm(self):
        return self.navigationcontrol.SubForm

    @property
    def TabIndex(self):
        return self.navigationcontrol.TabIndex

    @property
    def Tabs(self):
        return self.navigationcontrol.Tabs

    @property
    def TabStop(self):
        return self.navigationcontrol.TabStop

    @property
    def Tag(self):
        return self.navigationcontrol.Tag

    @property
    def Top(self):
        return self.navigationcontrol.Top

    @property
    def TopPadding(self):
        return self.navigationcontrol.TopPadding

    @property
    def Value(self):
        return self.navigationcontrol.Value

    @property
    def VerticalAnchor(self):
        return self.navigationcontrol.VerticalAnchor

    @property
    def Visible(self):
        return self.navigationcontrol.Visible

    @Visible.setter
    def Visible(self, value):
        self.navigationcontrol.Visible = value

    @property
    def Width(self):
        return self.navigationcontrol.Width

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

    @property
    def Application(self):
        return self.objectframe.Application

    @property
    def AutoActivate(self):
        return self.objectframe.AutoActivate

    @property
    def BackColor(self):
        return self.objectframe.BackColor

    @property
    def BackShade(self):
        return self.objectframe.BackShade

    @property
    def BackStyle(self):
        return self.objectframe.BackStyle

    @property
    def BackThemeColorIndex(self):
        return self.objectframe.BackThemeColorIndex

    @property
    def BackTint(self):
        return self.objectframe.BackTint

    @property
    def BorderColor(self):
        return self.objectframe.BorderColor

    @property
    def BorderShade(self):
        return self.objectframe.BorderShade

    @property
    def BorderStyle(self):
        return self.objectframe.BorderStyle

    @property
    def BorderThemeColorIndex(self):
        return self.objectframe.BorderThemeColorIndex

    @property
    def BorderTint(self):
        return self.objectframe.BorderTint

    @property
    def BorderWidth(self):
        return self.objectframe.BorderWidth

    @property
    def BottomPadding(self):
        return self.objectframe.BottomPadding

    @property
    def Class(self):
        return self.objectframe.Class

    @property
    def ColumnCount(self):
        return self.objectframe.ColumnCount

    @property
    def ColumnHeads(self):
        return self.objectframe.ColumnHeads

    @property
    def Controls(self):
        return Controls(self.objectframe.Controls)

    @property
    def ControlTipText(self):
        return self.objectframe.ControlTipText

    @property
    def ControlType(self):
        return self.objectframe.ControlType

    @property
    def DisplayType(self):
        return self.objectframe.DisplayType

    @property
    def DisplayWhen(self):
        return self.objectframe.DisplayWhen

    @property
    def Enabled(self):
        return self.objectframe.Enabled

    @property
    def EventProcPrefix(self):
        return self.objectframe.EventProcPrefix

    @property
    def GridlineColor(self):
        return self.objectframe.GridlineColor

    @property
    def GridlineShade(self):
        return self.objectframe.GridlineShade

    @property
    def GridlineStyleBottom(self):
        return self.objectframe.GridlineStyleBottom

    @property
    def GridlineStyleLeft(self):
        return self.objectframe.GridlineStyleLeft

    @property
    def GridlineStyleRight(self):
        return self.objectframe.GridlineStyleRight

    @property
    def GridlineStyleTop(self):
        return self.objectframe.GridlineStyleTop

    @property
    def GridlineThemeColorIndex(self):
        return self.objectframe.GridlineThemeColorIndex

    @property
    def GridlineTint(self):
        return self.objectframe.GridlineTint

    @property
    def GridlineWidthBottom(self):
        return self.objectframe.GridlineWidthBottom

    @property
    def GridlineWidthLeft(self):
        return self.objectframe.GridlineWidthLeft

    @property
    def GridlineWidthRight(self):
        return self.objectframe.GridlineWidthRight

    @property
    def GridlineWidthTop(self):
        return self.objectframe.GridlineWidthTop

    @property
    def Height(self):
        return self.objectframe.Height

    @property
    def HelpContextId(self):
        return self.objectframe.HelpContextId

    @property
    def HorizontalAnchor(self):
        return self.objectframe.HorizontalAnchor

    @property
    def InSelection(self):
        return self.objectframe.InSelection

    @property
    def IsVisible(self):
        return self.objectframe.IsVisible

    @property
    def Item(self):
        return self.objectframe.Item

    @property
    def Layout(self):
        return AcLayoutType(self.objectframe.Layout)

    @property
    def LayoutID(self):
        return self.objectframe.LayoutID

    @property
    def Left(self):
        return self.objectframe.Left

    @property
    def LeftPadding(self):
        return self.objectframe.LeftPadding

    @property
    def LinkChildFields(self):
        return self.objectframe.LinkChildFields

    @property
    def LinkMasterFields(self):
        return self.objectframe.LinkMasterFields

    @property
    def Locked(self):
        return self.objectframe.Locked

    @property
    def Name(self):
        return self.objectframe.Name

    @property
    def Object(self):
        return self.objectframe.Object

    @property
    def ObjectPalette(self):
        return self.objectframe.ObjectPalette

    def ObjectVerbs(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.objectframe.ObjectVerbs(*args, **arguments)

    @property
    def ObjectVerbsCount(self):
        return self.objectframe.ObjectVerbsCount

    @property
    def OldBorderStyle(self):
        return self.objectframe.OldBorderStyle

    @property
    def OldValue(self):
        return self.objectframe.OldValue

    @property
    def OLEClass(self):
        return self.objectframe.OLEClass

    @property
    def OLEType(self):
        return self.objectframe.OLEType

    @property
    def OLETypeAllowed(self):
        return self.objectframe.OLETypeAllowed

    @property
    def OnClick(self):
        return self.objectframe.OnClick

    @property
    def OnDblClick(self):
        return self.objectframe.OnDblClick

    @property
    def OnEnter(self):
        return self.objectframe.OnEnter

    @property
    def OnExit(self):
        return self.objectframe.OnExit

    @property
    def OnGotFocus(self):
        return self.objectframe.OnGotFocus

    @property
    def OnLostFocus(self):
        return self.objectframe.OnLostFocus

    @property
    def OnMouseDown(self):
        return self.objectframe.OnMouseDown

    @property
    def OnMouseMove(self):
        return self.objectframe.OnMouseMove

    @property
    def OnMouseUp(self):
        return self.objectframe.OnMouseUp

    @property
    def OnUpdated(self):
        return self.objectframe.OnUpdated

    @property
    def Parent(self):
        return self.objectframe.Parent

    @property
    def Properties(self):
        return Properties(self.objectframe.Properties)

    @property
    def RightPadding(self):
        return self.objectframe.RightPadding

    @property
    def RowSource(self):
        return self.objectframe.RowSource

    @property
    def RowSourceType(self):
        return self.objectframe.RowSourceType

    @property
    def Scaling(self):
        return self.objectframe.Scaling

    @property
    def Section(self):
        return self.objectframe.Section

    @property
    def ShortcutMenuBar(self):
        return self.objectframe.ShortcutMenuBar

    @property
    def SizeMode(self):
        return self.objectframe.SizeMode

    @property
    def SourceDoc(self):
        return self.objectframe.SourceDoc

    @property
    def SourceItem(self):
        return self.objectframe.SourceItem

    @property
    def SourceObject(self):
        return self.objectframe.SourceObject

    @property
    def SpecialEffect(self):
        return self.objectframe.SpecialEffect

    @property
    def StatusBarText(self):
        return self.objectframe.StatusBarText

    @property
    def TabIndex(self):
        return self.objectframe.TabIndex

    @property
    def TabStop(self):
        return self.objectframe.TabStop

    @property
    def Tag(self):
        return self.objectframe.Tag

    @property
    def Top(self):
        return self.objectframe.Top

    @property
    def TopPadding(self):
        return self.objectframe.TopPadding

    @property
    def UpdateMethod(self):
        return self.objectframe.UpdateMethod

    @property
    def UpdateOptions(self):
        return self.objectframe.UpdateOptions

    @property
    def VarOleObject(self):
        return self.objectframe.VarOleObject

    @property
    def Verb(self):
        return self.objectframe.Verb

    @property
    def VerticalAnchor(self):
        return self.objectframe.VerticalAnchor

    @property
    def Visible(self):
        return self.objectframe.Visible

    @Visible.setter
    def Visible(self, value):
        self.objectframe.Visible = value

    @property
    def Width(self):
        return self.objectframe.Width

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
        return self.operations.Item(*args, **arguments)

    @property
    def Parent(self):
        return self.operations.Parent

class OptionButton:

    def __init__(self, optionbutton=None):
        self.optionbutton = optionbutton

    @property
    def AddColon(self):
        return self.optionbutton.AddColon

    @property
    def Application(self):
        return self.optionbutton.Application

    @property
    def AutoLabel(self):
        return self.optionbutton.AutoLabel

    @property
    def BorderColor(self):
        return self.optionbutton.BorderColor

    @property
    def BorderShade(self):
        return self.optionbutton.BorderShade

    @property
    def BorderStyle(self):
        return self.optionbutton.BorderStyle

    @property
    def BorderThemeColorIndex(self):
        return self.optionbutton.BorderThemeColorIndex

    @property
    def BorderTint(self):
        return self.optionbutton.BorderTint

    @property
    def BorderWidth(self):
        return self.optionbutton.BorderWidth

    @property
    def BottomPadding(self):
        return self.optionbutton.BottomPadding

    @property
    def ColumnHidden(self):
        return self.optionbutton.ColumnHidden

    @property
    def ColumnOrder(self):
        return self.optionbutton.ColumnOrder

    @property
    def ColumnWidth(self):
        return self.optionbutton.ColumnWidth

    @property
    def Controls(self):
        return Controls(self.optionbutton.Controls)

    @property
    def ControlSource(self):
        return self.optionbutton.ControlSource

    @property
    def ControlTipText(self):
        return self.optionbutton.ControlTipText

    @property
    def ControlType(self):
        return self.optionbutton.ControlType

    @property
    def DefaultValue(self):
        return self.optionbutton.DefaultValue

    @property
    def DisplayWhen(self):
        return self.optionbutton.DisplayWhen

    @property
    def Enabled(self):
        return self.optionbutton.Enabled

    @property
    def EventProcPrefix(self):
        return self.optionbutton.EventProcPrefix

    @property
    def GridlineColor(self):
        return self.optionbutton.GridlineColor

    @property
    def GridlineShade(self):
        return self.optionbutton.GridlineShade

    @property
    def GridlineStyleBottom(self):
        return self.optionbutton.GridlineStyleBottom

    @property
    def GridlineStyleLeft(self):
        return self.optionbutton.GridlineStyleLeft

    @property
    def GridlineStyleRight(self):
        return self.optionbutton.GridlineStyleRight

    @property
    def GridlineStyleTop(self):
        return self.optionbutton.GridlineStyleTop

    @property
    def GridlineThemeColorIndex(self):
        return self.optionbutton.GridlineThemeColorIndex

    @property
    def GridlineTint(self):
        return self.optionbutton.GridlineTint

    @property
    def GridlineWidthBottom(self):
        return self.optionbutton.GridlineWidthBottom

    @property
    def GridlineWidthLeft(self):
        return self.optionbutton.GridlineWidthLeft

    @property
    def GridlineWidthRight(self):
        return self.optionbutton.GridlineWidthRight

    @property
    def GridlineWidthTop(self):
        return self.optionbutton.GridlineWidthTop

    @property
    def Height(self):
        return self.optionbutton.Height

    @property
    def HelpContextId(self):
        return self.optionbutton.HelpContextId

    @property
    def HideDuplicates(self):
        return self.optionbutton.HideDuplicates

    @property
    def HorizontalAnchor(self):
        return self.optionbutton.HorizontalAnchor

    @property
    def InSelection(self):
        return self.optionbutton.InSelection

    @property
    def IsVisible(self):
        return self.optionbutton.IsVisible

    @property
    def LabelAlign(self):
        return self.optionbutton.LabelAlign

    @property
    def LabelX(self):
        return self.optionbutton.LabelX

    @property
    def LabelY(self):
        return self.optionbutton.LabelY

    @property
    def Layout(self):
        return AcLayoutType(self.optionbutton.Layout)

    @property
    def LayoutID(self):
        return self.optionbutton.LayoutID

    @property
    def Left(self):
        return self.optionbutton.Left

    @property
    def LeftPadding(self):
        return self.optionbutton.LeftPadding

    @property
    def Locked(self):
        return self.optionbutton.Locked

    @property
    def Name(self):
        return self.optionbutton.Name

    @property
    def OldBorderStyle(self):
        return self.optionbutton.OldBorderStyle

    @property
    def OldValue(self):
        return self.optionbutton.OldValue

    @property
    def OnClick(self):
        return self.optionbutton.OnClick

    @property
    def OnDblClick(self):
        return self.optionbutton.OnDblClick

    @property
    def OnEnter(self):
        return self.optionbutton.OnEnter

    @property
    def OnExit(self):
        return self.optionbutton.OnExit

    @property
    def OnGotFocus(self):
        return self.optionbutton.OnGotFocus

    @property
    def OnKeyDown(self):
        return self.optionbutton.OnKeyDown

    @property
    def OnKeyPress(self):
        return self.optionbutton.OnKeyPress

    @property
    def OnKeyUp(self):
        return self.optionbutton.OnKeyUp

    @property
    def OnLostFocus(self):
        return self.optionbutton.OnLostFocus

    @property
    def OnMouseDown(self):
        return self.optionbutton.OnMouseDown

    @property
    def OnMouseMove(self):
        return self.optionbutton.OnMouseMove

    @property
    def OnMouseUp(self):
        return self.optionbutton.OnMouseUp

    @property
    def OptionValue(self):
        return self.optionbutton.OptionValue

    @property
    def Parent(self):
        return self.optionbutton.Parent

    @property
    def Properties(self):
        return Properties(self.optionbutton.Properties)

    @property
    def ReadingOrder(self):
        return self.optionbutton.ReadingOrder

    @property
    def RightPadding(self):
        return self.optionbutton.RightPadding

    @property
    def Section(self):
        return self.optionbutton.Section

    @property
    def ShortcutMenuBar(self):
        return self.optionbutton.ShortcutMenuBar

    @property
    def SpecialEffect(self):
        return self.optionbutton.SpecialEffect

    @property
    def StatusBarText(self):
        return self.optionbutton.StatusBarText

    @property
    def TabIndex(self):
        return self.optionbutton.TabIndex

    @property
    def TabStop(self):
        return self.optionbutton.TabStop

    @property
    def Tag(self):
        return self.optionbutton.Tag

    @property
    def Top(self):
        return self.optionbutton.Top

    @property
    def TopPadding(self):
        return self.optionbutton.TopPadding

    @property
    def TripleState(self):
        return self.optionbutton.TripleState

    @property
    def ValidationRule(self):
        return self.optionbutton.ValidationRule

    @property
    def ValidationText(self):
        return self.optionbutton.ValidationText

    @property
    def Value(self):
        return self.optionbutton.Value

    @property
    def VerticalAnchor(self):
        return self.optionbutton.VerticalAnchor

    @property
    def Visible(self):
        return self.optionbutton.Visible

    @Visible.setter
    def Visible(self, value):
        self.optionbutton.Visible = value

    @property
    def Width(self):
        return self.optionbutton.Width

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

    @property
    def Application(self):
        return self.optiongroup.Application

    @property
    def AutoLabel(self):
        return self.optiongroup.AutoLabel

    @property
    def BackColor(self):
        return self.optiongroup.BackColor

    @property
    def BackShade(self):
        return self.optiongroup.BackShade

    @property
    def BackStyle(self):
        return self.optiongroup.BackStyle

    @property
    def BackThemeColorIndex(self):
        return self.optiongroup.BackThemeColorIndex

    @property
    def BackTint(self):
        return self.optiongroup.BackTint

    @property
    def BorderColor(self):
        return self.optiongroup.BorderColor

    @property
    def BorderShade(self):
        return self.optiongroup.BorderShade

    @property
    def BorderStyle(self):
        return self.optiongroup.BorderStyle

    @property
    def BorderThemeColorIndex(self):
        return self.optiongroup.BorderThemeColorIndex

    @property
    def BorderTint(self):
        return self.optiongroup.BorderTint

    @property
    def BorderWidth(self):
        return self.optiongroup.BorderWidth

    @property
    def ColumnHidden(self):
        return self.optiongroup.ColumnHidden

    @property
    def ColumnOrder(self):
        return self.optiongroup.ColumnOrder

    @property
    def ColumnWidth(self):
        return self.optiongroup.ColumnWidth

    @property
    def Controls(self):
        return Controls(self.optiongroup.Controls)

    @property
    def ControlSource(self):
        return self.optiongroup.ControlSource

    @property
    def ControlTipText(self):
        return self.optiongroup.ControlTipText

    @property
    def ControlType(self):
        return self.optiongroup.ControlType

    @property
    def DefaultValue(self):
        return self.optiongroup.DefaultValue

    @property
    def DisplayWhen(self):
        return self.optiongroup.DisplayWhen

    @property
    def Enabled(self):
        return self.optiongroup.Enabled

    @property
    def EventProcPrefix(self):
        return self.optiongroup.EventProcPrefix

    @property
    def Height(self):
        return self.optiongroup.Height

    @property
    def HelpContextId(self):
        return self.optiongroup.HelpContextId

    @property
    def HideDuplicates(self):
        return self.optiongroup.HideDuplicates

    @property
    def HorizontalAnchor(self):
        return self.optiongroup.HorizontalAnchor

    @property
    def InSelection(self):
        return self.optiongroup.InSelection

    @property
    def IsVisible(self):
        return self.optiongroup.IsVisible

    @property
    def LabelAlign(self):
        return self.optiongroup.LabelAlign

    @property
    def LabelX(self):
        return self.optiongroup.LabelX

    @property
    def LabelY(self):
        return self.optiongroup.LabelY

    @property
    def Left(self):
        return self.optiongroup.Left

    @property
    def Locked(self):
        return self.optiongroup.Locked

    @property
    def Name(self):
        return self.optiongroup.Name

    @property
    def OldBorderStyle(self):
        return self.optiongroup.OldBorderStyle

    @property
    def OldValue(self):
        return self.optiongroup.OldValue

    @property
    def OnClick(self):
        return self.optiongroup.OnClick

    @property
    def OnDblClick(self):
        return self.optiongroup.OnDblClick

    @property
    def OnEnter(self):
        return self.optiongroup.OnEnter

    @property
    def OnExit(self):
        return self.optiongroup.OnExit

    @property
    def OnMouseDown(self):
        return self.optiongroup.OnMouseDown

    @property
    def OnMouseMove(self):
        return self.optiongroup.OnMouseMove

    @property
    def OnMouseUp(self):
        return self.optiongroup.OnMouseUp

    @property
    def Parent(self):
        return self.optiongroup.Parent

    @property
    def Properties(self):
        return Properties(self.optiongroup.Properties)

    @property
    def Section(self):
        return self.optiongroup.Section

    @property
    def ShortcutMenuBar(self):
        return self.optiongroup.ShortcutMenuBar

    @property
    def SpecialEffect(self):
        return self.optiongroup.SpecialEffect

    @property
    def StatusBarText(self):
        return self.optiongroup.StatusBarText

    @property
    def TabIndex(self):
        return self.optiongroup.TabIndex

    @property
    def TabStop(self):
        return self.optiongroup.TabStop

    @property
    def Tag(self):
        return self.optiongroup.Tag

    @property
    def Top(self):
        return self.optiongroup.Top

    @property
    def ValidationRule(self):
        return self.optiongroup.ValidationRule

    @property
    def ValidationText(self):
        return self.optiongroup.ValidationText

    @property
    def Value(self):
        return self.optiongroup.Value

    @property
    def VerticalAnchor(self):
        return self.optiongroup.VerticalAnchor

    @property
    def Visible(self):
        return self.optiongroup.Visible

    @Visible.setter
    def Visible(self, value):
        self.optiongroup.Visible = value

    @property
    def Width(self):
        return self.optiongroup.Width

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

    @property
    def Controls(self):
        return Controls(self.page.Controls)

    @property
    def ControlTipText(self):
        return self.page.ControlTipText

    @property
    def ControlType(self):
        return self.page.ControlType

    @property
    def Enabled(self):
        return self.page.Enabled

    @property
    def EventProcPrefix(self):
        return self.page.EventProcPrefix

    @property
    def Height(self):
        return self.page.Height

    @property
    def HelpContextId(self):
        return self.page.HelpContextId

    @property
    def InSelection(self):
        return self.page.InSelection

    @property
    def IsVisible(self):
        return self.page.IsVisible

    @property
    def Left(self):
        return self.page.Left

    @property
    def Name(self):
        return self.page.Name

    @property
    def OnClick(self):
        return self.page.OnClick

    @property
    def OnDblClick(self):
        return self.page.OnDblClick

    @property
    def OnMouseDown(self):
        return self.page.OnMouseDown

    @property
    def OnMouseMove(self):
        return self.page.OnMouseMove

    @property
    def OnMouseUp(self):
        return self.page.OnMouseUp

    @property
    def PageIndex(self):
        return self.page.PageIndex

    @property
    def Parent(self):
        return self.page.Parent

    @property
    def Picture(self):
        return self.page.Picture

    @property
    def PictureData(self):
        return self.page.PictureData

    @property
    def PictureType(self):
        return self.page.PictureType

    @property
    def Properties(self):
        return Properties(self.page.Properties)

    @property
    def Section(self):
        return self.page.Section

    @property
    def ShortcutMenuBar(self):
        return self.page.ShortcutMenuBar

    @property
    def StatusBarText(self):
        return self.page.StatusBarText

    @property
    def Tag(self):
        return self.page.Tag

    @property
    def Top(self):
        return self.page.Top

    @property
    def Visible(self):
        return self.page.Visible

    @Visible.setter
    def Visible(self, value):
        self.page.Visible = value

    @property
    def Width(self):
        return self.page.Width

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

    @property
    def EventProcPrefix(self):
        return self.pagebreak.EventProcPrefix

    @property
    def InSelection(self):
        return self.pagebreak.InSelection

    @property
    def IsVisible(self):
        return self.pagebreak.IsVisible

    @property
    def Left(self):
        return self.pagebreak.Left

    @property
    def Name(self):
        return self.pagebreak.Name

    @property
    def Parent(self):
        return self.pagebreak.Parent

    @property
    def Properties(self):
        return Properties(self.pagebreak.Properties)

    @property
    def Section(self):
        return self.pagebreak.Section

    @property
    def Tag(self):
        return self.pagebreak.Tag

    @property
    def Top(self):
        return self.pagebreak.Top

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
        return self.pages.Item(*args, **arguments)

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

    @property
    def DefaultSize(self):
        return self.printer.DefaultSize

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

    @property
    def Orientation(self):
        return self.printer.Orientation

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

    @property
    def RowSpacing(self):
        return self.printer.RowSpacing

    @RowSpacing.setter
    def RowSpacing(self, value):
        self.printer.RowSpacing = value

    @property
    def TopMargin(self):
        return self.printer.TopMargin

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
        return self.printers.Item(*args, **arguments)

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
        return self.properties.Item(*args, **arguments)

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

    @property
    def BackShade(self):
        return self.rectangle.BackShade

    @property
    def BackStyle(self):
        return self.rectangle.BackStyle

    @property
    def BackThemeColorIndex(self):
        return self.rectangle.BackThemeColorIndex

    @property
    def BackTint(self):
        return self.rectangle.BackTint

    @property
    def BorderColor(self):
        return self.rectangle.BorderColor

    @property
    def BorderShade(self):
        return self.rectangle.BorderShade

    @property
    def BorderStyle(self):
        return self.rectangle.BorderStyle

    @property
    def BorderThemeColorIndex(self):
        return self.rectangle.BorderThemeColorIndex

    @property
    def BorderTint(self):
        return self.rectangle.BorderTint

    @property
    def BorderWidth(self):
        return self.rectangle.BorderWidth

    @property
    def ControlType(self):
        return self.rectangle.ControlType

    @property
    def DisplayWhen(self):
        return self.rectangle.DisplayWhen

    @property
    def EventProcPrefix(self):
        return self.rectangle.EventProcPrefix

    @property
    def Height(self):
        return self.rectangle.Height

    @property
    def HorizontalAnchor(self):
        return self.rectangle.HorizontalAnchor

    @property
    def InSelection(self):
        return self.rectangle.InSelection

    @property
    def IsVisible(self):
        return self.rectangle.IsVisible

    @property
    def Left(self):
        return self.rectangle.Left

    @property
    def Name(self):
        return self.rectangle.Name

    @property
    def OldBorderStyle(self):
        return self.rectangle.OldBorderStyle

    @property
    def OnClick(self):
        return self.rectangle.OnClick

    @property
    def OnDblClick(self):
        return self.rectangle.OnDblClick

    @property
    def OnMouseDown(self):
        return self.rectangle.OnMouseDown

    @property
    def OnMouseMove(self):
        return self.rectangle.OnMouseMove

    @property
    def OnMouseUp(self):
        return self.rectangle.OnMouseUp

    @property
    def Parent(self):
        return self.rectangle.Parent

    @property
    def Properties(self):
        return Properties(self.rectangle.Properties)

    @property
    def Section(self):
        return self.rectangle.Section

    @property
    def SpecialEffect(self):
        return self.rectangle.SpecialEffect

    @property
    def Tag(self):
        return self.rectangle.Tag

    @property
    def Top(self):
        return self.rectangle.Top

    @property
    def VerticalAnchor(self):
        return self.rectangle.VerticalAnchor

    @property
    def Visible(self):
        return self.rectangle.Visible

    @Visible.setter
    def Visible(self, value):
        self.rectangle.Visible = value

    @property
    def Width(self):
        return self.rectangle.Width

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

    @property
    def AllowReportView(self):
        return self.report.AllowReportView

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

    @property
    def Caption(self):
        return self.report.Caption

    @property
    def CloseButton(self):
        return self.report.CloseButton

    @property
    def ControlBox(self):
        return self.report.ControlBox

    @property
    def Controls(self):
        return Controls(self.report.Controls)

    @property
    def Count(self):
        return self.report.Count

    @property
    def CurrentRecord(self):
        return self.report.CurrentRecord

    @property
    def CurrentView(self):
        return self.report.CurrentView

    @property
    def CurrentX(self):
        return self.report.CurrentX

    @property
    def CurrentY(self):
        return self.report.CurrentY

    @property
    def Cycle(self):
        return self.report.Cycle

    @property
    def DateGrouping(self):
        return self.report.DateGrouping

    def DefaultControl(self, *args, ControlType=None):
        arguments = {"ControlType": ControlType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.report.DefaultControl(*args, **arguments)

    @property
    def DefaultView(self):
        return self.report.DefaultView

    @property
    def Dirty(self):
        return self.report.Dirty

    @property
    def DisplayOnSharePointSite(self):
        return self.report.DisplayOnSharePointSite

    @property
    def DrawMode(self):
        return self.report.DrawMode

    @property
    def DrawStyle(self):
        return self.report.DrawStyle

    @property
    def DrawWidth(self):
        return self.report.DrawWidth

    @property
    def FastLaserPrinting(self):
        return self.report.FastLaserPrinting

    @property
    def FillColor(self):
        return self.report.FillColor

    @property
    def FillStyle(self):
        return self.report.FillStyle

    @property
    def Filter(self):
        return self.report.Filter

    @property
    def FilterOn(self):
        return self.report.FilterOn

    @property
    def FilterOnLoad(self):
        return self.report.FilterOnLoad

    @property
    def FitToPage(self):
        return self.report.FitToPage

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

    @property
    def FormatCount(self):
        return self.report.FormatCount

    @property
    def GridX(self):
        return self.report.GridX

    @property
    def GridY(self):
        return self.report.GridY

    def GroupLevel(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.report.GroupLevel(*args, **arguments)

    @property
    def GrpKeepTogether(self):
        return self.report.GrpKeepTogether

    @property
    def HasData(self):
        return self.report.HasData

    @property
    def HasModule(self):
        return self.report.HasModule

    @property
    def Height(self):
        return self.report.Height

    @property
    def HelpContextId(self):
        return self.report.HelpContextId

    @property
    def HelpFile(self):
        return self.report.HelpFile

    @property
    def Hwnd(self):
        return self.report.Hwnd

    @property
    def KeyPreview(self):
        return self.report.KeyPreview

    @property
    def LayoutForPrint(self):
        return self.report.LayoutForPrint

    @property
    def Left(self):
        return self.report.Left

    @property
    def MenuBar(self):
        return self.report.MenuBar

    @property
    def MinMaxButtons(self):
        return self.report.MinMaxButtons

    @property
    def Modal(self):
        return self.report.Modal

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

    @property
    def Name(self):
        return self.report.Name

    @property
    def NextRecord(self):
        return self.report.NextRecord

    @property
    def OnActivate(self):
        return self.report.OnActivate

    @property
    def OnApplyFilter(self):
        return self.report.OnApplyFilter

    @property
    def OnClick(self):
        return self.report.OnClick

    @property
    def OnClose(self):
        return self.report.OnClose

    @property
    def OnCurrent(self):
        return self.report.OnCurrent

    @property
    def OnDblClick(self):
        return self.report.OnDblClick

    @property
    def OnDeactivate(self):
        return self.report.OnDeactivate

    @property
    def OnError(self):
        return self.report.OnError

    @property
    def OnFilter(self):
        return self.report.OnFilter

    @property
    def OnGotFocus(self):
        return self.report.OnGotFocus

    @property
    def OnKeyDown(self):
        return self.report.OnKeyDown

    @property
    def OnKeyPress(self):
        return self.report.OnKeyPress

    @property
    def OnKeyUp(self):
        return self.report.OnKeyUp

    @property
    def OnLoad(self):
        return self.report.OnLoad

    @property
    def OnLostFocus(self):
        return self.report.OnLostFocus

    @property
    def OnMouseDown(self):
        return self.report.OnMouseDown

    @property
    def OnMouseMove(self):
        return self.report.OnMouseMove

    @property
    def OnMouseUp(self):
        return self.report.OnMouseUp

    @property
    def OnNoData(self):
        return self.report.OnNoData

    @property
    def OnOpen(self):
        return self.report.OnOpen

    @property
    def OnPage(self):
        return self.report.OnPage

    @property
    def OnResize(self):
        return self.report.OnResize

    @property
    def OnTimer(self):
        return self.report.OnTimer

    @property
    def OnUnload(self):
        return self.report.OnUnload

    @property
    def OpenArgs(self):
        return self.report.OpenArgs

    @property
    def OrderBy(self):
        return self.report.OrderBy

    @property
    def OrderByOn(self):
        return self.report.OrderByOn

    @property
    def OrderByOnLoad(self):
        return self.report.OrderByOnLoad

    @property
    def Orientation(self):
        return self.report.Orientation

    @property
    def Page(self):
        return self.report.Page

    @property
    def PageFooter(self):
        return self.report.PageFooter

    @property
    def PageHeader(self):
        return self.report.PageHeader

    @property
    def Pages(self):
        return self.report.Pages

    @property
    def Painting(self):
        return self.report.Painting

    @property
    def PaintPalette(self):
        return self.report.PaintPalette

    @property
    def PaletteSource(self):
        return self.report.PaletteSource

    @property
    def Parent(self):
        return self.report.Parent

    @property
    def Picture(self):
        return self.report.Picture

    @property
    def PictureAlignment(self):
        return self.report.PictureAlignment

    @property
    def PictureData(self):
        return self.report.PictureData

    @property
    def PicturePages(self):
        return self.report.PicturePages

    @property
    def PicturePalette(self):
        return self.report.PicturePalette

    @property
    def PictureSizeMode(self):
        return self.report.PictureSizeMode

    @property
    def PictureTiling(self):
        return self.report.PictureTiling

    @property
    def PictureType(self):
        return self.report.PictureType

    @property
    def PopUp(self):
        return self.report.PopUp

    @property
    def PrintCount(self):
        return self.report.PrintCount

    @property
    def Printer(self):
        return Printer(self.report.Printer)

    @Printer.setter
    def Printer(self, value):
        self.report.Printer = value

    @property
    def PrintSection(self):
        return self.report.PrintSection

    @property
    def Properties(self):
        return Properties(self.report.Properties)

    @property
    def PrtDevMode(self):
        return self.report.PrtDevMode

    @property
    def PrtDevNames(self):
        return self.report.PrtDevNames

    @property
    def PrtMip(self):
        return self.report.PrtMip

    @property
    def RecordLocks(self):
        return self.report.RecordLocks

    @property
    def Recordset(self):
        return self.report.Recordset

    @Recordset.setter
    def Recordset(self, value):
        self.report.Recordset = value

    @property
    def RecordSource(self):
        return self.report.RecordSource

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

    @property
    def ScaleHeight(self):
        return self.report.ScaleHeight

    @property
    def ScaleLeft(self):
        return self.report.ScaleLeft

    @property
    def ScaleMode(self):
        return self.report.ScaleMode

    @property
    def ScaleTop(self):
        return self.report.ScaleTop

    @property
    def ScaleWidth(self):
        return self.report.ScaleWidth

    @property
    def ScrollBars(self):
        return self.report.ScrollBars

    def Section(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.report.Section(*args, **arguments)

    @property
    def ServerFilter(self):
        return self.report.ServerFilter

    @property
    def ShortcutMenuBar(self):
        return self.report.ShortcutMenuBar

    @property
    def ShowPageMargins(self):
        return self.report.ShowPageMargins

    @property
    def Tag(self):
        return self.report.Tag

    @property
    def TimerInterval(self):
        return self.report.TimerInterval

    @property
    def Toolbar(self):
        return self.report.Toolbar

    @property
    def Top(self):
        return self.report.Top

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

    def Circle(self, *args, Step_ (_x=None, y=None):
        arguments = {"Step_ (_x": Step_ (_x, "y": y}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.report.Circle(*args, **arguments)

    def Line(self, *args, Step_ (_x1=None, y1=None):
        arguments = {"Step_ (_x1": Step_ (_x1, "y1": y1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.report.Line(*args, **arguments)

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
        return self.reports.Item(*args, **arguments)

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
        return self.returnvars.Item(*args, **arguments)

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

    @property
    def AlternateBackShade(self):
        return self.section.AlternateBackShade

    @property
    def AlternateBackThemeColorIndex(self):
        return self.section.AlternateBackThemeColorIndex

    @property
    def AlternateBackTint(self):
        return self.section.AlternateBackTint

    @property
    def Application(self):
        return self.section.Application

    @property
    def AutoHeight(self):
        return self.section.AutoHeight

    @property
    def BackColor(self):
        return self.section.BackColor

    @property
    def BackShade(self):
        return self.section.BackShade

    @property
    def BackThemeColorIndex(self):
        return self.section.BackThemeColorIndex

    @property
    def BackTint(self):
        return self.section.BackTint

    @property
    def CanGrow(self):
        return self.section.CanGrow

    @property
    def CanShrink(self):
        return self.section.CanShrink

    @property
    def Controls(self):
        return Controls(self.section.Controls)

    @property
    def DisplayWhen(self):
        return self.section.DisplayWhen

    @property
    def EventProcPrefix(self):
        return self.section.EventProcPrefix

    @property
    def ForceNewPage(self):
        return self.section.ForceNewPage

    @property
    def HasContinued(self):
        return self.section.HasContinued

    @property
    def Height(self):
        return self.section.Height

    @property
    def InSelection(self):
        return self.section.InSelection

    @property
    def KeepTogether(self):
        return self.section.KeepTogether

    @property
    def Name(self):
        return self.section.Name

    @property
    def NewRowOrCol(self):
        return self.section.NewRowOrCol

    @property
    def OnClick(self):
        return self.section.OnClick

    @property
    def OnDblClick(self):
        return self.section.OnDblClick

    @property
    def OnFormat(self):
        return self.section.OnFormat

    @property
    def OnMouseDown(self):
        return self.section.OnMouseDown

    @property
    def OnMouseMove(self):
        return self.section.OnMouseMove

    @property
    def OnMouseUp(self):
        return self.section.OnMouseUp

    @property
    def OnPaint(self):
        return self.section.OnPaint

    @property
    def OnPrint(self):
        return self.section.OnPrint

    @property
    def OnRetreat(self):
        return self.section.OnRetreat

    @property
    def Parent(self):
        return self.section.Parent

    @property
    def Properties(self):
        return Properties(self.section.Properties)

    @property
    def RepeatSection(self):
        return self.section.RepeatSection

    @property
    def SpecialEffect(self):
        return self.section.SpecialEffect

    @property
    def Tag(self):
        return self.section.Tag

    @property
    def Visible(self):
        return self.section.Visible

    @Visible.setter
    def Visible(self, value):
        self.section.Visible = value

    @property
    def WillContinue(self):
        return self.section.WillContinue

    def SetTabOrder(self):
        self.section.SetTabOrder()

class SharedResource:

    def __init__(self, sharedresource=None):
        self.sharedresource = sharedresource

    @property
    def Name(self):
        return self.sharedresource.Name

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
        return self.sharedresources.Item(*args, **arguments)

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
        return self.smarttagactions.Item(*args, **arguments)

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
        return self.smarttagproperties.Item(*args, **arguments)

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

    @property
    def Value(self):
        return self.smarttagproperty.Value

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
        return self.smarttags.Item(*args, **arguments)

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

    @property
    def Application(self):
        return self.subform.Application

    @property
    def AutoLabel(self):
        return self.subform.AutoLabel

    @property
    def BorderColor(self):
        return self.subform.BorderColor

    @property
    def BorderShade(self):
        return self.subform.BorderShade

    @property
    def BorderStyle(self):
        return self.subform.BorderStyle

    @property
    def BorderThemeColorIndex(self):
        return self.subform.BorderThemeColorIndex

    @property
    def BorderTint(self):
        return self.subform.BorderTint

    @property
    def BorderWidth(self):
        return self.subform.BorderWidth

    @property
    def BottomPadding(self):
        return self.subform.BottomPadding

    @property
    def CanGrow(self):
        return self.subform.CanGrow

    @property
    def CanShrink(self):
        return self.subform.CanShrink

    @property
    def Controls(self):
        return Controls(self.subform.Controls)

    @property
    def ControlType(self):
        return self.subform.ControlType

    @property
    def DisplayWhen(self):
        return self.subform.DisplayWhen

    @property
    def Enabled(self):
        return self.subform.Enabled

    @property
    def EventProcPrefix(self):
        return self.subform.EventProcPrefix

    @property
    def FilterOnEmptyMaster(self):
        return self.subform.FilterOnEmptyMaster

    @property
    def Form(self):
        return self.subform.Form

    @property
    def GridlineColor(self):
        return self.subform.GridlineColor

    @property
    def GridlineShade(self):
        return self.subform.GridlineShade

    @property
    def GridlineStyleBottom(self):
        return self.subform.GridlineStyleBottom

    @property
    def GridlineStyleLeft(self):
        return self.subform.GridlineStyleLeft

    @property
    def GridlineStyleRight(self):
        return self.subform.GridlineStyleRight

    @property
    def GridlineStyleTop(self):
        return self.subform.GridlineStyleTop

    @property
    def GridlineThemeColorIndex(self):
        return self.subform.GridlineThemeColorIndex

    @property
    def GridlineTint(self):
        return self.subform.GridlineTint

    @property
    def GridlineWidthBottom(self):
        return self.subform.GridlineWidthBottom

    @property
    def GridlineWidthLeft(self):
        return self.subform.GridlineWidthLeft

    @property
    def GridlineWidthRight(self):
        return self.subform.GridlineWidthRight

    @property
    def GridlineWidthTop(self):
        return self.subform.GridlineWidthTop

    @property
    def Height(self):
        return self.subform.Height

    @property
    def HorizontalAnchor(self):
        return self.subform.HorizontalAnchor

    @property
    def InSelection(self):
        return self.subform.InSelection

    @property
    def IsVisible(self):
        return self.subform.IsVisible

    @property
    def LabelAlign(self):
        return self.subform.LabelAlign

    @property
    def LabelX(self):
        return self.subform.LabelX

    @property
    def LabelY(self):
        return self.subform.LabelY

    @property
    def Layout(self):
        return AcLayoutType(self.subform.Layout)

    @property
    def LayoutID(self):
        return self.subform.LayoutID

    @property
    def Left(self):
        return self.subform.Left

    @property
    def LeftPadding(self):
        return self.subform.LeftPadding

    @property
    def LinkChildFields(self):
        return self.subform.LinkChildFields

    @property
    def LinkMasterFields(self):
        return self.subform.LinkMasterFields

    @property
    def Locked(self):
        return self.subform.Locked

    @property
    def Name(self):
        return self.subform.Name

    @property
    def OldBorderStyle(self):
        return self.subform.OldBorderStyle

    @property
    def OnEnter(self):
        return self.subform.OnEnter

    @property
    def OnExit(self):
        return self.subform.OnExit

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

    @property
    def Section(self):
        return self.subform.Section

    @property
    def SourceObject(self):
        return self.subform.SourceObject

    @property
    def SpecialEffect(self):
        return self.subform.SpecialEffect

    @property
    def StatusBarText(self):
        return self.subform.StatusBarText

    @property
    def TabIndex(self):
        return self.subform.TabIndex

    @property
    def TabStop(self):
        return self.subform.TabStop

    @property
    def Tag(self):
        return self.subform.Tag

    @property
    def Top(self):
        return self.subform.Top

    @property
    def TopPadding(self):
        return self.subform.TopPadding

    @property
    def VerticalAnchor(self):
        return self.subform.VerticalAnchor

    @property
    def Visible(self):
        return self.subform.Visible

    @Visible.setter
    def Visible(self, value):
        self.subform.Visible = value

    @property
    def Width(self):
        return self.subform.Width

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

    @property
    def BackShade(self):
        return self.tabcontrol.BackShade

    @property
    def BackStyle(self):
        return self.tabcontrol.BackStyle

    @property
    def BackThemeColorIndex(self):
        return self.tabcontrol.BackThemeColorIndex

    @property
    def BackTint(self):
        return self.tabcontrol.BackTint

    @property
    def BorderColor(self):
        return self.tabcontrol.BorderColor

    @property
    def BorderShade(self):
        return self.tabcontrol.BorderShade

    @property
    def BorderStyle(self):
        return self.tabcontrol.BorderStyle

    @property
    def BorderThemeColorIndex(self):
        return self.tabcontrol.BorderThemeColorIndex

    @property
    def BorderTint(self):
        return self.tabcontrol.BorderTint

    @property
    def BottomPadding(self):
        return self.tabcontrol.BottomPadding

    @property
    def ControlType(self):
        return self.tabcontrol.ControlType

    @property
    def DisplayWhen(self):
        return self.tabcontrol.DisplayWhen

    @property
    def Enabled(self):
        return self.tabcontrol.Enabled

    @property
    def EventProcPrefix(self):
        return self.tabcontrol.EventProcPrefix

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

    @property
    def ForeColor(self):
        return self.tabcontrol.ForeColor

    @property
    def ForeShade(self):
        return self.tabcontrol.ForeShade

    @property
    def ForeThemeColorIndex(self):
        return self.tabcontrol.ForeThemeColorIndex

    @property
    def ForeTint(self):
        return self.tabcontrol.ForeTint

    @property
    def Gradient(self):
        return self.tabcontrol.Gradient

    @property
    def GridlineColor(self):
        return self.tabcontrol.GridlineColor

    @property
    def GridlineShade(self):
        return self.tabcontrol.GridlineShade

    @property
    def GridlineStyleBottom(self):
        return self.tabcontrol.GridlineStyleBottom

    @property
    def GridlineStyleLeft(self):
        return self.tabcontrol.GridlineStyleLeft

    @property
    def GridlineStyleRight(self):
        return self.tabcontrol.GridlineStyleRight

    @property
    def GridlineStyleTop(self):
        return self.tabcontrol.GridlineStyleTop

    @property
    def GridlineThemeColorIndex(self):
        return self.tabcontrol.GridlineThemeColorIndex

    @property
    def GridlineTint(self):
        return self.tabcontrol.GridlineTint

    @property
    def GridlineWidthBottom(self):
        return self.tabcontrol.GridlineWidthBottom

    @property
    def GridlineWidthLeft(self):
        return self.tabcontrol.GridlineWidthLeft

    @property
    def GridlineWidthRight(self):
        return self.tabcontrol.GridlineWidthRight

    @property
    def GridlineWidthTop(self):
        return self.tabcontrol.GridlineWidthTop

    @property
    def Height(self):
        return self.tabcontrol.Height

    @property
    def HelpContextId(self):
        return self.tabcontrol.HelpContextId

    @property
    def HorizontalAnchor(self):
        return self.tabcontrol.HorizontalAnchor

    @property
    def HoverColor(self):
        return self.tabcontrol.HoverColor

    @property
    def HoverForeColor(self):
        return self.tabcontrol.HoverForeColor

    @property
    def HoverForeShade(self):
        return self.tabcontrol.HoverForeShade

    @property
    def HoverForeThemeColorIndex(self):
        return self.tabcontrol.HoverForeThemeColorIndex

    @property
    def HoverForeTint(self):
        return self.tabcontrol.HoverForeTint

    @property
    def HoverShade(self):
        return self.tabcontrol.HoverShade

    @property
    def HoverThemeColorIndex(self):
        return self.tabcontrol.HoverThemeColorIndex

    @property
    def HoverTint(self):
        return self.tabcontrol.HoverTint

    @property
    def InSelection(self):
        return self.tabcontrol.InSelection

    @property
    def IsVisible(self):
        return self.tabcontrol.IsVisible

    @property
    def Layout(self):
        return AcLayoutType(self.tabcontrol.Layout)

    @property
    def LayoutID(self):
        return self.tabcontrol.LayoutID

    @property
    def Left(self):
        return self.tabcontrol.Left

    @property
    def LeftPadding(self):
        return self.tabcontrol.LeftPadding

    @property
    def MultiRow(self):
        return self.tabcontrol.MultiRow

    @property
    def Name(self):
        return self.tabcontrol.Name

    @property
    def OldValue(self):
        return self.tabcontrol.OldValue

    @property
    def OnChange(self):
        return self.tabcontrol.OnChange

    @property
    def OnClick(self):
        return self.tabcontrol.OnClick

    @property
    def OnDblClick(self):
        return self.tabcontrol.OnDblClick

    @property
    def OnKeyDown(self):
        return self.tabcontrol.OnKeyDown

    @property
    def OnKeyPress(self):
        return self.tabcontrol.OnKeyPress

    @property
    def OnKeyUp(self):
        return self.tabcontrol.OnKeyUp

    @property
    def OnMouseDown(self):
        return self.tabcontrol.OnMouseDown

    @property
    def OnMouseMove(self):
        return self.tabcontrol.OnMouseMove

    @property
    def OnMouseUp(self):
        return self.tabcontrol.OnMouseUp

    @property
    def Pages(self):
        return Pages(self.tabcontrol.Pages)

    @property
    def Parent(self):
        return self.tabcontrol.Parent

    @property
    def PressedColor(self):
        return self.tabcontrol.PressedColor

    @property
    def PressedForeColor(self):
        return self.tabcontrol.PressedForeColor

    @property
    def PressedForeShade(self):
        return self.tabcontrol.PressedForeShade

    @property
    def PressedForeThemeColorIndex(self):
        return self.tabcontrol.PressedForeThemeColorIndex

    @property
    def PressedForeTint(self):
        return self.tabcontrol.PressedForeTint

    @property
    def PressedShade(self):
        return self.tabcontrol.PressedShade

    @property
    def PressedThemeColorIndex(self):
        return self.tabcontrol.PressedThemeColorIndex

    @property
    def PressedTint(self):
        return self.tabcontrol.PressedTint

    @property
    def Properties(self):
        return Properties(self.tabcontrol.Properties)

    @property
    def RightPadding(self):
        return self.tabcontrol.RightPadding

    @property
    def Section(self):
        return self.tabcontrol.Section

    @property
    def Shape(self):
        return self.tabcontrol.Shape

    @Shape.setter
    def Shape(self, value):
        self.tabcontrol.Shape = value

    @property
    def ShortcutMenuBar(self):
        return self.tabcontrol.ShortcutMenuBar

    @property
    def StatusBarText(self):
        return self.tabcontrol.StatusBarText

    @property
    def Style(self):
        return self.tabcontrol.Style

    @property
    def TabFixedHeight(self):
        return self.tabcontrol.TabFixedHeight

    @property
    def TabFixedWidth(self):
        return self.tabcontrol.TabFixedWidth

    @property
    def TabIndex(self):
        return self.tabcontrol.TabIndex

    @property
    def TabStop(self):
        return self.tabcontrol.TabStop

    @property
    def Tag(self):
        return self.tabcontrol.Tag

    @property
    def ThemeFontIndex(self):
        return self.tabcontrol.ThemeFontIndex

    @property
    def Top(self):
        return self.tabcontrol.Top

    @property
    def TopPadding(self):
        return self.tabcontrol.TopPadding

    @property
    def UseTheme(self):
        return self.tabcontrol.UseTheme

    @property
    def Value(self):
        return self.tabcontrol.Value

    @property
    def VerticalAnchor(self):
        return self.tabcontrol.VerticalAnchor

    @property
    def Visible(self):
        return self.tabcontrol.Visible

    @Visible.setter
    def Visible(self, value):
        self.tabcontrol.Visible = value

    @property
    def Width(self):
        return self.tabcontrol.Width

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
        return self.tempvars.Item(*args, **arguments)

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

    @property
    def AllowAutoCorrect(self):
        return self.textbox.AllowAutoCorrect

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

    @property
    def AutoTab(self):
        return self.textbox.AutoTab

    @property
    def BackColor(self):
        return self.textbox.BackColor

    @property
    def BackShade(self):
        return self.textbox.BackShade

    @property
    def BackStyle(self):
        return self.textbox.BackStyle

    @property
    def BackThemeColorIndex(self):
        return self.textbox.BackThemeColorIndex

    @property
    def BackTint(self):
        return self.textbox.BackTint

    @property
    def BorderColor(self):
        return self.textbox.BorderColor

    @property
    def BorderShade(self):
        return self.textbox.BorderShade

    @property
    def BorderStyle(self):
        return self.textbox.BorderStyle

    @property
    def BorderThemeColorIndex(self):
        return self.textbox.BorderThemeColorIndex

    @property
    def BorderTint(self):
        return self.textbox.BorderTint

    @property
    def BorderWidth(self):
        return self.textbox.BorderWidth

    @property
    def BottomMargin(self):
        return self.textbox.BottomMargin

    @property
    def BottomPadding(self):
        return self.textbox.BottomPadding

    @property
    def CanGrow(self):
        return self.textbox.CanGrow

    @property
    def CanShrink(self):
        return self.textbox.CanShrink

    @property
    def ColumnHidden(self):
        return self.textbox.ColumnHidden

    @property
    def ColumnOrder(self):
        return self.textbox.ColumnOrder

    @property
    def ColumnWidth(self):
        return self.textbox.ColumnWidth

    @property
    def Controls(self):
        return Controls(self.textbox.Controls)

    @property
    def ControlSource(self):
        return self.textbox.ControlSource

    @property
    def ControlTipText(self):
        return self.textbox.ControlTipText

    @property
    def ControlType(self):
        return self.textbox.ControlType

    @property
    def DecimalPlaces(self):
        return self.textbox.DecimalPlaces

    @property
    def DefaultValue(self):
        return self.textbox.DefaultValue

    @property
    def DisplayAsHyperlink(self):
        return self.textbox.DisplayAsHyperlink

    @property
    def DisplayWhen(self):
        return self.textbox.DisplayWhen

    @property
    def Enabled(self):
        return self.textbox.Enabled

    @property
    def EnterKeyBehavior(self):
        return self.textbox.EnterKeyBehavior

    @property
    def EventProcPrefix(self):
        return self.textbox.EventProcPrefix

    @property
    def FilterLookup(self):
        return self.textbox.FilterLookup

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

    @property
    def ForeColor(self):
        return self.textbox.ForeColor

    @property
    def ForeShade(self):
        return self.textbox.ForeShade

    @property
    def ForeThemeColorIndex(self):
        return self.textbox.ForeThemeColorIndex

    @property
    def ForeTint(self):
        return self.textbox.ForeTint

    @property
    def Format(self):
        return self.textbox.Format

    @property
    def FormatConditions(self):
        return self.textbox.FormatConditions

    @property
    def GridlineColor(self):
        return self.textbox.GridlineColor

    @property
    def GridlineShade(self):
        return self.textbox.GridlineShade

    @property
    def GridlineStyleBottom(self):
        return self.textbox.GridlineStyleBottom

    @property
    def GridlineStyleLeft(self):
        return self.textbox.GridlineStyleLeft

    @property
    def GridlineStyleRight(self):
        return self.textbox.GridlineStyleRight

    @property
    def GridlineStyleTop(self):
        return self.textbox.GridlineStyleTop

    @property
    def GridlineThemeColorIndex(self):
        return self.textbox.GridlineThemeColorIndex

    @property
    def GridlineTint(self):
        return self.textbox.GridlineTint

    @property
    def GridlineWidthBottom(self):
        return self.textbox.GridlineWidthBottom

    @property
    def GridlineWidthLeft(self):
        return self.textbox.GridlineWidthLeft

    @property
    def GridlineWidthRight(self):
        return self.textbox.GridlineWidthRight

    @property
    def GridlineWidthTop(self):
        return self.textbox.GridlineWidthTop

    @property
    def Height(self):
        return self.textbox.Height

    @property
    def HelpContextId(self):
        return self.textbox.HelpContextId

    @property
    def HideDuplicates(self):
        return self.textbox.HideDuplicates

    @property
    def HorizontalAnchor(self):
        return self.textbox.HorizontalAnchor

    @property
    def Hyperlink(self):
        return self.textbox.Hyperlink

    @property
    def IMEHold(self):
        return self.textbox.IMEHold

    @property
    def InputMask(self):
        return self.textbox.InputMask

    @property
    def InSelection(self):
        return self.textbox.InSelection

    @property
    def IsHyperlink(self):
        return self.textbox.IsHyperlink

    @property
    def IsVisible(self):
        return self.textbox.IsVisible

    @property
    def LabelAlign(self):
        return self.textbox.LabelAlign

    @property
    def LabelX(self):
        return self.textbox.LabelX

    @property
    def LabelY(self):
        return self.textbox.LabelY

    @property
    def Layout(self):
        return AcLayoutType(self.textbox.Layout)

    @property
    def LayoutID(self):
        return self.textbox.LayoutID

    @property
    def Left(self):
        return self.textbox.Left

    @property
    def LeftMargin(self):
        return self.textbox.LeftMargin

    @property
    def LeftPadding(self):
        return self.textbox.LeftPadding

    @property
    def LineSpacing(self):
        return self.textbox.LineSpacing

    @property
    def Locked(self):
        return self.textbox.Locked

    @property
    def Name(self):
        return self.textbox.Name

    @property
    def OldBorderStyle(self):
        return self.textbox.OldBorderStyle

    @property
    def OldValue(self):
        return self.textbox.OldValue

    @property
    def OnChange(self):
        return self.textbox.OnChange

    @property
    def OnClick(self):
        return self.textbox.OnClick

    @property
    def OnDblClick(self):
        return self.textbox.OnDblClick

    @property
    def OnDirty(self):
        return self.textbox.OnDirty

    @property
    def OnEnter(self):
        return self.textbox.OnEnter

    @property
    def OnExit(self):
        return self.textbox.OnExit

    @property
    def OnGotFocus(self):
        return self.textbox.OnGotFocus

    @property
    def OnKeyDown(self):
        return self.textbox.OnKeyDown

    @property
    def OnKeyPress(self):
        return self.textbox.OnKeyPress

    @property
    def OnKeyUp(self):
        return self.textbox.OnKeyUp

    @property
    def OnLostFocus(self):
        return self.textbox.OnLostFocus

    @property
    def OnMouseDown(self):
        return self.textbox.OnMouseDown

    @property
    def OnMouseMove(self):
        return self.textbox.OnMouseMove

    @property
    def OnMouseUp(self):
        return self.textbox.OnMouseUp

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

    @property
    def Properties(self):
        return Properties(self.textbox.Properties)

    @property
    def ReadingOrder(self):
        return self.textbox.ReadingOrder

    @property
    def RightMargin(self):
        return self.textbox.RightMargin

    @property
    def RightPadding(self):
        return self.textbox.RightPadding

    @property
    def RunningSum(self):
        return self.textbox.RunningSum

    @property
    def ScrollBarAlign(self):
        return self.textbox.ScrollBarAlign

    @property
    def ScrollBars(self):
        return self.textbox.ScrollBars

    @property
    def Section(self):
        return self.textbox.Section

    @property
    def SelLength(self):
        return self.textbox.SelLength

    @property
    def SelStart(self):
        return self.textbox.SelStart

    @property
    def SelText(self):
        return self.textbox.SelText

    @property
    def ShortcutMenuBar(self):
        return self.textbox.ShortcutMenuBar

    @property
    def ShowDatePicker(self):
        return self.textbox.ShowDatePicker

    @property
    def SmartTags(self):
        return SmartTags(self.textbox.SmartTags)

    @property
    def SpecialEffect(self):
        return self.textbox.SpecialEffect

    @property
    def StatusBarText(self):
        return self.textbox.StatusBarText

    @property
    def TabIndex(self):
        return self.textbox.TabIndex

    @property
    def TabStop(self):
        return self.textbox.TabStop

    @property
    def Tag(self):
        return self.textbox.Tag

    @property
    def Text(self):
        return self.textbox.Text

    @property
    def TextAlign(self):
        return self.textbox.TextAlign

    @property
    def TextFormat(self):
        return self.textbox.TextFormat

    @property
    def ThemeFontIndex(self):
        return self.textbox.ThemeFontIndex

    @property
    def Top(self):
        return self.textbox.Top

    @property
    def TopMargin(self):
        return self.textbox.TopMargin

    @property
    def TopPadding(self):
        return self.textbox.TopPadding

    @property
    def ValidationRule(self):
        return self.textbox.ValidationRule

    @property
    def ValidationText(self):
        return self.textbox.ValidationText

    @property
    def Value(self):
        return self.textbox.Value

    @property
    def Vertical(self):
        return self.textbox.Vertical

    @property
    def VerticalAnchor(self):
        return self.textbox.VerticalAnchor

    @property
    def Visible(self):
        return self.textbox.Visible

    @Visible.setter
    def Visible(self, value):
        self.textbox.Visible = value

    @property
    def Width(self):
        return self.textbox.Width

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

    @property
    def Application(self):
        return self.togglebutton.Application

    @property
    def AutoLabel(self):
        return self.togglebutton.AutoLabel

    @property
    def BackColor(self):
        return self.togglebutton.BackColor

    @property
    def BackShade(self):
        return self.togglebutton.BackShade

    @property
    def BackThemeColorIndex(self):
        return self.togglebutton.BackThemeColorIndex

    @property
    def BackTint(self):
        return self.togglebutton.BackTint

    @property
    def Bevel(self):
        return self.togglebutton.Bevel

    @property
    def BorderColor(self):
        return self.togglebutton.BorderColor

    @property
    def BorderShade(self):
        return self.togglebutton.BorderShade

    @property
    def BorderStyle(self):
        return self.togglebutton.BorderStyle

    @property
    def BorderThemeColorIndex(self):
        return self.togglebutton.BorderThemeColorIndex

    @property
    def BorderTint(self):
        return self.togglebutton.BorderTint

    @property
    def BorderWidth(self):
        return self.togglebutton.BorderWidth

    @property
    def BottomPadding(self):
        return self.togglebutton.BottomPadding

    @property
    def Caption(self):
        return self.togglebutton.Caption

    @property
    def ColumnHidden(self):
        return self.togglebutton.ColumnHidden

    @property
    def ColumnOrder(self):
        return self.togglebutton.ColumnOrder

    @property
    def ColumnWidth(self):
        return self.togglebutton.ColumnWidth

    @property
    def Controls(self):
        return Controls(self.togglebutton.Controls)

    @property
    def ControlSource(self):
        return self.togglebutton.ControlSource

    @property
    def ControlTipText(self):
        return self.togglebutton.ControlTipText

    @property
    def ControlType(self):
        return self.togglebutton.ControlType

    @property
    def DefaultValue(self):
        return self.togglebutton.DefaultValue

    @property
    def DisplayWhen(self):
        return self.togglebutton.DisplayWhen

    @property
    def Enabled(self):
        return self.togglebutton.Enabled

    @property
    def EventProcPrefix(self):
        return self.togglebutton.EventProcPrefix

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

    @property
    def ForeColor(self):
        return self.togglebutton.ForeColor

    @property
    def ForeShade(self):
        return self.togglebutton.ForeShade

    @property
    def ForeThemeColorIndex(self):
        return self.togglebutton.ForeThemeColorIndex

    @property
    def ForeTint(self):
        return self.togglebutton.ForeTint

    @property
    def Glow(self):
        return self.togglebutton.Glow

    @property
    def Gradient(self):
        return self.togglebutton.Gradient

    @property
    def GridlineColor(self):
        return self.togglebutton.GridlineColor

    @property
    def GridlineShade(self):
        return self.togglebutton.GridlineShade

    @property
    def GridlineStyleBottom(self):
        return self.togglebutton.GridlineStyleBottom

    @property
    def GridlineStyleLeft(self):
        return self.togglebutton.GridlineStyleLeft

    @property
    def GridlineStyleRight(self):
        return self.togglebutton.GridlineStyleRight

    @property
    def GridlineStyleTop(self):
        return self.togglebutton.GridlineStyleTop

    @property
    def GridlineThemeColorIndex(self):
        return self.togglebutton.GridlineThemeColorIndex

    @property
    def GridlineTint(self):
        return self.togglebutton.GridlineTint

    @property
    def GridlineWidthBottom(self):
        return self.togglebutton.GridlineWidthBottom

    @property
    def GridlineWidthLeft(self):
        return self.togglebutton.GridlineWidthLeft

    @property
    def GridlineWidthRight(self):
        return self.togglebutton.GridlineWidthRight

    @property
    def GridlineWidthTop(self):
        return self.togglebutton.GridlineWidthTop

    @property
    def Height(self):
        return self.togglebutton.Height

    @property
    def HelpContextId(self):
        return self.togglebutton.HelpContextId

    @property
    def HideDuplicates(self):
        return self.togglebutton.HideDuplicates

    @property
    def HorizontalAnchor(self):
        return self.togglebutton.HorizontalAnchor

    @property
    def HoverColor(self):
        return self.togglebutton.HoverColor

    @property
    def HoverForeColor(self):
        return self.togglebutton.HoverForeColor

    @property
    def HoverForeShade(self):
        return self.togglebutton.HoverForeShade

    @property
    def HoverForeThemeColorIndex(self):
        return self.togglebutton.HoverForeThemeColorIndex

    @property
    def HoverForeTint(self):
        return self.togglebutton.HoverForeTint

    @property
    def HoverShade(self):
        return self.togglebutton.HoverShade

    @property
    def HoverThemeColorIndex(self):
        return self.togglebutton.HoverThemeColorIndex

    @property
    def HoverTint(self):
        return self.togglebutton.HoverTint

    @property
    def InSelection(self):
        return self.togglebutton.InSelection

    @property
    def IsVisible(self):
        return self.togglebutton.IsVisible

    @property
    def LabelAlign(self):
        return self.togglebutton.LabelAlign

    @property
    def LabelX(self):
        return self.togglebutton.LabelX

    @property
    def LabelY(self):
        return self.togglebutton.LabelY

    @property
    def Layout(self):
        return AcLayoutType(self.togglebutton.Layout)

    @property
    def LayoutID(self):
        return self.togglebutton.LayoutID

    @property
    def Left(self):
        return self.togglebutton.Left

    @property
    def LeftPadding(self):
        return self.togglebutton.LeftPadding

    @property
    def Locked(self):
        return self.togglebutton.Locked

    @property
    def Name(self):
        return self.togglebutton.Name

    @property
    def ObjectPalette(self):
        return self.togglebutton.ObjectPalette

    @property
    def OldValue(self):
        return self.togglebutton.OldValue

    @property
    def OnClick(self):
        return self.togglebutton.OnClick

    @property
    def OnDblClick(self):
        return self.togglebutton.OnDblClick

    @property
    def OnEnter(self):
        return self.togglebutton.OnEnter

    @property
    def OnExit(self):
        return self.togglebutton.OnExit

    @property
    def OnGotFocus(self):
        return self.togglebutton.OnGotFocus

    @property
    def OnKeyDown(self):
        return self.togglebutton.OnKeyDown

    @property
    def OnKeyPress(self):
        return self.togglebutton.OnKeyPress

    @property
    def OnKeyUp(self):
        return self.togglebutton.OnKeyUp

    @property
    def OnLostFocus(self):
        return self.togglebutton.OnLostFocus

    @property
    def OnMouseDown(self):
        return self.togglebutton.OnMouseDown

    @property
    def OnMouseMove(self):
        return self.togglebutton.OnMouseMove

    @property
    def OnMouseUp(self):
        return self.togglebutton.OnMouseUp

    @property
    def OptionValue(self):
        return self.togglebutton.OptionValue

    @property
    def Parent(self):
        return self.togglebutton.Parent

    @property
    def Picture(self):
        return self.togglebutton.Picture

    @property
    def PictureData(self):
        return self.togglebutton.PictureData

    @property
    def PictureType(self):
        return self.togglebutton.PictureType

    @property
    def PressedColor(self):
        return self.togglebutton.PressedColor

    @property
    def PressedForeColor(self):
        return self.togglebutton.PressedForeColor

    @property
    def PressedForeShade(self):
        return self.togglebutton.PressedForeShade

    @property
    def PressedForeThemeColorIndex(self):
        return self.togglebutton.PressedForeThemeColorIndex

    @property
    def PressedForeTint(self):
        return self.togglebutton.PressedForeTint

    @property
    def PressedShade(self):
        return self.togglebutton.PressedShade

    @property
    def PressedThemeColorIndex(self):
        return self.togglebutton.PressedThemeColorIndex

    @property
    def PressedTint(self):
        return self.togglebutton.PressedTint

    @property
    def Properties(self):
        return Properties(self.togglebutton.Properties)

    @property
    def QuickStyle(self):
        return self.togglebutton.QuickStyle

    @property
    def ReadingOrder(self):
        return self.togglebutton.ReadingOrder

    @property
    def RightPadding(self):
        return self.togglebutton.RightPadding

    @property
    def Section(self):
        return self.togglebutton.Section

    @property
    def Shadow(self):
        return self.togglebutton.Shadow

    @property
    def Shape(self):
        return self.togglebutton.Shape

    @Shape.setter
    def Shape(self, value):
        self.togglebutton.Shape = value

    @property
    def ShortcutMenuBar(self):
        return self.togglebutton.ShortcutMenuBar

    @property
    def SoftEdges(self):
        return self.togglebutton.SoftEdges

    @property
    def StatusBarText(self):
        return self.togglebutton.StatusBarText

    @property
    def TabIndex(self):
        return self.togglebutton.TabIndex

    @property
    def TabStop(self):
        return self.togglebutton.TabStop

    @property
    def Tag(self):
        return self.togglebutton.Tag

    @property
    def ThemeFontIndex(self):
        return self.togglebutton.ThemeFontIndex

    @property
    def Top(self):
        return self.togglebutton.Top

    @property
    def TopPadding(self):
        return self.togglebutton.TopPadding

    @property
    def TripleState(self):
        return self.togglebutton.TripleState

    @property
    def UseTheme(self):
        return self.togglebutton.UseTheme

    @property
    def ValidationRule(self):
        return self.togglebutton.ValidationRule

    @property
    def ValidationText(self):
        return self.togglebutton.ValidationText

    @property
    def Value(self):
        return self.togglebutton.Value

    @property
    def VerticalAnchor(self):
        return self.togglebutton.VerticalAnchor

    @property
    def Visible(self):
        return self.togglebutton.Visible

    @Visible.setter
    def Visible(self, value):
        self.togglebutton.Visible = value

    @property
    def Width(self):
        return self.togglebutton.Width

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

    @property
    def BorderShade(self):
        return self.webbrowsercontrol.BorderShade

    @property
    def BorderStyle(self):
        return self.webbrowsercontrol.BorderStyle

    @property
    def BorderThemeColorIndex(self):
        return self.webbrowsercontrol.BorderThemeColorIndex

    @property
    def BorderTint(self):
        return self.webbrowsercontrol.BorderTint

    @property
    def BorderWidth(self):
        return self.webbrowsercontrol.BorderWidth

    @property
    def BottomPadding(self):
        return self.webbrowsercontrol.BottomPadding

    @property
    def Controls(self):
        return Controls(self.webbrowsercontrol.Controls)

    @property
    def ControlSource(self):
        return self.webbrowsercontrol.ControlSource

    @property
    def ControlTipText(self):
        return self.webbrowsercontrol.ControlTipText

    @property
    def ControlType(self):
        return self.webbrowsercontrol.ControlType

    @property
    def DisplayWhen(self):
        return self.webbrowsercontrol.DisplayWhen

    @property
    def Enabled(self):
        return self.webbrowsercontrol.Enabled

    @property
    def EventProcPrefix(self):
        return self.webbrowsercontrol.EventProcPrefix

    @property
    def GridlineColor(self):
        return self.webbrowsercontrol.GridlineColor

    @property
    def GridlineShade(self):
        return self.webbrowsercontrol.GridlineShade

    @property
    def GridlineStyleBottom(self):
        return self.webbrowsercontrol.GridlineStyleBottom

    @property
    def GridlineStyleLeft(self):
        return self.webbrowsercontrol.GridlineStyleLeft

    @property
    def GridlineStyleRight(self):
        return self.webbrowsercontrol.GridlineStyleRight

    @property
    def GridlineStyleTop(self):
        return self.webbrowsercontrol.GridlineStyleTop

    @property
    def GridlineThemeColorIndex(self):
        return self.webbrowsercontrol.GridlineThemeColorIndex

    @property
    def GridlineTint(self):
        return self.webbrowsercontrol.GridlineTint

    @property
    def GridlineWidthBottom(self):
        return self.webbrowsercontrol.GridlineWidthBottom

    @property
    def GridlineWidthLeft(self):
        return self.webbrowsercontrol.GridlineWidthLeft

    @property
    def GridlineWidthRight(self):
        return self.webbrowsercontrol.GridlineWidthRight

    @property
    def GridlineWidthTop(self):
        return self.webbrowsercontrol.GridlineWidthTop

    @property
    def Height(self):
        return self.webbrowsercontrol.Height

    @property
    def HelpContextId(self):
        return self.webbrowsercontrol.HelpContextId

    @property
    def HorizontalAnchor(self):
        return self.webbrowsercontrol.HorizontalAnchor

    @property
    def Hyperlink(self):
        return self.webbrowsercontrol.Hyperlink

    @property
    def InSelection(self):
        return self.webbrowsercontrol.InSelection

    @property
    def Layout(self):
        return AcLayoutType(self.webbrowsercontrol.Layout)

    @property
    def LayoutID(self):
        return self.webbrowsercontrol.LayoutID

    @property
    def Left(self):
        return self.webbrowsercontrol.Left

    @property
    def LeftPadding(self):
        return self.webbrowsercontrol.LeftPadding

    @property
    def LocationURL(self):
        return self.webbrowsercontrol.LocationURL

    @property
    def Name(self):
        return self.webbrowsercontrol.Name

    @property
    def Object(self):
        return self.webbrowsercontrol.Object

    @property
    def OldValue(self):
        return self.webbrowsercontrol.OldValue

    @property
    def OnBeforeNavigate(self):
        return self.webbrowsercontrol.OnBeforeNavigate

    @property
    def OnDocumentComplete(self):
        return self.webbrowsercontrol.OnDocumentComplete

    @property
    def OnKeyDown(self):
        return self.webbrowsercontrol.OnKeyDown

    @property
    def OnKeyPress(self):
        return self.webbrowsercontrol.OnKeyPress

    @property
    def OnKeyUp(self):
        return self.webbrowsercontrol.OnKeyUp

    @property
    def OnMouseDown(self):
        return self.webbrowsercontrol.OnMouseDown

    @property
    def OnMouseMove(self):
        return self.webbrowsercontrol.OnMouseMove

    @property
    def OnMouseUp(self):
        return self.webbrowsercontrol.OnMouseUp

    @property
    def OnNavigateError(self):
        return self.webbrowsercontrol.OnNavigateError

    @property
    def OnProgressChange(self):
        return self.webbrowsercontrol.OnProgressChange

    @property
    def OnUpdated(self):
        return self.webbrowsercontrol.OnUpdated

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

    @property
    def ScrollBars(self):
        return self.webbrowsercontrol.ScrollBars

    @property
    def ScrollLeft(self):
        return self.webbrowsercontrol.ScrollLeft

    @property
    def ScrollTop(self):
        return self.webbrowsercontrol.ScrollTop

    @property
    def Section(self):
        return self.webbrowsercontrol.Section

    @property
    def SpecialEffect(self):
        return self.webbrowsercontrol.SpecialEffect

    @property
    def StatusBarText(self):
        return self.webbrowsercontrol.StatusBarText

    @property
    def TabIndex(self):
        return self.webbrowsercontrol.TabIndex

    @property
    def TabStop(self):
        return self.webbrowsercontrol.TabStop

    @property
    def Tag(self):
        return self.webbrowsercontrol.Tag

    @property
    def Top(self):
        return self.webbrowsercontrol.Top

    @property
    def TopPadding(self):
        return self.webbrowsercontrol.TopPadding

    @property
    def Transform(self):
        return self.webbrowsercontrol.Transform

    @property
    def Value(self):
        return self.webbrowsercontrol.Value

    @property
    def VerticalAnchor(self):
        return self.webbrowsercontrol.VerticalAnchor

    @property
    def Visible(self):
        return self.webbrowsercontrol.Visible

    @Visible.setter
    def Visible(self, value):
        self.webbrowsercontrol.Visible = value

    @property
    def Width(self):
        return self.webbrowsercontrol.Width

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
        return self.webservices.Item(*args, **arguments)

    @property
    def Parent(self):
        return self.webservices.Parent

class WSParameter:

    def __init__(self, wsparameter=None):
        self.wsparameter = wsparameter

    @property
    def Name(self):
        return self.wsparameter.Name

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
        return self.wsparameters.Item(*args, **arguments)

    @property
    def Parent(self):
        return self.wsparameters.Parent
