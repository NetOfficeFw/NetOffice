using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _WizHook 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _WizHook : COMObject
	{
		#pragma warning disable

		#region Type Information

		/// <summary>
		/// Instance Type
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
		public override Type InstanceType
		{
			get
			{
				return LateBindingApiWrapperType;
			}
		}

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(_WizHook);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _WizHook(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _WizHook(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _WizHook(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _WizHook(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _WizHook(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _WizHook(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _WizHook() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _WizHook(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 Key
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Key");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Key", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[BaseResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VBIDEApi._VBProject DbcVbProject
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VBIDEApi._VBProject>(this, "DbcVbProject");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="bstrConnectionString">string bstrConnectionString</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool get_IsMatchToDbcConnectString(string bstrConnectionString)
		{
			return Factory.ExecuteBoolPropertyGet(this, "IsMatchToDbcConnectString", bstrConnectionString);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_IsMatchToDbcConnectString
		/// </summary>
		/// <param name="bstrConnectionString">string bstrConnectionString</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_IsMatchToDbcConnectString")]
		public bool IsMatchToDbcConnectString(string bstrConnectionString)
		{
			return get_IsMatchToDbcConnectString(bstrConnectionString);
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="actid">Int32 actid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string NameFromActid(Int32 actid)
		{
			return Factory.ExecuteStringMethodGet(this, "NameFromActid", actid);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="actid">Int32 actid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 ArgsOfActid(Int32 actid)
		{
			return Factory.ExecuteInt32MethodGet(this, "ArgsOfActid", actid);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="script">string script</param>
		/// <param name="label">string label</param>
		/// <param name="openMode">Int32 openMode</param>
		/// <param name="extra">Int32 extra</param>
		/// <param name="version">Int32 version</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 OpenScript(string script, string label, Int32 openMode, Int32 extra, Int32 version)
		{
			return Factory.ExecuteInt32MethodGet(this, "OpenScript", new object[]{ script, label, openMode, extra, version });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="hScr">Int32 hScr</param>
		/// <param name="scriptColumn">Int32 scriptColumn</param>
		/// <param name="value">string value</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool GetScriptString(Int32 hScr, Int32 scriptColumn, string value)
		{
			return Factory.ExecuteBoolMethodGet(this, "GetScriptString", hScr, scriptColumn, value);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="hScr">Int32 hScr</param>
		/// <param name="scriptColumn">Int32 scriptColumn</param>
		/// <param name="value">string value</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool SaveScriptString(Int32 hScr, Int32 scriptColumn, string value)
		{
			return Factory.ExecuteBoolMethodGet(this, "SaveScriptString", hScr, scriptColumn, value);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool GlobalProcExists(string name)
		{
			return Factory.ExecuteBoolMethodGet(this, "GlobalProcExists", name);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="table">string table</param>
		/// <param name="columns">string columns</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool TableFieldHasUniqueIndex(string table, string columns)
		{
			return Factory.ExecuteBoolMethodGet(this, "TableFieldHasUniqueIndex", table, columns);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="_string">string string</param>
		/// <param name="flags">Int32 flags</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool BracketString(string _string, Int32 flags)
		{
			return Factory.ExecuteBoolMethodGet(this, "BracketString", _string, flags);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="helpFile">string helpFile</param>
		/// <param name="wCmd">Int32 wCmd</param>
		/// <param name="contextID">Int32 contextID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool WizHelp(string helpFile, Int32 wCmd, Int32 contextID)
		{
			return Factory.ExecuteBoolMethodGet(this, "WizHelp", helpFile, wCmd, contextID);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="file">string file</param>
		/// <param name="cancelled">bool cancelled</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool OpenPictureFile(string file, bool cancelled)
		{
			return Factory.ExecuteBoolMethodGet(this, "OpenPictureFile", file, cancelled);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="_in">string in</param>
		/// <param name="_out">string out</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool EnglishPictToLocal(string _in, string _out)
		{
			return Factory.ExecuteBoolMethodGet(this, "EnglishPictToLocal", _in, _out);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="_in">string in</param>
		/// <param name="_out">string out</param>
		/// <param name="parseFlags">Int32 parseFlags</param>
		/// <param name="translateFlags">Int32 translateFlags</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool TranslateExpression(string _in, string _out, Int32 parseFlags, Int32 translateFlags)
		{
			return Factory.ExecuteBoolMethodGet(this, "TranslateExpression", _in, _out, parseFlags, translateFlags);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="file">string file</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool FileExists(string file)
		{
			return Factory.ExecuteBoolMethodGet(this, "FileExists", file);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="relativePath">string relativePath</param>
		/// <param name="fullPath">string fullPath</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 FullPath(string relativePath, string fullPath)
		{
			return Factory.ExecuteInt16MethodGet(this, "FullPath", relativePath, fullPath);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="path">string path</param>
		/// <param name="drive">string drive</param>
		/// <param name="dir">string dir</param>
		/// <param name="file">string file</param>
		/// <param name="ext">string ext</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void SplitPath(string path, string drive, string dir, string file, string ext)
		{
			 Factory.ExecuteMethod(this, "SplitPath", new object[]{ path, drive, dir, file, ext });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fontName">string fontName</param>
		/// <param name="size">Int32 size</param>
		/// <param name="weight">Int32 weight</param>
		/// <param name="italic">bool italic</param>
		/// <param name="underline">bool underline</param>
		/// <param name="cch">Int32 cch</param>
		/// <param name="caption">string caption</param>
		/// <param name="maxWidthCch">Int32 maxWidthCch</param>
		/// <param name="dx">Int32 dx</param>
		/// <param name="dy">Int32 dy</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool TwipsFromFont(string fontName, Int32 size, Int32 weight, bool italic, bool underline, Int32 cch, string caption, Int32 maxWidthCch, Int32 dx, Int32 dy)
		{
			return Factory.ExecuteBoolMethodGet(this, "TwipsFromFont", new object[]{ fontName, size, weight, italic, underline, cch, caption, maxWidthCch, dx, dy });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="recordSource">string recordSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 ObjTypOfRecordSource(string recordSource)
		{
			return Factory.ExecuteInt16MethodGet(this, "ObjTypOfRecordSource", recordSource);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="identifier">string identifier</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool IsValidIdent(string identifier)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsValidIdent", identifier);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="array">String[] array</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void SortStringArray(String[] array)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)array);
            Invoker.Method(this, "SortStringArray", paramsArray);
        }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="workspace">NetOffice.DAOApi.Workspace workspace</param>
		/// <param name="database">NetOffice.DAOApi.Database database</param>
		/// <param name="table">string table</param>
		/// <param name="returnDebugInfo">bool returnDebugInfo</param>
		/// <param name="results">string results</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 AnalyzeTable(NetOffice.DAOApi.Workspace workspace, NetOffice.DAOApi.Database database, string table, bool returnDebugInfo, string results)
		{
			return Factory.ExecuteInt32MethodGet(this, "AnalyzeTable", new object[]{ workspace, database, table, returnDebugInfo, results });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="workspace">NetOffice.DAOApi.Workspace workspace</param>
		/// <param name="database">NetOffice.DAOApi.Database database</param>
		/// <param name="query">string query</param>
		/// <param name="results">string results</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 AnalyzeQuery(NetOffice.DAOApi.Workspace workspace, NetOffice.DAOApi.Database database, string query, string results)
		{
			return Factory.ExecuteInt32MethodGet(this, "AnalyzeQuery", workspace, database, query, results);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="hwndOwner">Int32 hwndOwner</param>
		/// <param name="appName">string appName</param>
		/// <param name="dlgTitle">string dlgTitle</param>
		/// <param name="openTitle">string openTitle</param>
		/// <param name="file">string file</param>
		/// <param name="initialDir">string initialDir</param>
		/// <param name="filter">string filter</param>
		/// <param name="filterIndex">Int32 filterIndex</param>
		/// <param name="view">Int32 view</param>
		/// <param name="flags">Int32 flags</param>
		/// <param name="fOpen">bool fOpen</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 GetFileName(Int32 hwndOwner, string appName, string dlgTitle, string openTitle, string file, string initialDir, string filter, Int32 filterIndex, Int32 view, Int32 flags, bool fOpen)
		{
			return Factory.ExecuteInt32MethodGet(this, "GetFileName", new object[]{ hwndOwner, appName, dlgTitle, openTitle, file, initialDir, filter, filterIndex, view, flags, fOpen });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dpName">string dpName</param>
		/// <param name="ctlName">string ctlName</param>
		/// <param name="typ">Int32 typ</param>
		/// <param name="section">string section</param>
		/// <param name="sectionType">Int32 sectionType</param>
		/// <param name="appletCode">string appletCode</param>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="dx">Int32 dx</param>
		/// <param name="dy">Int32 dy</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void CreateDataPageControl(string dpName, string ctlName, Int32 typ, string section, Int32 sectionType, string appletCode, Int32 x, Int32 y, Int32 dx, Int32 dy)
		{
			 Factory.ExecuteMethod(this, "CreateDataPageControl", new object[]{ dpName, ctlName, typ, section, sectionType, appletCode, x, y, dx, dy });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fStart">bool fStart</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void KnownWizLeaks(bool fStart)
		{
			 Factory.ExecuteMethod(this, "KnownWizLeaks", fStart);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrDbName">string bstrDbName</param>
		/// <param name="bstrConnect">string bstrConnect</param>
		/// <param name="bstrPasswd">string bstrPasswd</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool SetVbaPassword(string bstrDbName, string bstrConnect, string bstrPasswd)
		{
			return Factory.ExecuteBoolMethodGet(this, "SetVbaPassword", bstrDbName, bstrConnect, bstrPasswd);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string LocalFont()
		{
			return Factory.ExecuteStringMethodGet(this, "LocalFont");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		/// <param name="objtyp">Int16 objtyp</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void SaveObject(string bstrName, Int16 objtyp)
		{
			 Factory.ExecuteMethod(this, "SaveObject", bstrName, objtyp);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 CurrentLangID()
		{
			return Factory.ExecuteInt32MethodGet(this, "CurrentLangID");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 KeyboardLangID()
		{
			return Factory.ExecuteInt32MethodGet(this, "KeyboardLangID");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string AccessUserDataDir()
		{
			return Factory.ExecuteStringMethodGet(this, "AccessUserDataDir");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OfficeAddInDir()
		{
			return Factory.ExecuteStringMethodGet(this, "OfficeAddInDir");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dpName">string dpName</param>
		/// <param name="fileToInsert">string fileToInsert</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string EmbedFileOnDataPage(string dpName, string fileToInsert)
		{
			return Factory.ExecuteStringMethodGet(this, "EmbedFileOnDataPage", dpName, fileToInsert);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fRptToFile">bool fRptToFile</param>
		/// <param name="bstrFileOut">string bstrFileOut</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void ReportLeaksToFile(bool fRptToFile, string bstrFileOut)
		{
			 Factory.ExecuteMethod(this, "ReportLeaksToFile", fRptToFile, bstrFileOut);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrFilename">string bstrFilename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void LoadImexSpecSolution(string bstrFilename)
		{
			 Factory.ExecuteMethod(this, "LoadImexSpecSolution", bstrFilename);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fBlockKeys">bool fBlockKeys</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void SetDpBlockKeyInput(bool fBlockKeys)
		{
			 Factory.ExecuteMethod(this, "SetDpBlockKeyInput", fBlockKeys);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="objType">NetOffice.AccessApi.Enums.AcObjectType objType</param>
		/// <param name="attribs">Int32 attribs</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool FirstDbcDataObject(string name, NetOffice.AccessApi.Enums.AcObjectType objType, Int32 attribs)
		{
			return Factory.ExecuteBoolMethodGet(this, "FirstDbcDataObject", name, objType, attribs);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool CloseCurrentDatabase()
		{
			return Factory.ExecuteBoolMethodGet(this, "CloseCurrentDatabase");
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrWhich">string bstrWhich</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string AccessWizFilePath(string bstrWhich)
		{
			return Factory.ExecuteStringMethodGet(this, "AccessWizFilePath", bstrWhich);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool HideDates()
		{
			return Factory.ExecuteBoolMethodGet(this, "HideDates");
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrBase">string bstrBase</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string GetColumns(string bstrBase)
		{
			return Factory.ExecuteStringMethodGet(this, "GetColumns", bstrBase);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrExt">string bstrExt</param>
		/// <param name="bstrFilename">string bstrFilename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public Int32 GetFileOdso(string bstrExt, string bstrFilename)
		{
			return Factory.ExecuteInt32MethodGet(this, "GetFileOdso", bstrExt, bstrFilename);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrBase">string bstrBase</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string GetInfoForColumns(string bstrBase)
		{
			return Factory.ExecuteStringMethodGet(this, "GetInfoForColumns", bstrBase);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="hwndOwner">Int32 hwndOwner</param>
		/// <param name="appName">string appName</param>
		/// <param name="dlgTitle">string dlgTitle</param>
		/// <param name="openTitle">string openTitle</param>
		/// <param name="file">string file</param>
		/// <param name="initialDir">string initialDir</param>
		/// <param name="filter">string filter</param>
		/// <param name="filterIndex">Int32 filterIndex</param>
		/// <param name="view">Int32 view</param>
		/// <param name="flags">Int32 flags</param>
		/// <param name="fOpen">bool fOpen</param>
		/// <param name="fFileSystem">object fFileSystem</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public Int32 GetFileName2(Int32 hwndOwner, string appName, string dlgTitle, string openTitle, string file, string initialDir, string filter, Int32 filterIndex, Int32 view, Int32 flags, bool fOpen, object fFileSystem)
		{
			return Factory.ExecuteInt32MethodGet(this, "GetFileName2", new object[]{ hwndOwner, appName, dlgTitle, openTitle, file, initialDir, filter, filterIndex, view, flags, fOpen, fFileSystem });
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fBlockKeys">bool fBlockKeys</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool FGetMSDE(bool fBlockKeys)
		{
			return Factory.ExecuteBoolMethodGet(this, "FGetMSDE", fBlockKeys);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrText">string bstrText</param>
		/// <param name="bstrCaption">string bstrCaption</param>
		/// <param name="wStyle">Int32 wStyle</param>
		/// <param name="idHelpID">Int32 idHelpID</param>
		/// <param name="bstrHelpFileName">string bstrHelpFileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public Int32 WizMsgBox(string bstrText, string bstrCaption, Int32 wStyle, Int32 idHelpID, string bstrHelpFileName)
		{
			return Factory.ExecuteInt32MethodGet(this, "WizMsgBox", new object[]{ bstrText, bstrCaption, wStyle, idHelpID, bstrHelpFileName });
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pbstrUID">string pbstrUID</param>
		/// <param name="pbstrPwd">string pbstrPwd</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool AdpUIDPwd(string pbstrUID, string pbstrPwd)
		{
			return Factory.ExecuteBoolMethodGet(this, "AdpUIDPwd", pbstrUID, pbstrPwd);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="lWhich">Int32 lWhich</param>
		/// <param name="vValue">object vValue</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public void SetWizGlob(Int32 lWhich, object vValue)
		{
			 Factory.ExecuteMethod(this, "SetWizGlob", lWhich, vValue);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="lWhich">Int32 lWhich</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public object GetWizGlob(Int32 lWhich)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetWizGlob", lWhich);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrADPName">string bstrADPName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public void WizCopyCmdbars(string bstrADPName)
		{
			 Factory.ExecuteMethod(this, "WizCopyCmdbars", bstrADPName);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrTableName">string bstrTableName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public Int32 GetCurrentView(string bstrTableName)
		{
			return Factory.ExecuteInt32MethodGet(this, "GetCurrentView", bstrTableName);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wch">Int32 wch</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool FIsFEWch(Int32 wch)
		{
			return Factory.ExecuteBoolMethodGet(this, "FIsFEWch", wch);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public bool IsMemberSafe(Int32 dispid)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsMemberSafe", dispid);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public string GetAccWizRCPath()
		{
			return Factory.ExecuteStringMethodGet(this, "GetAccWizRCPath");
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objtyp">Int16 objtyp</param>
		/// <param name="bstrObjName">string bstrObjName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public bool FCreateNameMap(Int16 objtyp, string bstrObjName)
		{
			return Factory.ExecuteBoolMethodGet(this, "FCreateNameMap", objtyp, bstrObjName);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		public string GetAdeRegistryPath()
		{
			return Factory.ExecuteStringMethodGet(this, "GetAdeRegistryPath");
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrSpecXML">string bstrSpecXML</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		public void ExecuteTempImexSpec(string bstrSpecXML)
		{
			 Factory.ExecuteMethod(this, "ExecuteTempImexSpec", bstrSpecXML);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		public bool FCacheStatus()
		{
			return Factory.ExecuteBoolMethodGet(this, "FCacheStatus");
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrStatus">string bstrStatus</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		public void CacheStatus(string bstrStatus)
		{
			 Factory.ExecuteMethod(this, "CacheStatus", bstrStatus);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrSpecName">string bstrSpecName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		public void SetDefaultSpecName(string bstrSpecName)
		{
			 Factory.ExecuteMethod(this, "SetDefaultSpecName", bstrSpecName);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		public string GetImexTblName()
		{
			return Factory.ExecuteStringMethodGet(this, "GetImexTblName");
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrTableName">string bstrTableName</param>
		/// <param name="bstrPropertyName">string bstrPropertyName</param>
		/// <param name="fServer">bool fServer</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		public string GetLinkedListProperty(string bstrTableName, string bstrPropertyName, bool fServer)
		{
			return Factory.ExecuteStringMethodGet(this, "GetLinkedListProperty", bstrTableName, bstrPropertyName, fServer);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="pProperty">NetOffice.AccessApi._AccessProperty pProperty</param>
		/// <param name="openMode">Int32 openMode</param>
		/// <param name="extra">Int32 extra</param>
		/// <param name="version">Int32 version</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		public Int32 OpenEmScript(NetOffice.AccessApi._AccessProperty pProperty, Int32 openMode, Int32 extra, Int32 version)
		{
			return Factory.ExecuteInt32MethodGet(this, "OpenEmScript", pProperty, openMode, extra, version);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 14,15,16)]
		public string GetDisabledExtensions()
		{
			return Factory.ExecuteStringMethodGet(this, "GetDisabledExtensions");
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <param name="bstrObjectName">string bstrObjectName</param>
		/// <param name="iobjtyp">NetOffice.AccessApi.Enums.AcObjectType iobjtyp</param>
		/// <param name="fTablesAsClient">bool fTablesAsClient</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 14,15,16)]
		public Int32 GetObjPubOption(string bstrObjectName, NetOffice.AccessApi.Enums.AcObjectType iobjtyp, bool fTablesAsClient)
		{
			return Factory.ExecuteInt32MethodGet(this, "GetObjPubOption", bstrObjectName, iobjtyp, fTablesAsClient);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <param name="bstrObjectName">string bstrObjectName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 14,15,16)]
		public bool FIsPublishedXasTable(string bstrObjectName)
		{
			return Factory.ExecuteBoolMethodGet(this, "FIsPublishedXasTable", bstrObjectName);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 14,15,16)]
		public bool FIsXasDb()
		{
			return Factory.ExecuteBoolMethodGet(this, "FIsXasDb");
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <param name="bstrObjectName">string bstrObjectName</param>
		/// <param name="iobjtyp">NetOffice.AccessApi.Enums.AcObjectType iobjtyp</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 14,15,16)]
		public bool FIsValidXasObjectName(string bstrObjectName, NetOffice.AccessApi.Enums.AcObjectType iobjtyp)
		{
			return Factory.ExecuteBoolMethodGet(this, "FIsValidXasObjectName", bstrObjectName, iobjtyp);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrObjectName">string bstrObjectName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public object LoadResourceLibrary(string bstrObjectName)
		{
			return Factory.ExecuteVariantMethodGet(this, "LoadResourceLibrary", bstrObjectName);
		}

		#endregion

		#pragma warning restore
	}
}
