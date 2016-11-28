using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.AccessApi
{
	///<summary>
	/// DispatchInterface _WizHook 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _WizHook : COMObject
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 Key
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Key", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Key", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VBIDEApi._VBProject DbcVbProject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DbcVbProject", paramsArray);
				NetOffice.VBIDEApi._VBProject newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VBIDEApi._VBProject;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="bstrConnectionString">string bstrConnectionString</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool get_IsMatchToDbcConnectString(string bstrConnectionString)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(bstrConnectionString);
			object returnItem = Invoker.PropertyGet(this, "IsMatchToDbcConnectString", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_IsMatchToDbcConnectString
		/// </summary>
		/// <param name="bstrConnectionString">string bstrConnectionString</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public bool IsMatchToDbcConnectString(string bstrConnectionString)
		{
			return get_IsMatchToDbcConnectString(bstrConnectionString);
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="actid">Int32 Actid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public string NameFromActid(Int32 actid)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(actid);
			object returnItem = Invoker.MethodReturn(this, "NameFromActid", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="actid">Int32 Actid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public Int32 ArgsOfActid(Int32 actid)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(actid);
			object returnItem = Invoker.MethodReturn(this, "ArgsOfActid", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="script">string Script</param>
		/// <param name="label">string Label</param>
		/// <param name="openMode">Int32 OpenMode</param>
		/// <param name="extra">Int32 Extra</param>
		/// <param name="version">Int32 Version</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public Int32 OpenScript(string script, string label, Int32 openMode, Int32 extra, Int32 version)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(script, label, openMode, extra, version);
			object returnItem = Invoker.MethodReturn(this, "OpenScript", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="hScr">Int32 HScr</param>
		/// <param name="scriptColumn">Int32 ScriptColumn</param>
		/// <param name="value">string Value</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public bool GetScriptString(Int32 hScr, Int32 scriptColumn, string value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hScr, scriptColumn, value);
			object returnItem = Invoker.MethodReturn(this, "GetScriptString", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="hScr">Int32 HScr</param>
		/// <param name="scriptColumn">Int32 ScriptColumn</param>
		/// <param name="value">string Value</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public bool SaveScriptString(Int32 hScr, Int32 scriptColumn, string value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hScr, scriptColumn, value);
			object returnItem = Invoker.MethodReturn(this, "SaveScriptString", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public bool GlobalProcExists(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "GlobalProcExists", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="table">string Table</param>
		/// <param name="columns">string Columns</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public bool TableFieldHasUniqueIndex(string table, string columns)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(table, columns);
			object returnItem = Invoker.MethodReturn(this, "TableFieldHasUniqueIndex", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="_string">string String</param>
		/// <param name="flags">Int32 flags</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public bool BracketString(string _string, Int32 flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(_string, flags);
			object returnItem = Invoker.MethodReturn(this, "BracketString", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="helpFile">string HelpFile</param>
		/// <param name="wCmd">Int32 wCmd</param>
		/// <param name="contextID">Int32 ContextID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public bool WizHelp(string helpFile, Int32 wCmd, Int32 contextID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(helpFile, wCmd, contextID);
			object returnItem = Invoker.MethodReturn(this, "WizHelp", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="file">string File</param>
		/// <param name="cancelled">bool Cancelled</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public bool OpenPictureFile(string file, bool cancelled)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(file, cancelled);
			object returnItem = Invoker.MethodReturn(this, "OpenPictureFile", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="_in">string In</param>
		/// <param name="_out">string Out</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public bool EnglishPictToLocal(string _in, string _out)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(_in, _out);
			object returnItem = Invoker.MethodReturn(this, "EnglishPictToLocal", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="_in">string In</param>
		/// <param name="_out">string Out</param>
		/// <param name="parseFlags">Int32 ParseFlags</param>
		/// <param name="translateFlags">Int32 TranslateFlags</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public bool TranslateExpression(string _in, string _out, Int32 parseFlags, Int32 translateFlags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(_in, _out, parseFlags, translateFlags);
			object returnItem = Invoker.MethodReturn(this, "TranslateExpression", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="file">string File</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public bool FileExists(string file)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(file);
			object returnItem = Invoker.MethodReturn(this, "FileExists", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="relativePath">string RelativePath</param>
		/// <param name="fullPath">string FullPath</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public Int16 FullPath(string relativePath, string fullPath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(relativePath, fullPath);
			object returnItem = Invoker.MethodReturn(this, "FullPath", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="drive">string Drive</param>
		/// <param name="dir">string Dir</param>
		/// <param name="file">string File</param>
		/// <param name="ext">string Ext</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SplitPath(string path, string drive, string dir, string file, string ext)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, drive, dir, file, ext);
			Invoker.Method(this, "SplitPath", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fontName">string FontName</param>
		/// <param name="size">Int32 Size</param>
		/// <param name="weight">Int32 Weight</param>
		/// <param name="italic">bool Italic</param>
		/// <param name="underline">bool Underline</param>
		/// <param name="cch">Int32 Cch</param>
		/// <param name="caption">string Caption</param>
		/// <param name="maxWidthCch">Int32 MaxWidthCch</param>
		/// <param name="dx">Int32 dx</param>
		/// <param name="dy">Int32 dy</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public bool TwipsFromFont(string fontName, Int32 size, Int32 weight, bool italic, bool underline, Int32 cch, string caption, Int32 maxWidthCch, Int32 dx, Int32 dy)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fontName, size, weight, italic, underline, cch, caption, maxWidthCch, dx, dy);
			object returnItem = Invoker.MethodReturn(this, "TwipsFromFont", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="recordSource">string RecordSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public Int16 ObjTypOfRecordSource(string recordSource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recordSource);
			object returnItem = Invoker.MethodReturn(this, "ObjTypOfRecordSource", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="identifier">string Identifier</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public bool IsValidIdent(string identifier)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(identifier);
			object returnItem = Invoker.MethodReturn(this, "IsValidIdent", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="array">String[] Array</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SortStringArray(String[] array)
		{
			object[] paramsArray = Invoker.ValidateParamsArray((object)array);
			Invoker.Method(this, "SortStringArray", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="workspace">NetOffice.DAOApi.Workspace Workspace</param>
		/// <param name="database">NetOffice.DAOApi.Database Database</param>
		/// <param name="table">string Table</param>
		/// <param name="returnDebugInfo">bool ReturnDebugInfo</param>
		/// <param name="results">string Results</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public Int32 AnalyzeTable(NetOffice.DAOApi.Workspace workspace, NetOffice.DAOApi.Database database, string table, bool returnDebugInfo, string results)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(workspace, database, table, returnDebugInfo, results);
			object returnItem = Invoker.MethodReturn(this, "AnalyzeTable", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="workspace">NetOffice.DAOApi.Workspace Workspace</param>
		/// <param name="database">NetOffice.DAOApi.Database Database</param>
		/// <param name="query">string Query</param>
		/// <param name="results">string Results</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public Int32 AnalyzeQuery(NetOffice.DAOApi.Workspace workspace, NetOffice.DAOApi.Database database, string query, string results)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(workspace, database, query, results);
			object returnItem = Invoker.MethodReturn(this, "AnalyzeQuery", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="hwndOwner">Int32 hwndOwner</param>
		/// <param name="appName">string AppName</param>
		/// <param name="dlgTitle">string DlgTitle</param>
		/// <param name="openTitle">string OpenTitle</param>
		/// <param name="file">string File</param>
		/// <param name="initialDir">string InitialDir</param>
		/// <param name="filter">string Filter</param>
		/// <param name="filterIndex">Int32 FilterIndex</param>
		/// <param name="view">Int32 View</param>
		/// <param name="flags">Int32 flags</param>
		/// <param name="fOpen">bool fOpen</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public Int32 GetFileName(Int32 hwndOwner, string appName, string dlgTitle, string openTitle, string file, string initialDir, string filter, Int32 filterIndex, Int32 view, Int32 flags, bool fOpen)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hwndOwner, appName, dlgTitle, openTitle, file, initialDir, filter, filterIndex, view, flags, fOpen);
			object returnItem = Invoker.MethodReturn(this, "GetFileName", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dpName">string DpName</param>
		/// <param name="ctlName">string CtlName</param>
		/// <param name="typ">Int32 Typ</param>
		/// <param name="section">string Section</param>
		/// <param name="sectionType">Int32 SectionType</param>
		/// <param name="appletCode">string AppletCode</param>
		/// <param name="x">Int32 X</param>
		/// <param name="y">Int32 Y</param>
		/// <param name="dx">Int32 dx</param>
		/// <param name="dy">Int32 dy</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void CreateDataPageControl(string dpName, string ctlName, Int32 typ, string section, Int32 sectionType, string appletCode, Int32 x, Int32 y, Int32 dx, Int32 dy)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dpName, ctlName, typ, section, sectionType, appletCode, x, y, dx, dy);
			Invoker.Method(this, "CreateDataPageControl", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fStart">bool fStart</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void KnownWizLeaks(bool fStart)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fStart);
			Invoker.Method(this, "KnownWizLeaks", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrDbName">string bstrDbName</param>
		/// <param name="bstrConnect">string bstrConnect</param>
		/// <param name="bstrPasswd">string bstrPasswd</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public bool SetVbaPassword(string bstrDbName, string bstrConnect, string bstrPasswd)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrDbName, bstrConnect, bstrPasswd);
			object returnItem = Invoker.MethodReturn(this, "SetVbaPassword", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public string LocalFont()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "LocalFont", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		/// <param name="objtyp">Int16 objtyp</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SaveObject(string bstrName, Int16 objtyp)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrName, objtyp);
			Invoker.Method(this, "SaveObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public Int32 CurrentLangID()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CurrentLangID", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public Int32 KeyboardLangID()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "KeyboardLangID", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public string AccessUserDataDir()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AccessUserDataDir", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public string OfficeAddInDir()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "OfficeAddInDir", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dpName">string DpName</param>
		/// <param name="fileToInsert">string FileToInsert</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public string EmbedFileOnDataPage(string dpName, string fileToInsert)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dpName, fileToInsert);
			object returnItem = Invoker.MethodReturn(this, "EmbedFileOnDataPage", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fRptToFile">bool fRptToFile</param>
		/// <param name="bstrFileOut">string bstrFileOut</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void ReportLeaksToFile(bool fRptToFile, string bstrFileOut)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fRptToFile, bstrFileOut);
			Invoker.Method(this, "ReportLeaksToFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrFilename">string bstrFilename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void LoadImexSpecSolution(string bstrFilename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrFilename);
			Invoker.Method(this, "LoadImexSpecSolution", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fBlockKeys">bool fBlockKeys</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SetDpBlockKeyInput(bool fBlockKeys)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fBlockKeys);
			Invoker.Method(this, "SetDpBlockKeyInput", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="objType">NetOffice.AccessApi.Enums.AcObjectType ObjType</param>
		/// <param name="attribs">Int32 Attribs</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public bool FirstDbcDataObject(string name, NetOffice.AccessApi.Enums.AcObjectType objType, Int32 attribs)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, objType, attribs);
			object returnItem = Invoker.MethodReturn(this, "FirstDbcDataObject", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public bool CloseCurrentDatabase()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CloseCurrentDatabase", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrWhich">string bstrWhich</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public string AccessWizFilePath(string bstrWhich)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrWhich);
			object returnItem = Invoker.MethodReturn(this, "AccessWizFilePath", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public bool HideDates()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "HideDates", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrBase">string bstrBase</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public string GetColumns(string bstrBase)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrBase);
			object returnItem = Invoker.MethodReturn(this, "GetColumns", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrExt">string bstrExt</param>
		/// <param name="bstrFilename">string bstrFilename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public Int32 GetFileOdso(string bstrExt, string bstrFilename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrExt, bstrFilename);
			object returnItem = Invoker.MethodReturn(this, "GetFileOdso", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrBase">string bstrBase</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public string GetInfoForColumns(string bstrBase)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrBase);
			object returnItem = Invoker.MethodReturn(this, "GetInfoForColumns", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="hwndOwner">Int32 hwndOwner</param>
		/// <param name="appName">string AppName</param>
		/// <param name="dlgTitle">string DlgTitle</param>
		/// <param name="openTitle">string OpenTitle</param>
		/// <param name="file">string File</param>
		/// <param name="initialDir">string InitialDir</param>
		/// <param name="filter">string Filter</param>
		/// <param name="filterIndex">Int32 FilterIndex</param>
		/// <param name="view">Int32 View</param>
		/// <param name="flags">Int32 flags</param>
		/// <param name="fOpen">bool fOpen</param>
		/// <param name="fFileSystem">object fFileSystem</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public Int32 GetFileName2(Int32 hwndOwner, string appName, string dlgTitle, string openTitle, string file, string initialDir, string filter, Int32 filterIndex, Int32 view, Int32 flags, bool fOpen, object fFileSystem)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hwndOwner, appName, dlgTitle, openTitle, file, initialDir, filter, filterIndex, view, flags, fOpen, fFileSystem);
			object returnItem = Invoker.MethodReturn(this, "GetFileName2", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fBlockKeys">bool fBlockKeys</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public bool FGetMSDE(bool fBlockKeys)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fBlockKeys);
			object returnItem = Invoker.MethodReturn(this, "FGetMSDE", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrText">string bstrText</param>
		/// <param name="bstrCaption">string bstrCaption</param>
		/// <param name="wStyle">Int32 wStyle</param>
		/// <param name="idHelpID">Int32 idHelpID</param>
		/// <param name="bstrHelpFileName">string bstrHelpFileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public Int32 WizMsgBox(string bstrText, string bstrCaption, Int32 wStyle, Int32 idHelpID, string bstrHelpFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrText, bstrCaption, wStyle, idHelpID, bstrHelpFileName);
			object returnItem = Invoker.MethodReturn(this, "WizMsgBox", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pbstrUID">string pbstrUID</param>
		/// <param name="pbstrPwd">string pbstrPwd</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public bool AdpUIDPwd(string pbstrUID, string pbstrPwd)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pbstrUID, pbstrPwd);
			object returnItem = Invoker.MethodReturn(this, "AdpUIDPwd", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="lWhich">Int32 lWhich</param>
		/// <param name="vValue">object vValue</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void SetWizGlob(Int32 lWhich, object vValue)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lWhich, vValue);
			Invoker.Method(this, "SetWizGlob", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="lWhich">Int32 lWhich</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public object GetWizGlob(Int32 lWhich)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lWhich);
			object returnItem = Invoker.MethodReturn(this, "GetWizGlob", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrADPName">string bstrADPName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void WizCopyCmdbars(string bstrADPName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrADPName);
			Invoker.Method(this, "WizCopyCmdbars", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrTableName">string bstrTableName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public Int32 GetCurrentView(string bstrTableName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrTableName);
			object returnItem = Invoker.MethodReturn(this, "GetCurrentView", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wch">Int32 wch</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public bool FIsFEWch(Int32 wch)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wch);
			object returnItem = Invoker.MethodReturn(this, "FIsFEWch", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 11,12,14,15,16)]
		public bool IsMemberSafe(Int32 dispid)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dispid);
			object returnItem = Invoker.MethodReturn(this, "IsMemberSafe", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 11,12,14,15,16)]
		public string GetAccWizRCPath()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetAccWizRCPath", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objtyp">Int16 objtyp</param>
		/// <param name="bstrObjName">string bstrObjName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 11,12,14,15,16)]
		public bool FCreateNameMap(Int16 objtyp, string bstrObjName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objtyp, bstrObjName);
			object returnItem = Invoker.MethodReturn(this, "FCreateNameMap", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public string GetAdeRegistryPath()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetAdeRegistryPath", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrSpecXML">string bstrSpecXML</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void ExecuteTempImexSpec(string bstrSpecXML)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrSpecXML);
			Invoker.Method(this, "ExecuteTempImexSpec", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public bool FCacheStatus()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "FCacheStatus", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrStatus">string bstrStatus</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void CacheStatus(string bstrStatus)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrStatus);
			Invoker.Method(this, "CacheStatus", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrSpecName">string bstrSpecName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void SetDefaultSpecName(string bstrSpecName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrSpecName);
			Invoker.Method(this, "SetDefaultSpecName", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public string GetImexTblName()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetImexTblName", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrTableName">string bstrTableName</param>
		/// <param name="bstrPropertyName">string bstrPropertyName</param>
		/// <param name="fServer">bool fServer</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public string GetLinkedListProperty(string bstrTableName, string bstrPropertyName, bool fServer)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrTableName, bstrPropertyName, fServer);
			object returnItem = Invoker.MethodReturn(this, "GetLinkedListProperty", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pProperty">NetOffice.AccessApi._AccessProperty pProperty</param>
		/// <param name="openMode">Int32 OpenMode</param>
		/// <param name="extra">Int32 Extra</param>
		/// <param name="version">Int32 Version</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public Int32 OpenEmScript(NetOffice.AccessApi._AccessProperty pProperty, Int32 openMode, Int32 extra, Int32 version)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pProperty, openMode, extra, version);
			object returnItem = Invoker.MethodReturn(this, "OpenEmScript", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 14,15,16)]
		public string GetDisabledExtensions()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetDisabledExtensions", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrObjectName">string bstrObjectName</param>
		/// <param name="iobjtyp">NetOffice.AccessApi.Enums.AcObjectType iobjtyp</param>
		/// <param name="fTablesAsClient">bool fTablesAsClient</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 14,15,16)]
		public Int32 GetObjPubOption(string bstrObjectName, NetOffice.AccessApi.Enums.AcObjectType iobjtyp, bool fTablesAsClient)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrObjectName, iobjtyp, fTablesAsClient);
			object returnItem = Invoker.MethodReturn(this, "GetObjPubOption", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrObjectName">string bstrObjectName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 14,15,16)]
		public bool FIsPublishedXasTable(string bstrObjectName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrObjectName);
			object returnItem = Invoker.MethodReturn(this, "FIsPublishedXasTable", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 14,15,16)]
		public bool FIsXasDb()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "FIsXasDb", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrObjectName">string bstrObjectName</param>
		/// <param name="iobjtyp">NetOffice.AccessApi.Enums.AcObjectType iobjtyp</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 14,15,16)]
		public bool FIsValidXasObjectName(string bstrObjectName, NetOffice.AccessApi.Enums.AcObjectType iobjtyp)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrObjectName, iobjtyp);
			object returnItem = Invoker.MethodReturn(this, "FIsValidXasObjectName", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="bstrObjectName">string bstrObjectName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public object LoadResourceLibrary(string bstrObjectName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrObjectName);
			object returnItem = Invoker.MethodReturn(this, "LoadResourceLibrary", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		#endregion
		#pragma warning restore
	}
}