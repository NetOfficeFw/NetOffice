using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.VBIDEApi
{
	///<summary>
	/// DispatchInterface _CodeModule 
	/// SupportByVersion VBIDE, 12,14,5.3
	///</summary>
	[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _CodeModule : COMObject
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
                    _type = typeof(_CodeModule);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _CodeModule(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CodeModule(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CodeModule(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CodeModule(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CodeModule(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CodeModule() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CodeModule(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.VBComponent Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				NetOffice.VBIDEApi.VBComponent newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.VBIDEApi.VBComponent.LateBindingApiWrapperType) as NetOffice.VBIDEApi.VBComponent;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.VBE VBE
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "VBE", paramsArray);
				NetOffice.VBIDEApi.VBE newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.VBIDEApi.VBE.LateBindingApiWrapperType) as NetOffice.VBIDEApi.VBE;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public string Name
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Name", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Name", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		/// <param name="startLine">Int32 StartLine</param>
		/// <param name="count">Int32 Count</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Lines(Int32 startLine, Int32 count)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(startLine, count);
			object returnItem = Invoker.PropertyGet(this, "Lines", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Alias for get_Lines
		/// </summary>
		/// <param name="startLine">Int32 StartLine</param>
		/// <param name="count">Int32 Count</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public string Lines(Int32 startLine, Int32 count)
		{
			return get_Lines(startLine, count);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public Int32 CountOfLines
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CountOfLines", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		/// <param name="procName">string ProcName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind ProcKind</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_ProcStartLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(procName, procKind);
			object returnItem = Invoker.PropertyGet(this, "ProcStartLine", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Alias for get_ProcStartLine
		/// </summary>
		/// <param name="procName">string ProcName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind ProcKind</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public Int32 ProcStartLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return get_ProcStartLine(procName, procKind);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		/// <param name="procName">string ProcName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind ProcKind</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_ProcCountLines(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(procName, procKind);
			object returnItem = Invoker.PropertyGet(this, "ProcCountLines", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Alias for get_ProcCountLines
		/// </summary>
		/// <param name="procName">string ProcName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind ProcKind</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public Int32 ProcCountLines(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return get_ProcCountLines(procName, procKind);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		/// <param name="procName">string ProcName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind ProcKind</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_ProcBodyLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(procName, procKind);
			object returnItem = Invoker.PropertyGet(this, "ProcBodyLine", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Alias for get_ProcBodyLine
		/// </summary>
		/// <param name="procName">string ProcName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind ProcKind</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public Int32 ProcBodyLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return get_ProcBodyLine(procName, procKind);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		/// <param name="line">Int32 Line</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind ProcKind</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_ProcOfLine(Int32 line, out NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{		
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			procKind = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(line, procKind);
			object returnItem = Invoker.PropertyGet(this, "ProcOfLine", paramsArray);
			procKind = (NetOffice.VBIDEApi.Enums.vbext_ProcKind)paramsArray[1];
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Alias for get_ProcOfLine
		/// </summary>
		/// <param name="line">Int32 Line</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind ProcKind</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public string ProcOfLine(Int32 line, out NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return get_ProcOfLine(line, out procKind);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public Int32 CountOfDeclarationLines
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CountOfDeclarationLines", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.CodePane CodePane
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CodePane", paramsArray);
				NetOffice.VBIDEApi.CodePane newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.VBIDEApi.CodePane.LateBindingApiWrapperType) as NetOffice.VBIDEApi.CodePane;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="_string">string String</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public void AddFromString(string _string)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(_string);
			Invoker.Method(this, "AddFromString", paramsArray);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public void AddFromFile(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "AddFromFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="line">Int32 Line</param>
		/// <param name="_string">string String</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public void InsertLines(Int32 line, string _string)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(line, _string);
			Invoker.Method(this, "InsertLines", paramsArray);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="startLine">Int32 StartLine</param>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public void DeleteLines(Int32 startLine, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(startLine, count);
			Invoker.Method(this, "DeleteLines", paramsArray);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="startLine">Int32 StartLine</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public void DeleteLines(Int32 startLine)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(startLine);
			Invoker.Method(this, "DeleteLines", paramsArray);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="line">Int32 Line</param>
		/// <param name="_string">string String</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public void ReplaceLine(Int32 line, string _string)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(line, _string);
			Invoker.Method(this, "ReplaceLine", paramsArray);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="eventName">string EventName</param>
		/// <param name="objectName">string ObjectName</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public Int32 CreateEventProc(string eventName, string objectName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(eventName, objectName);
			object returnItem = Invoker.MethodReturn(this, "CreateEventProc", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="target">string Target</param>
		/// <param name="startLine">Int32 StartLine</param>
		/// <param name="startColumn">Int32 StartColumn</param>
		/// <param name="endLine">Int32 EndLine</param>
		/// <param name="endColumn">Int32 EndColumn</param>
		/// <param name="wholeWord">optional bool WholeWord = false</param>
		/// <param name="matchCase">optional bool MatchCase = false</param>
		/// <param name="patternSearch">optional bool PatternSearch = false</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn, object wholeWord, object matchCase, object patternSearch)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(target, startLine, startColumn, endLine, endColumn, wholeWord, matchCase, patternSearch);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="target">string Target</param>
		/// <param name="startLine">Int32 StartLine</param>
		/// <param name="startColumn">Int32 StartColumn</param>
		/// <param name="endLine">Int32 EndLine</param>
		/// <param name="endColumn">Int32 EndColumn</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(target, startLine, startColumn, endLine, endColumn);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="target">string Target</param>
		/// <param name="startLine">Int32 StartLine</param>
		/// <param name="startColumn">Int32 StartColumn</param>
		/// <param name="endLine">Int32 EndLine</param>
		/// <param name="endColumn">Int32 EndColumn</param>
		/// <param name="wholeWord">optional bool WholeWord = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn, object wholeWord)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(target, startLine, startColumn, endLine, endColumn, wholeWord);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="target">string Target</param>
		/// <param name="startLine">Int32 StartLine</param>
		/// <param name="startColumn">Int32 StartColumn</param>
		/// <param name="endLine">Int32 EndLine</param>
		/// <param name="endColumn">Int32 EndColumn</param>
		/// <param name="wholeWord">optional bool WholeWord = false</param>
		/// <param name="matchCase">optional bool MatchCase = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn, object wholeWord, object matchCase)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(target, startLine, startColumn, endLine, endColumn, wholeWord, matchCase);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}