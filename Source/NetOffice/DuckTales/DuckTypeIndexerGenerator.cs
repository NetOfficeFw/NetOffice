using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using NetOffice.Attributes;

namespace NetOffice.Duck
{
    internal class DuckTypeIndexerGenerator : IDisposable
    {       
        internal DuckTypeIndexerGenerator(StringBuilder builder, PropertyInfo[] properties, HasIndexPropertyAttribute info)
        {
            HasProperties = null != properties && properties.Length > 0;
            if (!HasProperties)
                return;

            Builder = builder;
            Info = info;
            Builder.AppendLine(Environment.NewLine + "\t\t#region Indexer" +  Environment.NewLine);

            foreach (PropertyInfo item in properties)
            {
                bool asMethod = IsInvokeAsMethod(item);
                string nameInternal = InternalName(item);

                if (IsEnumTypeProperty(item))
                {
                    BuildEnumTypeProperty(item, nameInternal, asMethod);
                }
                if (IsSystemValueTypeProperty(item))
                {
                    BuildSystemValueTypeProperty(item, nameInternal, asMethod);
                }
                else if (IsNetOfficeReferenceTypeProperty(item))
                {
                    BuildNetOfficeReferenceTypeProperty(item, nameInternal, asMethod);
                }
                else if (IsObjectTypeProperty(item))
                {
                    BuildObjectTypeProperty(item, nameInternal, asMethod);
                }
                else
                    throw new InvalidProgramException();
            }
        }

        private StringBuilder Builder { get; set; }

        private bool HasProperties { get; set; }

        private HasIndexPropertyAttribute Info{ get; set; }

        private string InternalName(PropertyInfo property)
        {
            if (null != Info)
                return Info.InvokeName;
            else
            { 
                object[] attributes = property.GetCustomAttributes(typeof(InternalNameAttribute), false);
                if (attributes.Length > 0)
                    return (attributes[0] as InternalNameAttribute).InternalName;
                else
                    return property.Name;
            }
        }

        private bool IsInvokeAsMethod(PropertyInfo property)
        {
            if (null != Info)
            {
                return Info.Invoke == IndexInvoke.Method;
            }
            else
            {
                object[] attributes = property.GetCustomAttributes(typeof(InvokeAsAttribute), false);
                if (attributes.Length > 0)
                    return (attributes[0] as InvokeAsAttribute).Invoke == Invoke.Method;
                else
                    return false;
            }
        }

        private bool IsNetOfficeReferenceTypeProperty(PropertyInfo property)
        {
            if (property.GetIndexParameters().Length ==  0)
                throw new NotSupportedException("Index property must have arguments");

            return property.PropertyType.IsValueType == false && property.PropertyType.FullName.StartsWith("NetOffice");
        }

        private void BuildNetOfficeReferenceTypeProperty(PropertyInfo property, string internalName, bool asMethod)
        {
            Builder.Append("\t\tpublic " + property.PropertyType.FullName + " this[");
            ParameterInfo[] arguments = property.GetIndexParameters();
            for (int i = 0; i < arguments.Length; i++)
            {
                ParameterInfo argument = arguments[i];
                Builder.Append(argument.ParameterType.FullName + " " + argument.Name);
                if (i < arguments.Length - 1)
                    Builder.Append(", ");
            }
            Builder.AppendLine("]");

            Builder.AppendLine("\t\t{");

            Builder.AppendLine("\t\t\tget");
            Builder.AppendLine("\t\t\t{");

            if (arguments.Length > 0)
            {
                Builder.Append("\t\t\t\tobject[] paramsArray = Invoker.ValidateParamsArray(");
                for (int i = 0; i < arguments.Length; i++)
                {
                    ParameterInfo argument = arguments[i];
                    Builder.Append(argument.Name);
                    if (i < arguments.Length - 1)
                        Builder.Append(", ");
                }
                Builder.AppendLine(");");
            }
            else
                Builder.AppendLine("\t\t\tobject[] paramsArray = null;");

            Builder.AppendLine("\t\t\t\tobject returnItem = Invoker." + (asMethod ? "MethodReturn" : "PropertyGet") + "(this, \"" + internalName + "\", paramsArray);");
            Builder.AppendLine("\t\t\t\treturn Factory.CreateDuckObjectFromComProxy(this, returnItem, typeof(" + property.PropertyType.FullName + ")) as " + property.PropertyType.FullName + ";");

            Builder.AppendLine("\t\t\t}");

            if (property.CanWrite)
            {
                Builder.AppendLine("\t\t\tset");
                Builder.AppendLine("\t\t\t{");

                if (arguments.Length > 0)
                {
                    Builder.Append("\t\t\t\tobject[] paramsArray = Invoker.ValidateParamsArray(");
                    for (int i = 0; i < arguments.Length; i++)
                    {
                        ParameterInfo argument = arguments[i];
                        Builder.Append(argument.Name);
                        if (i < arguments.Length - 1)
                            Builder.Append(", ");
                    }
                    Builder.AppendLine(");");
                }
                else
                    Builder.AppendLine("\t\t\tobject[] paramsArray = null;");

                Builder.AppendLine("\t\t\t\tInvoker." + (asMethod ? "Method" : "PropertySet") + "(this, \"" + internalName + "\", paramsArray, value);");

                Builder.AppendLine("\t\t\t}");
            }

            Builder.AppendLine("\t\t}" + Environment.NewLine);
        }

        private bool IsSystemValueTypeProperty(PropertyInfo item)
        {
            if (item.GetIndexParameters().Length == 0)
                throw new NotSupportedException("Index property must have arguments");

            return item.PropertyType.IsValueType || item.PropertyType.FullName == "System.String";
        }

        private bool IsEnumTypeProperty(PropertyInfo item)
        {
            if (item.GetIndexParameters().Length == 0)
                throw new NotSupportedException("Index property must have arguments");

            return item.PropertyType.IsEnum;
        }

        private void BuildEnumTypeProperty(PropertyInfo property, string internalName, bool asMethod)
        {
            Builder.Append("\t\tpublic " + property.PropertyType.FullName + " this[");
            ParameterInfo[] arguments = property.GetIndexParameters();
            for (int i = 0; i < arguments.Length; i++)
            {
                ParameterInfo argument = arguments[i];
                Builder.Append(argument.ParameterType.FullName + " " + argument.Name);
                if (i < arguments.Length - 1)
                    Builder.Append(", ");
            }
            Builder.AppendLine("]");
            Builder.AppendLine("\t\t{");

            Builder.AppendLine("\t\t\tget");
            Builder.AppendLine("\t\t\t{");

            if (arguments.Length > 0)
            {
                Builder.Append("\t\t\t\tobject[] paramsArray = Invoker.ValidateParamsArray(");
                for (int i = 0; i < arguments.Length; i++)
                {
                    ParameterInfo argument = arguments[i];
                    Builder.Append(argument.Name);
                    if (i < arguments.Length - 1)
                        Builder.Append(", ");
                }
                Builder.AppendLine(");");
            }
            else
                Builder.AppendLine("\t\t\tobject[] paramsArray = null;");

            Builder.AppendLine("\t\t\t\tobject returnItem = Invoker." + (asMethod ? "MethodReturn" : "PropertyGet") + "(this, \"" + internalName + "\", paramsArray);");
            Builder.AppendLine("\t\t\t\tint intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);");
            Builder.AppendLine("\t\t\t\treturn (" + property.PropertyType.FullName + ")intReturnItem;");

            Builder.AppendLine("\t\t\t}");

            if (property.CanWrite)
            {
                Builder.AppendLine("\t\t\tset");
                Builder.AppendLine("\t\t\t{");

                if (arguments.Length > 0)
                {
                    Builder.Append("\t\t\t\tobject[] paramsArray = Invoker.ValidateParamsArray(");
                    for (int i = 0; i < arguments.Length; i++)
                    {
                        ParameterInfo argument = arguments[i];
                        Builder.Append(argument.Name);
                        if (i < arguments.Length - 1)
                            Builder.Append(", ");
                    }
                    Builder.AppendLine(");");
                }
                else
                    Builder.AppendLine("\t\t\tobject[] paramsArray = null;");

                Builder.AppendLine("\t\t\t\tInvoker." + (asMethod ? "Method" : "PropertySet") +"(this, \"" + internalName + "\", paramsArray, value);");

                Builder.AppendLine("\t\t\t}");
            }
            Builder.AppendLine("\t\t}" + Environment.NewLine);
        }

        private bool IsObjectTypeProperty(PropertyInfo item)
        {
            if (item.GetIndexParameters().Length == 0)
                throw new NotSupportedException("Index property must have arguments");

            return item.PropertyType.FullName == "System.Object";
        }

        private void BuildObjectTypeProperty(PropertyInfo property, string internalName, bool asMethod)
        {
            Builder.Append("\t\tpublic " + property.PropertyType.FullName + " this[");
            ParameterInfo[] arguments = property.GetIndexParameters();
            for (int i = 0; i < arguments.Length; i++)
            {
                ParameterInfo argument = arguments[i];
                Builder.Append(argument.ParameterType.FullName + " " + argument.Name);
                if (i < arguments.Length - 1)
                    Builder.Append(", ");
            }
            Builder.AppendLine("]");
            Builder.AppendLine("\t\t{");

            Builder.AppendLine("\t\t\tget");
            Builder.AppendLine("\t\t\t{");

            if (arguments.Length > 0)
            {
                Builder.Append("\t\t\t\tobject[] paramsArray = Invoker.ValidateParamsArray(");
                for (int i = 0; i < arguments.Length; i++)
                {
                    ParameterInfo argument = arguments[i];
                    Builder.Append(argument.Name);
                    if (i < arguments.Length - 1)
                        Builder.Append(", ");
                }
                Builder.AppendLine(");");
            }
            else
                Builder.AppendLine("\t\t\tobject[] paramsArray = null;");
            Builder.AppendLine("\t\t\t\tobject returnItem = Invoker." + (asMethod ? "MethodReturn" : "PropertyGet") + "(this, \"" + internalName + "\", paramsArray);");

            Builder.AppendLine("\t\t\t\tif((null != returnItem) && (returnItem is MarshalByRefObject))");
            Builder.AppendLine("\t\t\t\t{");
            Builder.AppendLine("\t\t\t\t\tICOMObject newObject = Factory.CreateDuckObjectFromComProxy(this, returnItem);");
            Builder.AppendLine("\t\t\t\t\treturn newObject;");
            Builder.AppendLine("\t\t\t\t}");
            Builder.AppendLine("\t\t\t\telse");
            Builder.AppendLine("\t\t\t\t{");
            Builder.AppendLine("\t\t\t\t\treturn  returnItem;");
            Builder.AppendLine("\t\t\t\t}");

            Builder.AppendLine("\t\t\t}");

            if (property.CanWrite)
            {
                Builder.AppendLine("\t\t\tset");
                Builder.AppendLine("\t\t\t{");

                if (arguments.Length > 0)
                {
                    Builder.Append("\t\t\t\tobject[] paramsArray = Invoker.ValidateParamsArray(");
                    for (int i = 0; i < arguments.Length; i++)
                    {
                        ParameterInfo argument = arguments[i];
                        Builder.Append(argument.Name);
                        if (i < arguments.Length - 1)
                            Builder.Append(", ");
                    }
                    Builder.AppendLine(");");
                }
                else
                    Builder.AppendLine("\t\t\tobject[] paramsArray = null;");
                Builder.AppendLine("\t\t\t\tInvoker." + (asMethod ? "Method" : "PropertySet") +"(this, \"" + internalName + "\", paramsArray, value);");

                Builder.AppendLine("\t\t\t}");
            }
            Builder.AppendLine("\t\t}" + Environment.NewLine);
        }

        private void BuildSystemValueTypeProperty(PropertyInfo property, string internalName, bool asMethod)
        {
            Builder.Append("\t\tpublic " + property.PropertyType.FullName + " this[");
            ParameterInfo[] arguments = property.GetIndexParameters();
            for (int i = 0; i < arguments.Length; i++)
            {
                ParameterInfo argument = arguments[i];
                Builder.Append(argument.ParameterType.FullName + " " + argument.Name);
                if (i < arguments.Length - 1)
                    Builder.Append(", ");
            }
            Builder.AppendLine("]");
            Builder.AppendLine("\t\t{");

            Builder.AppendLine("\t\t\tget");
            Builder.AppendLine("\t\t\t{");

            if (arguments.Length > 0)
            {
                Builder.Append("\t\t\t\tobject[] paramsArray = Invoker.ValidateParamsArray(");
                for (int i = 0; i < arguments.Length; i++)
                {
                    ParameterInfo argument = arguments[i];
                    Builder.Append(argument.Name);
                    if (i < arguments.Length - 1)
                        Builder.Append(", ");
                }
                Builder.AppendLine(");");
            }
            else
                Builder.AppendLine("\t\t\tobject[] paramsArray = null;");
            Builder.AppendLine("\t\t\t\tobject returnItem = Invoker." + (asMethod ? "MethodReturn" : "PropertyGet") + "(this, \"" + internalName + "\", paramsArray);");
            Builder.AppendLine("\t\t\t\treturn NetRuntimeSystem.Convert.To" + property.PropertyType.Name + "(returnItem);");

            Builder.AppendLine("\t\t\t}");

            if (property.CanWrite)
            {
                Builder.AppendLine("\t\t\tset");
                Builder.AppendLine("\t\t\t{");

                if (arguments.Length > 0)
                {
                    Builder.Append("\t\t\t\tobject[] paramsArray = Invoker.ValidateParamsArray(");
                    for (int i = 0; i < arguments.Length; i++)
                    {
                        ParameterInfo argument = arguments[i];
                        Builder.Append(argument.Name);
                        if (i < arguments.Length - 1)
                            Builder.Append(", ");
                    }
                    Builder.AppendLine(");");
                }
                else
                    Builder.AppendLine("\t\t\tobject[] paramsArray = null;");
                Builder.AppendLine("\t\t\t\tInvoker." + (asMethod ? "Method" : "PropertySet") + "(this, \"" + internalName + "\", paramsArray, value);");

                Builder.AppendLine("\t\t\t}");
            }
            Builder.AppendLine("\t\t}" + Environment.NewLine);
        }

        public void Dispose()
        {
            if (HasProperties)
                Builder.AppendLine("\t\t#endregion");
        }
    }
}
