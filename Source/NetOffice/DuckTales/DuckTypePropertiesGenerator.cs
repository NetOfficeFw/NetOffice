using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace NetOffice.Duck
{
    internal class DuckTypePropertiesGenerator : IDisposable
    {
        internal DuckTypePropertiesGenerator(StringBuilder builder, PropertyInfo[] properties)
        {
            HasProperties = null != properties && properties.Length > 0;
            if (!HasProperties)
                return;

            Builder = builder;
            Builder.AppendLine(Environment.NewLine + "\t\t#region Properties" + Environment.NewLine + Environment.NewLine);

            foreach (PropertyInfo item in properties)
            {              
                if (IsEnumTypeProperty(item))
                {
                    BuildEnumTypeProperty(item);
                }
                if (IsSystemValueTypeProperty(item))
                {
                    BuildSystemValueTypeProperty(item);
                }
                else if (IsNetOfficeReferenceTypeProperty(item))
                {
                    BuildNetOfficeReferenceTypeProperty(item);
                }
                else if (IsObjectTypeProperty(item))
                {
                    BuildObjectTypeProperty(item);
                }
                else
                    throw new InvalidProgramException();
            }
        }

        private StringBuilder Builder { get; set; }

        private bool HasProperties { get; set; }
        
        private bool IsNetOfficeReferenceTypeProperty(PropertyInfo item)
        {
            if (item.GetIndexParameters().Length > 0)
                throw new NotSupportedException("Property must have 0 arguments");

            return item.PropertyType.IsValueType == false && item.PropertyType.FullName.StartsWith("NetOffice");
        }

        private void BuildNetOfficeReferenceTypeProperty(PropertyInfo property)
        {          
            Builder.AppendLine("\t\tpublic " + property.PropertyType.FullName + " " +  property.Name);
            Builder.AppendLine("\t\t{");

            Builder.AppendLine("\t\t\tget");
            Builder.AppendLine("\t\t\t{");

            Builder.AppendLine("\t\t\t\tobject returnItem = Invoker.PropertyGet(this, \"" + property.Name + "\", null);");
            Builder.AppendLine("\t\t\t\treturn Factory.CreateDuckObjectFromComProxy(this, returnItem, typeof(" + property.PropertyType.FullName + ")) as " + property.PropertyType.FullName + ";");
     
            Builder.AppendLine("\t\t\t}");

            if (property.CanWrite)
            {
                Builder.AppendLine("\t\t\tset");
                Builder.AppendLine("\t\t\t{");

                Builder.AppendLine("\t\t\t\tobject[] paramsArray = Invoker.ValidateParamsArray(value);");
                Builder.AppendLine("\t\t\t\tInvoker.PropertySet(this, \"" + property.Name + "\", paramsArray);");

                Builder.AppendLine("\t\t\t}");
            }

            Builder.AppendLine("\t\t}" + Environment.NewLine);
        }

        private bool IsSystemValueTypeProperty(PropertyInfo item)
        {
            if (item.GetIndexParameters().Length > 0)
                throw new NotSupportedException("Property must have 0 arguments");

            return item.PropertyType.IsValueType || item.PropertyType.FullName == "System.String";
        }

        private bool IsEnumTypeProperty(PropertyInfo item)
        {
            if (item.GetIndexParameters().Length > 0)
                throw new NotSupportedException("Property must have 0 arguments");

            return item.PropertyType.IsEnum;
        }

        private void BuildEnumTypeProperty(PropertyInfo property)
        {
            Builder.AppendLine("\t\tpublic " + property.PropertyType.FullName + " " + property.Name);
            Builder.AppendLine("\t\t{");

            Builder.AppendLine("\t\t\tget");
            Builder.AppendLine("\t\t\t{");

            Builder.AppendLine("\t\t\t\tobject returnItem = Invoker.PropertyGet(this, \"" + property.Name + "\", null);");
            Builder.AppendLine("\t\t\t\tint intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);");
            Builder.AppendLine("\t\t\t\treturn (" + property.PropertyType.FullName + ")intReturnItem;");
            
            Builder.AppendLine("\t\t\t}");

            if (property.CanWrite)
            {
                Builder.AppendLine("\t\t\tset");
                Builder.AppendLine("\t\t\t{");

                Builder.AppendLine("\t\t\t\tobject[] paramsArray = Invoker.ValidateParamsArray(value);");
                Builder.AppendLine("\t\t\t\tInvoker.PropertySet(this, \"" + property.Name + "\", paramsArray);");

                Builder.AppendLine("\t\t\t}");
            }
            Builder.AppendLine("\t\t}" + Environment.NewLine);
        }

        private bool IsObjectTypeProperty(PropertyInfo item)
        {
            if (item.GetIndexParameters().Length > 0)
                throw new NotSupportedException("Property must have 0 arguments");
            return item.PropertyType.FullName == "System.Object";
        }

        private void BuildObjectTypeProperty(PropertyInfo property)
        {
            Builder.AppendLine("\t\tpublic " + property.PropertyType.FullName + " " + property.Name);
            Builder.AppendLine("\t\t{");

            Builder.AppendLine("\t\t\tget");
            Builder.AppendLine("\t\t\t{");

            Builder.AppendLine("\t\t\t\tobject returnItem = Invoker.PropertyGet(this, \"" + property.Name + "\", null);");

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

                Builder.AppendLine("\t\t\t\tobject[] paramsArray = Invoker.ValidateParamsArray(value);");
                Builder.AppendLine("\t\t\t\tInvoker.PropertySet(this, \"" + property.Name + "\", paramsArray);");

                Builder.AppendLine("\t\t\t}");
            }
            Builder.AppendLine("\t\t}" + Environment.NewLine);
        }
        
        private void BuildSystemValueTypeProperty(PropertyInfo property)
        {
            Builder.AppendLine("\t\tpublic " + property.PropertyType.FullName + " " + property.Name);
            Builder.AppendLine("\t\t{");

            Builder.AppendLine("\t\t\tget");
            Builder.AppendLine("\t\t\t{");

            Builder.AppendLine("\t\t\t\tobject returnItem = Invoker.PropertyGet(this, \"" + property.Name + "\", null);");
            Builder.AppendLine("\t\t\t\treturn NetRuntimeSystem.Convert.To" + property.PropertyType.Name + "(returnItem);");

            Builder.AppendLine("\t\t\t}");

            if (property.CanWrite)
            {
                Builder.AppendLine("\t\t\tset");
                Builder.AppendLine("\t\t\t{");

                Builder.AppendLine("\t\t\t\tobject[] paramsArray = Invoker.ValidateParamsArray(value);");
                Builder.AppendLine("\t\t\t\tInvoker.PropertySet(this, \"" + property.Name + "\", paramsArray);");

                Builder.AppendLine("\t\t\t}");
            }
            Builder.AppendLine("\t\t}" + Environment.NewLine);
        }

        public void Dispose()
        {
            if (HasProperties)
                Builder.AppendLine(Environment.NewLine + "\t\t#endregion");
        }
    }
}
