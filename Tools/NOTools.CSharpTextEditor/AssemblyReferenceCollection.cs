using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.CSharpTextEditor
{
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class AssemblyReferenceCollection : BindingList<AssemblyReference>
    {
        internal AssemblyReference this[string name]
        {
            get
            {
                foreach (var item in this)
                {
               
                    if (item.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase))
                        return item;
                }
                return null;
            }
        }

        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public new bool RaiseListChangedEvents
        {
            get
            {
                return base.RaiseListChangedEvents;
            }
            set
            {
                base.RaiseListChangedEvents = value;
            }
        }

        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public new bool AllowNew
        {
            get
            {
                return base.AllowNew;
            }
            set
            {
                base.AllowNew = value;
            }
        }

        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public new bool AllowEdit
        {
            get
            {
                return base.AllowEdit;
            }
            set
            {
                base.AllowEdit = value;
            }
        }

        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public new bool AllowRemove
        {
            get
            {
                return base.AllowRemove;
            }
            set
            {
                base.AllowRemove = value;
            }
        }

        public override string ToString()
        {
            return "References";
        }

        /// <summary>
        /// 
        /// </summary>
        public string[] ToStringPathArray()
        {
            List<string> list = new List<string>();

            foreach (AssemblyReference item in this)
            {
                if (!String.IsNullOrWhiteSpace(item.Path))
                    list.Add(item.Path);
                else
                    list.Add(item.Name + ".dll");
            }

            if (!list.Contains("mscorlib.dll"))
                list.Add("mscorlib.dll");

            return list.ToArray();
        }
    }
}
