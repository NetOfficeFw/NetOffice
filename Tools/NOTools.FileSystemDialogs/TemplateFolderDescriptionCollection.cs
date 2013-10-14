using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class TemplateFolderDescriptionCollection : BindingList<TemplateFolderDescription> , INotifyPropertyChanged
    {
        public TemplateFolderDescriptionCollection(PropertyChangedEventHandler eventHandler = null)
        {
            base.RaiseListChangedEvents = true;
            if (null != eventHandler)
                PropertyChanged += eventHandler;
        }

        protected override void OnListChanged(ListChangedEventArgs e)
        {
            if (e.ListChangedType == ListChangedType.ItemAdded)
                foreach (var item in this)
                    item.Parent = this;
            RaisePropertyChanged();
            base.OnListChanged(e);
            
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
            return String.Format("{0} Items", Count);
        }

        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public event PropertyChangedEventHandler PropertyChanged;

        private void RaisePropertyChanged()
        {
            if (null != PropertyChanged)
                PropertyChanged(this, new PropertyChangedEventArgs(""));
        }
    }
}
