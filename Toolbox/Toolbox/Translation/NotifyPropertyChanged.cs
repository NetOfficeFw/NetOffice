using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.Translation
{
    public class NotifyPropertyChanged : INotifyPropertyChanged
    {
        #region Fields

        protected string _value;
        protected string _value2;

        #endregion

        #region Properties

        /// <summary>
        /// Name
        /// </summary>
        public string Value
        {
            get
            {
                return _value;
            }
            set
            {
                _value = value;
                RaisePropertyChanged("Value");
            }
        }
       
        /// <summary>
        /// Value
        /// </summary>
        public string Value2
        {
            get
            {
                return _value2;
            }
            set
            {
                _value2 = value;
                RaisePropertyChanged("Value2");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Create a deep copy of the instance
        /// </summary>
        /// <returns>clone</returns>
        public virtual NotifyPropertyChanged Clone()
        {
            throw new NotSupportedException();
        }

        #endregion

        #region INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        protected internal void RaisePropertyChanged(string propertyName)
        {
            if (null != PropertyChanged)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion

        #region Overrides

        /// <summary>
        /// Returns a System.String that represents the instance
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return String.Format("Value:{0}   Value2:{1}", _value, _value2);
        }

        #endregion
    }
}
