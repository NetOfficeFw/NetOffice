using System;
using System.Windows.Forms;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.Translation
{
    /// <summary>
    /// Represents a localizable component
    /// </summary>
    internal class LocalizableCompoment : NotifyPropertyChanged
    {
        #region Fields

        private Type _controlType;
        private UserControl _control;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class as root component
        /// </summary>
        /// <param name="parent">parent language</param>
        /// <param name="name">name of the component</param>
        /// <param name="controlType">component type to create an instance from</param>
        internal LocalizableCompoment(ToolLanguage parent, string name, Type controlType)
        {
            Parent = parent;
            _value = name;
            _controlType = controlType;
            ControlRessources = new ItemCollection();

            string[] names = RessourceTableAttribute.GetRessourceNames(controlType);
            Dictionary<string, string> values = RessourceTableAttribute.GetRessourceValues(Design, parent.LCID);
            foreach (var resName in names)
            {
                string resValue ="";
                values.TryGetValue(resName, out resValue);

                Control ctrl = Translator.TryGetControl(Design, resName);
                Controls.Text.AdvRichTextBox advrichText = ctrl as Controls.Text.AdvRichTextBox;
                if (null != advrichText)
                {
                    ControlRessources.Add(new LocalizableWideString(resName, resValue));
                }
                else
                {
                    RichTextBox richBox = ctrl as RichTextBox;
                    if (null != richBox)
                    {
                        ControlRessources.Add(new LocalizableRTFString(resName, resValue));
                    }
                    else
                    {
                        TextBox textBox = ctrl as TextBox;
                        if (null != textBox && textBox.Multiline)
                            ControlRessources.Add(new LocalizableWideString(resName, resValue));
                        else
                            ControlRessources.Add(new LocalizableString(resName, resValue));
                    }

                }
            }
        }

        /// <summary>
        /// Creates an instance as sub component
        /// </summary>
        /// <param name="parent">parent language</param>
        /// <param name="parentComponentName">name of the parent component</param>
        /// <param name="name">name of the component</param>
        /// <param name="controlType">component type to create an instance from</param>
        internal LocalizableCompoment(ToolLanguage parent, string parentComponentName, string name, Type controlType)
        {
            Parent = parent;
            _value = name;
            _value2 = parentComponentName;
            _controlType = controlType;
            ControlRessources = new ItemCollection();

            string[] names = RessourceTableAttribute.GetRessourceNames(controlType);
            Dictionary<string, string> values = RessourceTableAttribute.GetRessourceValues(Design, parent.LCID);
            foreach (var resName in names)
            {
                string resValue = "";
                values.TryGetValue(resName, out resValue);

                Control ctrl = Translator.TryGetControl(Design, resName);
                Controls.Text.AdvRichTextBox advrichText = ctrl as Controls.Text.AdvRichTextBox;
                if (null != advrichText)
                {
                    ControlRessources.Add(new LocalizableWideString(resName, resValue));
                }
                else
                {
                    RichTextBox richBox = ctrl as RichTextBox;
                    if (null != richBox)
                    {
                        ControlRessources.Add(new LocalizableRTFString(resName, resValue));
                    }
                    else
                    {
                        TextBox textBox = ctrl as TextBox;
                        if (null != textBox && textBox.Multiline)
                            ControlRessources.Add(new LocalizableWideString(resName, resValue));
                        else
                            ControlRessources.Add(new LocalizableString(resName, resValue));
                    }

                }
            }
        }

        #endregion

        #region Properties

        /// <summary>
        /// Parent Language
        /// </summary>
        internal ToolLanguage Parent { get; private set; }

        /// <summary>
        /// Attribute from the component class
        /// </summary>
        internal RessourceTableAttribute Attribute
        {
            get
            {
                object[] obj = _controlType.GetCustomAttributes(typeof(RessourceTableAttribute), false);
                RessourceTableAttribute attrib = obj[0] as RessourceTableAttribute;
                return attrib;
            }
        }

        /// <summary>
        /// Component instance in design mode
        /// </summary>
        internal UserControl Design
        {
            get             
            {
                if (null == _control)
                {
                    _control = Activator.CreateInstance(_controlType) as UserControl;
                    ILocalizationDesign designSupport = _control as ILocalizationDesign;
                    if (null != designSupport)
                        designSupport.EnableDesignView(Parent.LCID, Value2);
                }
                return _control;
            }
        }

        /// <summary>
        /// Localizable resources from the component
        /// </summary>
        internal ItemCollection ControlRessources { get; private set; }

        #endregion
    }
}
