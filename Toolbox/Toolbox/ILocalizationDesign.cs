using System;
using System.Collections.Generic;

namespace NetOffice.DeveloperToolbox
{
    /// <summary>
    /// Additional support for the language editor
    /// </summary>
    public interface ILocalizationDesign
    {
        /// <summary>
        /// make hidden controls visible to help the user
        /// </summary>
        /// <param name="strings">lcid</param>
        /// <param name="strings">parentComponentName</param>
        void EnableDesignView(int lcid, string parentComponentName);

        /// <summary>
        /// Update control captions
        /// </summary>
        /// <param name="strings">target language strings</param>
        void Localize(Translation.ItemCollection strings);

        /// <summary>
        /// Update control caption
        /// </summary>
        /// <param name="name">name of the control</param>
        /// <param name="text">caption for the control</param>
        void Localize(string name, string text);

        /// <summary>
        /// Returns current caption
        /// </summary>
        /// <param name="name">name of the control</param>
        /// <returns>current control text</returns>
        string GetCurrentText(string name);

        /// <summary>
        /// To localize non-control instances (toolstrip, etc.)
        /// </summary>
        System.ComponentModel.IContainer Components { get; }

        /// <summary>
        /// Caption in the Language editor.
        /// </summary>
        string NameLocalization { get; }

        /// <summary>
        /// Additional child controls
        /// </summary>
        IEnumerable<ILocalizationChildInfo> Childs { get; }
    }

    /// <summary>
    /// Optional childs informations
    /// </summary>
    public interface ILocalizationChildInfo
    {
        /// <summary>
        /// Caption in the Language editor.
        /// </summary>
        string NameLocalization { get; }

        /// <summary>
        /// Type of the child control
        /// </summary>
        Type TypeLocalization { get; }
    }
}
