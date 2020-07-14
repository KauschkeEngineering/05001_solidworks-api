﻿using System;
using static AngelSix.SolidDna.CustomPropertyEditor;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// A custom property of a model that can be edit directly
    /// </summary>
    public class CustomProperty
    {
        public Type DataType
        {
            get
            {
                if (mEditor != null && Name.Equals("") == false)
                {
                    switch (mEditor.GetTypeOfProperty(Name))
                    {
                        case CustomPropertyTypes.Unknown:
                            return typeof(object);
                        case CustomPropertyTypes.Number:
                            return typeof(long);
                        case CustomPropertyTypes.Double:
                            return typeof(double);
                        case CustomPropertyTypes.YesOrNo:
                            return typeof(bool);
                        case CustomPropertyTypes.Text:
                            return typeof(string);
                        case CustomPropertyTypes.Date:
                            return typeof(DateTime);
                        default:
                            return null;
                    }
                }
                return null;
            }
        }
        #region Private Members

        /// <summary>
        /// The editor used for this custom property
        /// </summary>
        private CustomPropertyEditor mEditor;

        #endregion

        #region Public Properties

        /// <summary>
        /// The name of the custom property
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// The value of the custom property
        /// </summary>
        public string Value
        {
            get => mEditor.GetCustomPropertyValue(Name).Item2;
            set => mEditor.SetCustomPropertyValue(Name, value);
        }

        /// <summary>
        /// The resolved value of the custom property
        /// </summary>
        public string ResolvedValue => mEditor.GetCustomPropertyValue(Name, true).Item2;

        #endregion

        #region Constructor

        /// <summary>
        /// Default Constructor
        /// </summary>
        public CustomProperty(CustomPropertyEditor editor, string name)
        {
            // Store reference
            mEditor = editor;

            // Set name
            Name = name;
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Deletes this custom property
        /// </summary>
        public CustomPropertyDeleteResult Delete()
        {
            return mEditor.DeleteCustomProperty(Name);
        }

        #endregion
    }
}
