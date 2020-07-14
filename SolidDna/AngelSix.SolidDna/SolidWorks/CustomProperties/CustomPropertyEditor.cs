using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Represents a SolidWorks custom property manager for a model
    /// </summary>
    public class CustomPropertyEditor : SolidDnaObject<CustomPropertyManager>
    {
        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public CustomPropertyEditor(CustomPropertyManager model) : base(model)
        {

        }

        #endregion

        /// <summary>
        /// Checks if a custom property exists
        /// </summary>
        /// <param name="name">The name of the custom property</param>
        /// <returns></returns>
        public bool CustomPropertyExists(string name)
        {
            // TODO: Add error checking and exception catching

            return GetCustomProperties().Any(f => string.Equals(f.Name, name, System.StringComparison.InvariantCultureIgnoreCase));
        }

        /// <summary>
        /// Gets the value of a custom property by name
        /// </summary>
        /// <param name="name">The name of the custom property</param>
        /// <param name="resolve">True to resolve the custom property value</param>
        /// <returns></returns>
        public string GetCustomProperty(string name, bool resolve = false)
        {
            // TODO: Add error checking and exception catching

            // Get custom property
            BaseObject.Get5(name, false, out var val, out var resolvedVal, out var wasResolved);

            // Return desired result
            return resolve ? resolvedVal : val;
        }

        /// <summary>
        /// Gets the value of a custom property by name
        /// </summary>
        /// <param name="name">The name of the custom property</param>
        /// <param name="resolve">True to resolve the custom property value</param>
        /// <returns></returns>
        public Tuple<CustomPropertyGetResult, string> GetCustomPropertyValue(string name, bool resolve = false)
        {
            // TODO: Add error checking and exception catching

            // Get custom property
            var result = BaseObject.Get5(name, false, out var val, out var resolvedVal, out var wasResolved);

            // Return desired result
            return resolve ? new Tuple<CustomPropertyGetResult, string>((CustomPropertyGetResult)result, resolvedVal) : new Tuple<CustomPropertyGetResult, string>((CustomPropertyGetResult)result, val);
        }

        /// <summary>
        /// Sets the value of a custom property by name
        /// </summary>
        /// <param name="name">The name of the custom property</param>
        /// <param name="value">The value of the custom property</param>
        /// <param name="type">The type of the custom property</param>
        /// <returns></returns>
        public void SetCustomProperty(string name, string value, swCustomInfoType_e type = swCustomInfoType_e.swCustomInfoText)
        {
            // TODO: Add error checking and exception catching

            // NOTE: We use Add here to create a property if one doesn't exist
            //       I feel this is the expected behaviour of Set
            //
            //       To mimic the Set behaviour of the SolidWorks API
            //       Simply do CustomPropertyExists() to check first if it exists
            //

            // Set new one
            BaseObject.Add3(name, (int)type, value, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
        }

        /// <summary>
        /// Sets the value of a custom property by name
        /// </summary>
        /// <param name="name">The name of the custom property</param>
        /// <param name="value">The value of the custom property</param>
        /// <param name="type">The type of the custom property</param>
        /// <returns></returns>
        public CustomPropertySetResult SetCustomPropertyValue(string name, string value)
        {
            // TODO: Add error checking and exception catching

            // NOTE: We use Add here to create a property if one doesn't exist
            //       I feel this is the expected behaviour of Set
            //
            //       To mimic the Set behaviour of the SolidWorks API
            //       Simply do CustomPropertyExists() to check first if it exists
            //
            return (CustomPropertySetResult)BaseObject.Set2(name, value);
        }

        /// <summary>
        /// Add a new custom property with the given name, type and value
        /// </summary>
        /// <param name="name">The name of the custom property</param>
        /// <param name="type">The type of the custom property</param>
        /// <param name="option">The option how to handle existing custom property</param>
        /// <param name="value">The value of the custom property</param>
        /// <returns></returns>
        public CustomPropertyAddResult AddCustomPropertyValue(string name, CustomPropertyTypes type, string value, CustomPropertyAddOption option)
        {
            // TODO: Add error checking and exception catching

            // NOTE: We use Add here to create a property if one doesn't exist
            //       I feel this is the expected behaviour of Set
            //
            //       To mimic the Set behaviour of the SolidWorks API
            //       Simply do CustomPropertyExists() to check first if it exists
            //

            return (CustomPropertyAddResult)BaseObject.Add3(name, (int)type, value, (int)option);
        }

        /// <summary>
        /// Deletes a custom property by name
        /// </summary>
        /// <param name="name">The name of the custom property</param>
        public void DeleteCustomProperty(string name)
        {
            // TODO: Add error checking and exception catching

            BaseObject.Delete2(name);
        }

        /// <summary>
        /// Deletes a custom property by name
        /// </summary>
        /// <param name="name">The name of the custom property</param>
        public CustomPropertyDeleteResult DeleteCustomPropertyWithResult(string name)
        {
            // TODO: Add error checking and exception catching

            return (CustomPropertyDeleteResult)BaseObject.Delete2(name);
        }

        /// <summary>
        /// Gets a list of all custom properties
        /// </summary>
        /// <returns></returns>
        public List<CustomProperty> GetCustomProperties()
        {
            // TODO: Add error checking and exception catching

            // Create an empty list
            var list = new List<CustomProperty>();

            // Get all properties
            var names = (string[])BaseObject.GetNames();

            // Create custom property objects for each
            if (names?.Length > 0)
                list.AddRange(names.Select(name => new CustomProperty(this, name)).ToList());

            // Return the list
            return list;
        }

        /// <summary>
        /// Gets a specified custom properties
        /// </summary>
        /// <returns></returns>
        public Type GetCustomPropertyType(string name)
        {
            // TODO: Add error checking and exception catching

            // Create an empty list
            var list = new List<CustomProperty>();

            // Get all properties
            var names = (string[])BaseObject.GetNames();

            return new CustomProperty(this, names.FirstOrDefault(propertyName => propertyName.Equals(name))).DataType;
        }


        public List<string> GetCustomPropertyNames()
        {
            if (BaseObject != null && BaseObject.Count > 0)
                return ((string[])BaseObject.GetNames()).ToList();
            return new List<string>();
        }

        public CustomPropertyTypes GetTypeOfProperty(string name)
        {
            if (BaseObject != null)
                return (CustomPropertyTypes)BaseObject.GetType2(name);
            return CustomPropertyTypes.Unknown;
        }

    }
}
