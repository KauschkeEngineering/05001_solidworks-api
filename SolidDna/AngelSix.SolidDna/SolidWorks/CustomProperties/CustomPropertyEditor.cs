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
        public enum CustomInfoTypes
        {
            Unknown = swCustomInfoType_e.swCustomInfoUnknown,
            Number = swCustomInfoType_e.swCustomInfoNumber,
            Double = swCustomInfoType_e.swCustomInfoDouble,
            YesOrNo = swCustomInfoType_e.swCustomInfoYesOrNo,
            Text = swCustomInfoType_e.swCustomInfoText,
            Date = swCustomInfoType_e.swCustomInfoDate
        }

        public enum CustomInfoGetResult
        {
            // Cached value was returned
            CachedValue = swCustomInfoGetResult_e.swCustomInfoGetResult_CachedValue,
            // Custom property does not exist
            NotPresent = swCustomInfoGetResult_e.swCustomInfoGetResult_NotPresent,
            // Resolved value was returned
            ResolvedValue = swCustomInfoGetResult_e.swCustomInfoGetResult_ResolvedValue
        }

        public enum CustomInfoSetResult
        {
            // Success
            OK = swCustomInfoSetResult_e.swCustomInfoSetResult_OK,
            // Custom property does not exist
            NotPresent = swCustomInfoSetResult_e.swCustomInfoSetResult_NotPresent,
            // Specified value has an incorrect typeSpecified value has an incorrect type
            TypeMismatch = swCustomInfoSetResult_e.swCustomInfoSetResult_TypeMismatch,
            // ?? not described in SOLIDWORKS API see: https://help.solidworks.com/2020/English/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swCustomInfoSetResult_e.html?verRedirect=1
            // but this is returned if the property is linked for e.g. to parent.
            LinkedProp = swCustomInfoSetResult_e.swCustomInfoSetResult_LinkedProp
        }

        public enum CustomInfoAddResult
        {
            // Success
            AddedOrChanged = swCustomInfoAddResult_e.swCustomInfoAddResult_AddedOrChanged,
            // Failed to add the custom property
            GenericFail = swCustomInfoAddResult_e.swCustomInfoAddResult_GenericFail,
            // Existing custom property with the same name has a different type
            MismatchAgainstExistingType = swCustomInfoAddResult_e.swCustomInfoAddResult_MismatchAgainstExistingType,
            //Specified value of the custom property does not match the specified type
            MismatchAgainstSpecifiedType = swCustomInfoAddResult_e.swCustomInfoAddResult_MismatchAgainstSpecifiedType,
        }

        public enum CustomPropertyAddOption
        {
            // Add the custom property only if it is new
            OnlyIfNew = swCustomPropertyAddOption_e.swCustomPropertyOnlyIfNew,
            // Delete an existing custom property having the same name and add the new custom property
            DeleteAndAdd = swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd,
            // Replace the value of an existing custom property having the same name
            ReplaceValue = swCustomPropertyAddOption_e.swCustomPropertyReplaceValue
        }

        public enum CustomInfoDeleteResult
        {
            // Success
            OK = swCustomInfoDeleteResult_e.swCustomInfoDeleteResult_OK,
            // Custom property does not exist
            NotPresent = swCustomInfoDeleteResult_e.swCustomInfoDeleteResult_NotPresent,
            // ?? not described in SOLIDWORKS API see: http://help.solidworks.com/2020/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swCustomInfoDeleteResult_e.html?verRedirect=1
            // but this is returned if the property is linked for e.g. to parent.
            LinkedProp = swCustomInfoDeleteResult_e.swCustomInfoDeleteResult_LinkedProp
        }

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
        public Tuple<CustomInfoGetResult, string> GetCustomPropertyValue(string name, bool resolve = false)
        {
            // TODO: Add error checking and exception catching

            // Get custom property
            var result = BaseObject.Get5(name, false, out var val, out var resolvedVal, out var wasResolved);

            // Return desired result
            return resolve ? new Tuple<CustomInfoGetResult, string>((CustomInfoGetResult)result, resolvedVal) : new Tuple<CustomInfoGetResult, string>((CustomInfoGetResult)result, val);
        }

        /// <summary>
        /// Sets the value of a custom property by name
        /// </summary>
        /// <param name="name">The name of the custom property</param>
        /// <param name="value">The value of the custom property</param>
        /// <param name="type">The type of the custom property</param>
        /// <returns></returns>
        public CustomInfoSetResult SetCustomPropertyValue(string name, string value)
        {
            // TODO: Add error checking and exception catching

            // NOTE: We use Add here to create a property if one doesn't exist
            //       I feel this is the expected behaviour of Set
            //
            //       To mimic the Set behaviour of the SolidWorks API
            //       Simply do CustomPropertyExists() to check first if it exists
            //
            return (CustomInfoSetResult)BaseObject.Set2(name, value);
        }

        /// <summary>
        /// Add a new custom property with the given name, type and value
        /// </summary>
        /// <param name="name">The name of the custom property</param>
        /// <param name="type">The type of the custom property</param>
        /// <param name="option">The option how to handle existing custom property</param>
        /// <param name="value">The value of the custom property</param>
        /// <returns></returns>
        public CustomInfoAddResult AddCustomPropertyValue(string name, CustomInfoTypes type, string value, CustomPropertyAddOption option)
        {
            // TODO: Add error checking and exception catching

            // NOTE: We use Add here to create a property if one doesn't exist
            //       I feel this is the expected behaviour of Set
            //
            //       To mimic the Set behaviour of the SolidWorks API
            //       Simply do CustomPropertyExists() to check first if it exists
            //

            return (CustomInfoAddResult)BaseObject.Add3(name, (int)type, value, (int)option);
        }

        /// <summary>
        /// Deletes a custom property by name
        /// </summary>
        /// <param name="name">The name of the custom property</param>
        public CustomInfoDeleteResult DeleteCustomProperty(string name)
        {
            // TODO: Add error checking and exception catching

            return (CustomInfoDeleteResult)BaseObject.Delete2(name);
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

        public CustomInfoTypes GetTypeOfProperty(string name)
        {
            if (BaseObject != null)
                return (CustomInfoTypes)BaseObject.GetType2(name);
            return CustomInfoTypes.Unknown;
        }

    }
}
