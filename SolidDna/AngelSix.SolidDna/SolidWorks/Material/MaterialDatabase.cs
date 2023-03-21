
using System.IO;

namespace AngelSix.SolidDna
{
    public class MaterialDatabase
    {

        #region Constants

        public const string DEFAULT_EXTENSION = ".sldmat";
        public const string FILE_FILTER = " (" + DEFAULT_EXTENSION + ")|*" + DEFAULT_EXTENSION;

        #endregion

        #region Private Members

        #endregion

        #region Properties

        public string Name
        {
            get
            {
                if (!FullPath.Equals(""))
                {
                    var split = FullPath.Split(Path.DirectorySeparatorChar);
                    if (split.Length > 0)
                    {
                        return split[split.Length - 1].Replace(DEFAULT_EXTENSION, "");
                    }
                    return FullPath;
                }
                return "";
            }
        }

        public string FullPath { get; private set; }
        public bool Found 
        { 
            get
            {
                if (Name.Equals("") || Name.Equals(FullPath))
                {
                    return false;
                }
                return true;
            }
        }

        public bool IsNone { get; } = false;

        #endregion

        #region Constructor

        public MaterialDatabase(string fullPath)
        {
            FullPath = fullPath;
            if (fullPath.Equals("Keine"))
            {
                IsNone = true;
            }
        }

        #endregion

        #region Private Methods

        #endregion

        #region Public Methods

        #endregion

    }
}
