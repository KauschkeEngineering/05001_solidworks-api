using SolidWorks.Interop.swdocumentmgr;

namespace AngelSix.SolidDna
{
    public class DrawingSheet : SolidDnaObject<SwDMSheet2>
    {
        #region Protected Members


        #endregion

        #region Public Properties

        public string Name => BaseObject.Name;

        #endregion

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public DrawingSheet(SwDMSheet2 model) : base(model)
        {
        }

        #endregion

        #region Sheet Methods

        #endregion

        #region View Methods

        #endregion
    }
}
