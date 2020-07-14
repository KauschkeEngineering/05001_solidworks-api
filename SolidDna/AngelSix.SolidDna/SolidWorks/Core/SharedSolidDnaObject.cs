using System.Runtime.InteropServices;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Represents a core SolidDNA object, that is disposable
    /// and needs a COM object disposing cleanly on disposal
    /// 
    /// NOTE: Use this shared type if another part of the application may have access to this
    ///       same COM object and the life-cycle for each reference is managed independently
    /// </summary>
    public class SharedSolidDnaObject<T> : SolidDnaObject<T>
    {
        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="comObject">The COM object to wrap</param>
        public SharedSolidDnaObject(T comObject) : base(comObject)
        {

        }

        #endregion

        #region Dispose

        /// <summary>
        /// Disposal
        /// </summary>
        public override void Dispose()
        {
            if (BaseObject == null)
                return;

            // Do any specific disposal
            SolidDnaObjectDisposal.Dispose<T>(BaseObject);

            // COM release object
            // TODO: Add references why to not use Marshal.FinalReleaseComObject here
            // if called here it will lead to several problems
            // like in my case the stand alone app and windows was completly freezing
            // becuase the servicehost remote procedure call is stuck
            //Marshal.ReleaseComObject(BaseObject);

            // Clear reference
            BaseObject = default(T);
        }

        #endregion
    }
}
