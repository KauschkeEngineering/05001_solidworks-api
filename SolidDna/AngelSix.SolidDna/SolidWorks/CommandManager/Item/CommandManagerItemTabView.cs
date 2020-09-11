using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// CommandManager tab button text display.
    /// The view style of a <see cref="CommandManagerItem"/> for a Tab (the large icons above opened files).
    /// References <see cref="swCommandTabButtonTextDisplay_e"/>
    /// </summary>
    public enum CommandManagerItemTabView
    {
        /// <summary>
        /// The item is not shown in the tab
        /// </summary>
        None = 0,

        /// <summary>
        /// The item is shown with an icon only
        /// </summary>
        IconOnly = swCommandTabButtonTextDisplay_e.swCommandTabButton_NoText,

        /// <summary>
        /// The item is shown with the icon, then the text below it
        /// </summary>
        IconWithTextBelow = swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow,

        /// <summary>
        /// The item is shown with the icon then the text to the right
        /// </summary>
        IconWithTextAtRight = swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal
    }
}
