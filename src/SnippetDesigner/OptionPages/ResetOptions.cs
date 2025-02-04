using System;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.VisualStudio.Shell;

namespace Microsoft.SnippetDesigner.OptionPages
{
    /// <summary>
    /// Reset options page
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("0F9D79A2-121F-484e-8DE9-62A1EF289301")]
    public class ResetOptions : DialogPage
    {
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        protected override IWin32Window Window => new ResetOptionsControl
        {
            Location = new Point(0, 0)
        };
    }
}
