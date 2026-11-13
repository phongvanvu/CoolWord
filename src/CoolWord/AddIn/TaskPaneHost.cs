using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using CoolWord.Core;
using CoolWord.UI;

namespace CoolWord.AddIn
{
    [ComVisible(true)]
    [Guid("7C5E8A12-3D4F-4E9B-8A1C-2F3E4D5C6B7A")]
    [ProgId("CoolWord.TaskPaneHost")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class TaskPaneHost : UserControl
    {
        private ElementHost _elementHost;
        private MainTaskPaneControl _wpfControl;

        public TaskPaneHost()
        {
            InitializeHost();
        }

        private void InitializeHost()
        {
            SuspendLayout();
            _elementHost = new ElementHost
            {
                Dock = DockStyle.Fill
            };
            Controls.Add(_elementHost);
            ResumeLayout(false);
        }

        public void InitializeControl(IWordDocumentService wordService)
        {
            if (_wpfControl == null)
            {
                _wpfControl = new MainTaskPaneControl(wordService);
                _elementHost.Child = _wpfControl;
            }
        }

        public MainTaskPaneControl WpfControl
        {
            get { return _wpfControl; }
        }
    }
}
