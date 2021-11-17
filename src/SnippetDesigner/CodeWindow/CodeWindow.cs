using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Utilities;

namespace Microsoft.SnippetDesigner
{
    /// <summary>
    /// The code window which is held inside the snippet codeWindowHost form
    /// </summary>
    public partial class CodeWindow : UserControl, ISnippetCodeWindow
    {
        private const string CodeSnippetContentTypeString = "codesnippet";
        private readonly IContentTypeRegistryService contentTypeService;
        private readonly IVsEditorAdaptersFactoryService editorAdapterFactoryService;
        private IVsTextBuffer bufferAdapter;
        private IContentType codeSnippetContentType;
        private ICodeWindowHost codeWindowHost;
        private uint cookieTextLineEvents;
        private uint cookieTextViewEvents;
        private IntPtr hWndCodeWindow;
        private bool isHandleCreated;
        private bool isTextInitialized;
        private SnippetEditor snippetEditor;
        private IConnectionPoint textLinesEventsConnectionPoint;
        private IConnectionPoint textViewEventsConnectionPoint;
        private IVsTextView viewAdapter;

        /// <summary>
        /// Constructor for the code window which is a user snippetExplorerForm that hosts a vscodewindow
        /// </summary>
        public CodeWindow()
        {
            InitializeComponent();

            if (SnippetDesignerPackage.Instance == null)
            {
                return;
            }

            editorAdapterFactoryService =
                SnippetDesignerPackage.Instance.ComponentModel.GetService<IVsEditorAdaptersFactoryService>();
            SnippetDesignerPackage.Instance.ComponentModel.GetService<ITextSearchService>();
            contentTypeService =
                SnippetDesignerPackage.Instance.ComponentModel.GetService<IContentTypeRegistryService>();

            RegisterCodeSnippetContentType();
        }

        public CodeWindow(IVsEditorAdaptersFactoryService editorAdapterFactoryService, IContentTypeRegistryService contentTypeService)
        {
            this.editorAdapterFactoryService = editorAdapterFactoryService;
            this.contentTypeService = contentTypeService;
        }

        //get and set the text in the code window
        public string CodeText
        {
            get
            {
                //   IVsTextLines vsTextLines = OldTextLines;

                if (TextBuffer != null)
                {
                    //string codeText = String.Empty;
                    //int numLines;
                    //int lastLineIndex;

                    return TextBuffer.CurrentSnapshot.GetText();
                    //ErrorHandler.ThrowOnFailure(vsTextLines.GetLastLineIndex(out numLines, out lastLineIndex));
                    //ErrorHandler.ThrowOnFailure(vsTextLines.GetLineText(0, 0, numLines, lastLineIndex, out codeText));
                    //return codeText;
                }
                else
                {
                    return string.Empty;
                }
            }
            set
            {
                try
                {
                    if (TextBufferAdapter != null)
                    {
                        if (isTextInitialized)
                        {
                            SetText(value);
                        }
                        else
                        {
                            if (InitializeText(value))
                            {
                                isTextInitialized = true;
                            }
                        }
                    }
                }
                catch (NullReferenceException ex)
                {
                    SnippetDesignerPackage.Instance.Logger.Log("Text Lines not ready yet?", "CodeWindow::CodeText", ex);
                }
            }
        }

        /// <summary>
        /// The code windows parent CodeWindowHost.
        /// This is the Snippet CodeWindowHost that this code window needs a reference to.
        ///
        /// The Set method is very important.  Our codewindow needs two things to be created.
        /// A reference to the parent codeWindowHost and the window handle.  Since we have both of these we can create
        /// the code window.  This code window can be created during the set here if we have the handle already
        /// or after this set in the OnHandleCreated event since by that point
        /// we will have both the handle and the parent codeWindowHost set
        /// </summary>
        internal ICodeWindowHost CodeWindowHost
        {
            get => codeWindowHost;
            //this may only be called once per snippet instance of the codewindow
            set
            {
                if (codeWindowHost == null)
                {
                    codeWindowHost = value;
                    if (isHandleCreated)
                    {
                        CreateVsCodeWindow();
                        codeWindowHost.SetupContextMenus();
                    }

                    if (codeWindowHost is SnippetEditor)
                    {
                        snippetEditor = value as SnippetEditor;
                    }
                }
            }
        }

        /// <summary>
        /// The number of lines
        /// </summary>
        internal int LineCount
        {
            get
            {
                if (TextBufferAdapter != null)
                {
                    TextBufferAdapter.GetLineCount(out var lineCount);
                    return lineCount;
                }
                else
                {
                    return -1;
                }
            }
        }

        /// <summary>
        /// get the selected text
        /// </summary>
        /// <returns>selected text</returns>
        internal string SelectedText
        {
            get
            {
                if (TextViewAdapter != null)
                {
                    var selectedText = string.Empty;
                    TextViewAdapter.GetSelectedText(out selectedText);
                    return selectedText;
                }
                else
                {
                    return string.Empty;
                }
            }
        }

        /// <summary>
        /// gets or sets the selection text span
        /// </summary>
        /// <returns>selected text</returns>
        internal SnapshotSpan Selection
        {
            get => TextView.Selection.SelectedSpans[0];
            set => TextView.Selection.Select(value, false);
        }

        /// <summary>
        /// get the length of the selection
        /// </summary>
        /// <returns>length of selection</returns>
        internal int SelectionLength => SelectedText.Length;

        internal ITextBuffer TextBuffer
        {
            get
            {
                if (TextView == null)
                {
                    return null;
                }

                return TextView.TextBuffer;
            }
        }

        /// <summary>
        /// The TextLines interface of the codewindows textbuffer
        /// </summary>
        internal IVsTextLines TextBufferAdapter => (IVsTextLines)bufferAdapter;

        internal ITextView TextView
        {
            get
            {
                if (viewAdapter != null)
                {
                    return editorAdapterFactoryService.GetWpfTextView(viewAdapter);
                }
                return null;
            }
        }

        /// <summary>
        /// the primary view for the code window
        /// </summary>
        internal IVsTextView TextViewAdapter => viewAdapter;

        public static bool IsWordChar(char c) => char.IsLetterOrDigit(c) || c == '_';

        public string GetCharacterAtPosition(TextPoint positon)
        {
            var textLines = TextBufferAdapter;
            textLines.GetLineText(positon.Line, positon.Index, positon.Line, positon.Index + 1, out var charAtPos);
            return charAtPos;
        }

        public string GetSpanText(SnapshotSpan span) => span.GetText();

        public string GetWordFromCurrentPosition() => GetWordFromPosition(TextView.Caret.Position.BufferPosition);

        public string GetWordFromPosition(SnapshotPoint positon)
        {
            var span = GetWordSpanFromPosition(positon);

            return span.GetText();
        }

        public SnapshotSpan GetWordTextSpanFromCurrentPosition() => GetWordSpanFromPosition(TextView.Caret.Position.BufferPosition);

        public void InitializeEditor()
        {
            VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            var oleServiceProvider = codeWindowHost.ServiceProvider;
            bufferAdapter = editorAdapterFactoryService.CreateVsTextBufferAdapter(oleServiceProvider, codeSnippetContentType);
            var result = bufferAdapter.InitializeContent("", 0);
            viewAdapter = editorAdapterFactoryService.CreateVsTextViewAdapter(oleServiceProvider);
            ((IVsWindowPane)viewAdapter).SetSite(oleServiceProvider);

            var initView = new[] { new INITVIEW() };
            initView[0].fSelectionMargin = 1;
            initView[0].fWidgetMargin = 0;
            initView[0].fVirtualSpace = 0;
            initView[0].fDragDropMove = 1;
            initView[0].fVirtualSpace = 0;

            uint readOnlyValue = 0;
            if (codeWindowHost.ReadOnlyCodeWindow)
            {
                readOnlyValue = (uint)TextViewInitFlags2.VIF_READONLY;
            }

            var flags =
                (uint)TextViewInitFlags.VIF_SET_SELECTION_MARGIN |
                (uint)TextViewInitFlags.VIF_SET_DRAGDROPMOVE |
                (uint)TextViewInitFlags2.VIF_SUPPRESS_STATUS_BAR_UPDATE |
                (uint)TextViewInitFlags2.VIF_SUPPRESSBORDER |
                (uint)TextViewInitFlags2.VIF_SUPPRESSTRACKCHANGES |
                (uint)TextViewInitFlags2.VIF_SUPPRESSTRACKGOBACK |
                (uint)TextViewInitFlags.VIF_HSCROLL |
                (uint)TextViewInitFlags.VIF_VSCROLL |
                (uint)TextViewInitFlags3.VIF_NO_HWND_SUPPORT |
                readOnlyValue;

            viewAdapter.Initialize(bufferAdapter as IVsTextLines, Handle, flags, initView);
        }

        /// <summary>
        /// Length of a line in the buffer
        /// </summary>
        /// <param name="line">the line</param>
        /// <returns>the length of the line</returns>
        internal int LineLength(int line)
        {
            if (TextBufferAdapter != null)
            {
                TextBufferAdapter.GetLengthOfLine(line, out var lineLength);
                return lineLength;
            }
            else
            {
                return -1;
            }
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
                if (cookieTextLineEvents != 0)
                {
                    textLinesEventsConnectionPoint.Unadvise(cookieTextLineEvents);
                }
                if (cookieTextViewEvents != 0)
                {
                    textViewEventsConnectionPoint.Unadvise(cookieTextViewEvents);
                }

                if (viewAdapter != null)
                {
                    ((IVsWindowPane)viewAdapter).ClosePane();
                }
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// This gets called once this snippetExplorerForm has recieved its windows handle
        /// This is important since we need this inorder to create the vs code window
        /// since we create the code window pane ourselves
        ///
        /// If this gets called before codeWindowHost is set then we must not create the code window yet
        /// it will be created when codeWindowHost gets set
        /// </summary>
        /// <param name="e"></param>
        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            if (codeWindowHost != null) //do we know our parent window?
            {
                //if so create code window and set up the context menus for this code window
                CreateVsCodeWindow();
                codeWindowHost.SetupContextMenus();
            }

            //mark that the handle is created
            isHandleCreated = true;
        }

        private static SnapshotSpan GetWordSpanFromPosition(SnapshotPoint positon)
        {
            if (positon.Position >= positon.Snapshot.Length)
            {
                return new SnapshotSpan(positon, 0);
            }

            var charAtPos = positon.GetChar();
            var text = positon.Snapshot.GetText();
            var lineSnapshot = positon.Snapshot.GetLineFromPosition(positon);
            int left, right;
            left = right = positon.Position;

            if (IsWordChar(charAtPos))
            {
                while (left - 1 >= lineSnapshot.Start.Position && IsWordChar(text[left - 1]))
                {
                    left--;
                }
                while (right + 1 < lineSnapshot.End.Position && IsWordChar(text[right + 1]))
                {
                    right++;
                }

                return new SnapshotSpan(positon.Snapshot, left, right + 1 - left);
            }
            else
            {
                return new SnapshotSpan(positon, 1);
            }
        }

        /// <summary>
        /// Keep the code window we created the right size when the window gets resized.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CodeWindow_SizeChanged(object sender, EventArgs e) => NativeMethods.SetWindowPos(hWndCodeWindow,
                                       IntPtr.Zero,
                                       0,
                                       0,
                                       Width,
                                       Height,
                                       0);

        /// <summary>
        /// Constructs the IVsCodeWindow and attaches an IVsTextBuffer
        /// </summary>
        /// <returns>S_OK if success</returns>
        private int CreateVsCodeWindow()
        {
            VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            InitializeEditor();

            var hr = VSConstants.S_OK;

            //create the codewindow as the size of the snippetExplorerForm its in
            hr = ((IVsWindowPane)viewAdapter).CreatePaneWindow(Handle, 0, 0, Parent.Size.Width, Parent.Size.Height, out hWndCodeWindow);

            //we are only getting events if the codewindowhost is the snippet editor)
            if (snippetEditor != null)
            {
                // sink IVsTextViewEvents, so we can determine when a VsCodeWindow object actually has the focus.
                var connptCntr = (IConnectionPointContainer)viewAdapter;
                var riid = typeof(IVsTextViewEvents).GUID;

                //find the desired connection point
                connptCntr.FindConnectionPoint(ref riid, out textViewEventsConnectionPoint);
                //connect to this connection point to be advised of changes
                textViewEventsConnectionPoint.Advise(snippetEditor, out cookieTextViewEvents);

                // sink IVsTextLineEvents, so we can determine when the buffer is changed
                connptCntr = (IConnectionPointContainer)bufferAdapter;
                riid = typeof(IVsTextLinesEvents).GUID;

                //find the desired connection point
                connptCntr.FindConnectionPoint(ref riid, out textLinesEventsConnectionPoint);
                //connect to this connection point to be advised of changes
                textLinesEventsConnectionPoint.Advise(snippetEditor, out cookieTextLineEvents);
            }

            return hr;
        }

        private bool InitializeText(string newText)
        {
            newText = newText ?? "";
            if (ErrorHandler.Failed(bufferAdapter.InitializeContent(newText, newText.Length)))
            {
                return false;
            }

            return true;
        }

        private void RegisterCodeSnippetContentType()
        {
            codeSnippetContentType = contentTypeService.GetContentType(CodeSnippetContentTypeString);
            if (codeSnippetContentType == null)
            {
                codeSnippetContentType = contentTypeService.AddContentType(CodeSnippetContentTypeString, new List<string> { "code" });
            }
        }

        /// <summary>
        /// This is a helper routine to replace the contents of the Text Buffer with "newText".
        /// This function handles the interop of passing the string as a block of memory
        /// via an IntPtr. It uses a CoTaskMemAlloc to allocate a block of memory at a fixed
        /// location.
        /// </summary>
        /// <param name="newText"></param>
        private void SetText(string newText) => TextBuffer.Replace(new Span(0, TextBuffer.CurrentSnapshot.Length), newText);
    }
}
