using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Operations;
using SnippetDesignerComponents;
using SnippetLibrary;

namespace Microsoft.SnippetDesigner
{
    /// <summary>
    /// The form which represents the GUI of the snippet editor.
    /// It implements ISnippetEditor which defines what properties and functions a snippet editor
    /// must use.
    /// </summary>
    [ComVisible(true)]
    public partial class SnippetEditorForm :
        UserControl,
        ISnippetEditor
    {
        //hash which maps the display names of the languages to the path to their user snippet directory
        internal readonly Dictionary<string, string> snippetDirectories = null;

        //if the user typed a single character then store it here so we know what it is
        //otherwise its null
        protected string lastCharacterEntered;

        protected ILogger logger;

        private const string EndMarkerFormat = "{0}end{0}";

        private const string SelectedMarkerFormat = "{0}selected{0}";

        private readonly IList<string> reservedReplacements = new List<string>();

        private readonly CollectionWithEvents<AlternativeShortcut> snippetAlternativeShortcuts = new CollectionWithEvents<AlternativeShortcut>();

        private readonly CollectionWithEvents<string> snippetImports = new CollectionWithEvents<string>();

        private readonly CollectionWithEvents<string> snippetKeywords = new CollectionWithEvents<string>();

        private readonly CollectionWithEvents<string> snippetReferences = new CollectionWithEvents<string>();

        private readonly CollectionWithEvents<SnippetType> snippetTypes = new CollectionWithEvents<SnippetType>();

        private readonly ITextSearchService textSearchService;

        private readonly Regex validPotentialReplacementRegex;

        //this is the id column value of the row you have currently selected
        //this value is use so that we know which ids to give the different higlighting
        private string currentlySelectedId;

        private bool formDirty;

        //the value of the id cell you entered before edit
        //the purpose of this is so if you modify a replcement id in the gridview we can know which ids in the codewindow
        //to update
        private string previousIDValue;

        //store the last language selected so we know when to remove adornments
        private string previousLanguageSelected = string.Empty;

        private string snippetAuthor = string.Empty;

        private string snippetDelimiter = Snippet.DefaultDelimiter;

        private string snippetDescription = string.Empty;

        //Snippy Library Access Code
        private SnippetFile snippetFile; //represents an instance of this snippet application

        private string snippetHelpUrl = string.Empty;
        private int snippetIndex; // index of the snippet in the snippetFile
        private string snippetKind = string.Empty;

        private string snippetShortcut = string.Empty;

        //header snippet data
        private string snippetTitle = string.Empty;

        public SnippetEditorForm()
        {
            snippetImports.CollectionChanged += snippet_CollectionChanged;
            snippetReferences.CollectionChanged += snippet_CollectionChanged;
            snippetAlternativeShortcuts.CollectionChanged += snippet_CollectionChanged;
            reservedReplacements.Add("end");
            reservedReplacements.Add("selected");

            textSearchService = SnippetDesignerPackage.Instance.ComponentModel.GetService<ITextSearchService>();

            validPotentialReplacementRegex = SnippetRegexPatterns.BuildValidPotentialReplacementRegex();
            snippetDirectories = SnippetDirectories.Instance.Value.UserSnippetDirectories;
        }

        /// <summary>
        /// the current snippet we are working with in the snippet file
        /// </summary>
        public Snippet ActiveSnippet { get; set; }

        /// <summary>
        /// Get the codewindow snippetExplorerForm
        /// </summary>
        public CodeWindow CodeWindow => snippetCodeWindow;

        /// <summary>
        /// This is the id column value of the row you have currently selected
        /// this is used for showing different markers for the active replcement
        /// </summary>
        public string CurrentlySelectedId => currentlySelectedId;

        public bool FormLoadFinished { get; set; }

        /// <summary>
        /// Is this form in a diry state
        /// </summary>
        public bool IsFormDirty
        {
            get => FormLoadFinished ? formDirty : false;
            set => formDirty = value;
        }

        public void CreateReplacementFromSelection()
        {
            string selectedText;
            var selectionLength = CodeWindow.Selection.Length;
            if (selectionLength > 0)
            {
                //trim any replacement symbols or spaces
                selectedText = CodeWindow.Selection.GetText().Trim();
            }
            else
            {
                selectedText = CodeWindow.GetWordFromCurrentPosition();
            }

            if (!IsValidReplaceableText(selectedText))
            {
                AlertInvalidReplacement(selectedText);
            }
            else if (CreateReplacement(selectedText) && selectionLength > 0)
            {
                var newSnapshot = CodeWindow.Selection.Snapshot.TextBuffer.CurrentSnapshot;
                // CodeWindow.Selection = new SnapshotSpan(newSnapshot, new Span(CodeWindow.Selection.Start.Position, CodeWindow.Selection.End.Position + (StringConstants.SymbolReplacement.Length * 2)));
            }
        }

        public void InsertEndMarker()
        {
            ReplaceAll(string.Format(EndMarkerFormat, SnippetDelimiter), "", false);
            var caretPosition = CodeWindow.TextView.Caret.Position.BufferPosition.Position;
            CodeWindow.TextBuffer.Insert(caretPosition, string.Format(EndMarkerFormat, SnippetDelimiter));
        }

        public void InsertSelectedMarker()
        {
            ReplaceAll(string.Format(SelectedMarkerFormat, SnippetDelimiter), "", false);
            var caretPosition = CodeWindow.TextView.Caret.Position.BufferPosition.Position;
            CodeWindow.TextBuffer.Insert(caretPosition, string.Format(SelectedMarkerFormat, SnippetDelimiter));
        }

        /// <summary>
        /// Load the snippet
        ///
        /// throws IOExcpetion if load failure
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public bool LoadSnippet(string fileName)
        {
            try
            {
                //load the snipept into memory
                snippetFile = new SnippetFile(fileName);

                snippetIndex = 0;

                //set this snippet as the active snippet
                ActiveSnippet = snippetFile.Snippets[snippetIndex];

                //populate the gui with this snippets information
                PullFieldsFromActiveSnippet();
                //indicate that this snippet is not dirty
                IsFormDirty = false;

                FormLoadFinished = true;
            }
            catch (IOException) //abort loading snippet, fail program
            {
                //since an io error occured
                throw;
            }

            return true;
        }

        public void MakeClickedReplacementActive()
        {
            //see if the person clicked inside of a replacement and return its span
            if (GetClickedOnReplacementSpan(out var currentWordSpan))
            {
                var currentWord = CodeWindow.GetSpanText(currentWordSpan);

                foreach (DataGridViewRow row in replacementGridView.Rows)
                {
                    if ((string)row.Cells[StringConstants.ColumnID].Value == currentWord)
                    {
                        replacementGridView.ClearSelection();
                        row.Selected = true;
                        currentlySelectedId = row.Cells[StringConstants.ColumnID].Value as string;
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Takes data from in memory snippet file and populates the gui form
        /// </summary>
        public void PullFieldsFromActiveSnippet()
        {
            SnippetDelimiter = ActiveSnippet.CodeDelimiterAttribute;
            SnippetTitle = ActiveSnippet.Title;
            SnippetAuthor = ActiveSnippet.Author;
            SnippetDescription = ActiveSnippet.Description;
            SnippetHelpUrl = ActiveSnippet.HelpUrl;
            SnippetShortcut = ActiveSnippet.Shortcut;
            SnippetKeywords = new CollectionWithEvents<string>(ActiveSnippet.Keywords);
            SnippetAlternativeShortcuts = new CollectionWithEvents<AlternativeShortcut>(ActiveSnippet.AlternativeShortcuts);

            SnippetTitles = new CollectionWithEvents<string>(GetSnippetTitles());

            if (ActiveSnippet.SnippetTypes.Count() <= 0)
            {
                //if no type specified then make it expansion by default
                snippetTypes.Add(new SnippetType(StringConstants.SnippetTypeExpansion));
            }
            else
            {
                SnippetTypes = new CollectionWithEvents<SnippetType>(ActiveSnippet.SnippetTypes);
            }

            SnippetReplacements = new CollectionWithEvents<Literal>(ActiveSnippet.Literals);

            //code - for some unknown reason this must be done before language is set to stop some inconsitency
            //including highlighting and color coding
            SnippetCode = ActiveSnippet.Code;

            SnippetKind = ActiveSnippet.CodeKindAttribute;
            SnippetLanguage = ActiveSnippet.CodeLanguageAttribute;
            SnippetImports = new CollectionWithEvents<string>(ActiveSnippet.Imports);
            SnippetReferences = new CollectionWithEvents<string>(ActiveSnippet.References);
        }

        /// <summary>
        /// Takes the data from the form and adds it to the in memory xml document
        /// </summary>
        public void PushFieldsIntoActiveSnippet()
        {
            //add header info
            ActiveSnippet.Title = SnippetTitle;
            ActiveSnippet.Author = SnippetAuthor;
            ActiveSnippet.Description = SnippetDescription;
            ActiveSnippet.HelpUrl = SnippetHelpUrl;
            ActiveSnippet.Shortcut = SnippetShortcut;
            ActiveSnippet.AlternativeShortcuts = SnippetAlternativeShortcuts;

            ActiveSnippet.Keywords = SnippetKeywords;
            ActiveSnippet.SnippetTypes = SnippetTypes;
            ActiveSnippet.CodeDelimiterAttribute = SnippetDelimiter;
            ActiveSnippet.Code = SnippetCode;

            //must be after code node is declared
            //kind and language values
            ActiveSnippet.CodeKindAttribute = SnippetKind;
            ActiveSnippet.CodeLanguageAttribute = SnippetLanguage;

            ActiveSnippet.Imports = SnippetImports;
            ActiveSnippet.References = SnippetReferences;
            ActiveSnippet.Literals = SnippetReplacements;
        }

        public void ReplacementRemove()
        {
            if (GetClickedOnReplacementSpan(out var currentWordSpan))
            {
                var currentWord = currentWordSpan.GetText();
                ReplacementRemove(currentWord);
            }
        }

        /// <summary>
        /// save snippet and update the snippets in memory
        /// </summary>
        /// <returns></returns>
        public bool SaveSnippet()
        {
            PushFieldsIntoActiveSnippet();
            if (SnippetDesignerPackage.Instance != null)
            {
                SnippetDesignerPackage.Instance.SnippetIndex.UpdateSnippetFile(snippetFile);
            }
            snippetFile.Save();
            return true;
        }

        /// <summary>
        /// Save snippet as
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public bool SaveSnippetAs(string fileName)
        {
            PushFieldsIntoActiveSnippet();
            foreach (var snippetItem in snippetFile.Snippets)
            {
                if (SnippetDesignerPackage.Instance != null)
                {
                    SnippetDesignerPackage.Instance.SnippetIndex.CreateIndexItemDataFromSnippet(snippetItem, fileName);
                }
            }
            snippetFile.SaveAs(fileName);
            return true;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            foreach (var displayLang in LanguageMaps.LanguageMap.DisplayLanguageToXML.Keys.Where(x => !string.IsNullOrEmpty(x)))
            {
                toolStripLanguageBox.Items.Add(displayLang);
            }
        }

        protected void RefreshReplacementMarkers()
        {
            var allReplacements = GetCurrentReplacements();

            //search through the code window and update all replcement highlight martkers
            MarkReplacements(allReplacements);
        }

        private void AlertInvalidReplacement(string newIdValue) => logger.MessageBox("Invalid Replacement", string.Format(Resources.ErrorInvalidReplacementID, newIdValue), LogType.Warning);

        private bool CreateReplacement(string textToChange)
        {
            if (reservedReplacements.Contains(textToChange.Trim()))
            {
                return false;
            }

            //check if replacement exists already
            var existsAlready = false;
            foreach (DataGridViewRow row in replacementGridView.Rows)
            {
                if ((string)row.Cells[StringConstants.ColumnID].EditedFormattedValue == textToChange ||
                    textToChange.Trim() == string.Empty)
                {
                    //this replacement already exists or is nothing don't add it to the replacement list
                    existsAlready = true;
                }
            }

            //build new replacement text
            var newText = TurnTextIntoReplacementSymbol(textToChange);
            if (!existsAlready)
            {
                object[] newRow = { textToChange, textToChange, textToChange, Resources.ReplacementLiteralName, string.Empty, string.Empty, true };
                try
                {
                    var rowIndex = replacementGridView.Rows.Add(newRow);
                    SetOrDisableTypeField(false, rowIndex);
                }
                catch (InvalidOperationException ex)
                {
                    logger.Log("Possible error when reloading snippet", "SnippetEditorForm", ex);
                }
            }

            //replace all occurances of the textToFind with $textToFind$
            var numFoundAndReplaced = ReplaceAll(textToChange, newText, true);

            return numFoundAndReplaced > 0;
        }

        private bool FindEnclosingReplacementQuoteSpan(SnapshotSpan span, out SnapshotSpan quoteSpan)
        {
            var lineSnapshot = span.Snapshot.GetLineFromPosition(span.Start.Position);

            var left = span.Start.Position - 1;
            var right = span.End.Position;
            var text = span.Snapshot.GetText();
            quoteSpan = new SnapshotSpan();
            while (left >= lineSnapshot.Start.Position && text[left].ToString() != StringConstants.DoubleQuoteString)
            {
                left--;
            }
            while (right < lineSnapshot.End.Position && text[right].ToString() != StringConstants.DoubleQuoteString)
            {
                right++;
            }
            if (right >= lineSnapshot.End.Position || left < lineSnapshot.Start.Position)
            {
                return false; //we didnt find a quoted replacement string
            }

            quoteSpan = new SnapshotSpan(span.Snapshot, new Span(left, right + 1 - left));

            //is this span surrounded by the replcement markers
            if (!IsSpanReplacement(quoteSpan))
            {
                return false;
            }

            //we have a correct quotes string replcement
            return true;
        }

        private bool GetClickedOnReplacementSpan(out SnapshotSpan replacementSpan)
        {
            replacementSpan = CodeWindow.GetWordTextSpanFromCurrentPosition();
            var currentWordSpan = replacementSpan;

            var currentWord = currentWordSpan.GetText();

            if (string.IsNullOrEmpty(currentWord)) //you might have selected more than a word, so use what you selected
            {
                replacementSpan = CodeWindow.Selection;
            }

            //make sure this is infact a replacement
            if (!IsSpanReplacement(currentWordSpan))
            {
                //this span doesnt seem to be a replacement but maybe its a string and the user just
                //clicked in the middle of it so lets intelligently see if thats true
                if (!FindEnclosingReplacementQuoteSpan(currentWordSpan, out replacementSpan))
                {
                    return false;
                }
            }

            //we have found a replacement span
            return true;
        }

        private List<string> GetCurrentReplacements()
        {
            var allReplacements = new List<string>();
            foreach (DataGridViewRow row in replacementGridView.Rows)
            {
                var idValue = ((string)row.Cells[StringConstants.ColumnID].EditedFormattedValue).Trim();
                if (idValue.Length > 0)
                {
                    allReplacements.Add(idValue);
                }
            }
            return allReplacements;
        }

        private IEnumerable<SnapshotSpan> GetReplaceableSpans(ITextSnapshot textSnapshot, string textToFind)
        {
            var isReplacement = IsTextReplacement(textToFind);
            var text = textSnapshot.GetText();
            var start = 0;
            while ((start = text.IndexOf(textToFind, start)) != -1)
            {
                var end = start + textToFind.Length;
                if ((start - 1 < 0 || !CodeWindow.IsWordChar(text[start - 1])) && (end >= text.Length || !CodeWindow.IsWordChar(text[end]))
                    || isReplacement)
                {
                    yield return new SnapshotSpan(textSnapshot, start, textToFind.Length);
                }

                start += textToFind.Length;
            }
        }

        private IEnumerable<string> GetSnippetTitles()
        {
            var snippetTitles = new List<string>();
            foreach (var s in snippetFile.Snippets)
            {
                snippetTitles.Add(s.Title);
            }
            return snippetTitles;
        }

        private bool IsSpanReplacement(SnapshotSpan replaceSpan)
        {
            var spanText = replaceSpan.GetText();
            if (string.IsNullOrEmpty(spanText))
            {
                return false;
            }

            if (IsTextReplacement(spanText))
            {
                return true;
            }

            //see if replacement symbols surround this span
            if (replaceSpan.Span.Start - 1 >= 0 && replaceSpan.Span.End < replaceSpan.Snapshot.Length)
            {
                var beforeFirstChar = replaceSpan.Snapshot[replaceSpan.Span.Start - 1];
                var afterLastChar = replaceSpan.Snapshot[replaceSpan.Span.End];

                if (beforeFirstChar.ToString() == SnippetDelimiter &&
                    afterLastChar.ToString() == SnippetDelimiter
                    )
                {
                    return true;
                }
            }

            return false;
        }

        private bool IsTextReplacement(string text)
        {
            var firstChar = text[0];
            var lastChar = text[text.Length - 1];

            //check the first and last characters of the span to see if they are the replacement symbols
            if (firstChar.ToString() == SnippetDelimiter &&
                lastChar.ToString() == SnippetDelimiter
                )
            {
                return true;
            }

            return false;
        }

        private bool IsValidReplaceableText(string text) => validPotentialReplacementRegex.IsMatch(text);

        private void languageComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            var langCombo = sender as ToolStripComboBox;
            if (langCombo != null)
            {
                var languageText = langCombo.SelectedItem.ToString();
                if (previousLanguageSelected != languageText) //make sure this is actually a change
                {
                    IsFormDirty = true;

                    //if (languageText == Resources.DisplayNameXML)
                    //{
                    //    //The XML Editor defines its own properties window and by removing adornments it will stop it
                    //    // from showing and allow ours to show
                    //    IOleServiceProvider sp = snippetCodeWindow.VsCodeWindow as IOleServiceProvider;
                    //    if (sp != null)
                    //    {
                    //        ServiceProvider site = new ServiceProvider(sp);
                    //        IVsCodeWindowManager cMan = site.GetService(typeof(SVsCodeWindowManager)) as IVsCodeWindowManager;
                    //        if (cMan != null)
                    //        {
                    //            cMan.RemoveAdornments();
                    //        }
                    //    }
                    //}

                    //store the last language
                    previousLanguageSelected = languageText;

                    RefreshPropertiesWindow();
                }
            }
        }

        private void mainObjectsRepaiont_Paint(object sender, PaintEventArgs e)
        {
            SnippetDesignerPackage.Instance.ActiveSnippetLanguage = SnippetLanguage;
            SnippetDesignerPackage.Instance.ActiveSnippetTitle = SnippetTitle;
        }

        private void MarkReplacements(ICollection<string> replaceIDs)
        {
            if (CodeWindow.TextBuffer == null)
            {
                return;
            }

            var findData = new FindData(SnippetRegexPatterns.BuildValidReplacementString(SnippetDelimiter),
                                        CodeWindow.TextBuffer.CurrentSnapshot,
                                        FindOptions.UseRegularExpressions,
                                        null);
            var candidateSpans = textSearchService.FindAll(findData);
            foreach (var candidateSpan in candidateSpans)
            {
                var replacementText = candidateSpan.GetText();
                var textBetween = TurnReplacementSymbolIntoText(replacementText);
                if (!replaceIDs.Contains(textBetween))
                {
                    CreateReplacement(textBetween);
                }
            }
        }

        private void RefreshPropertiesWindow()
        {
            //refresh the properties window
            var sEditor = (this as SnippetEditor);
            if (sEditor != null)
            {
                sEditor.RefreshPropertiesWindow();
            }
        }

        private int ReplaceAll(string currentWord, string newWord, bool skipIfAlreadyReplacement)
        {
            var textview = CodeWindow.TextView;
            var numberReplaced = 0;

            var textEdit = textview.TextBuffer.CreateEdit();
            foreach (var replaceableSpan in GetReplaceableSpans(textview.TextBuffer.CurrentSnapshot, currentWord))
            {
                //replace the span with the given text
                if (ReplaceSpanWithText(textEdit, newWord, replaceableSpan, skipIfAlreadyReplacement))
                {
                    numberReplaced++;
                }
            }
            textEdit.Apply();
            return numberReplaced;
        }

        private void ReplacementRemove(string textToChange)
        {
            DataGridViewRow rowToDelete = null;
            foreach (DataGridViewRow row in replacementGridView.Rows)
            {
                if ((string)row.Cells[StringConstants.ColumnID].Value == textToChange)
                {
                    rowToDelete = row;
                    break;
                }
            }

            if (rowToDelete != null)
            {
                UpdateMarkersAfterDeletedGridViewRow(rowToDelete);
                replacementGridView.Rows.Remove(rowToDelete);
            }
        }

        private bool ReplaceSpanWithText(ITextEdit textEdit, string newWord, SnapshotSpan replaceSpan, bool skipIfAlreadyReplacement)
        {
            try
            {
                if ((skipIfAlreadyReplacement && !IsSpanReplacement(replaceSpan)) || !skipIfAlreadyReplacement)
                {
                    textEdit.Replace(replaceSpan.Span, newWord);
                    return true;
                }
            }
            catch (Exception)
            {
                return false;
            }

            return false;
        }

        private void SetOrDisableTypeField(bool isObject, int rowIndex)
        {
            if (isObject) //if this is an object than enable the type field
            {
                replacementGridView.Rows[rowIndex].Cells[StringConstants.ColumnType].Value = string.Empty;
                replacementGridView.Rows[rowIndex].Cells[StringConstants.ColumnType].ReadOnly = false;
            }
            else //this is not a object so disable type field
            {
                replacementGridView.Rows[rowIndex].Cells[StringConstants.ColumnType].Value = Resources.TypeInvalidForLiteralSymbol;
                replacementGridView.Rows[rowIndex].Cells[StringConstants.ColumnType].ReadOnly = true;
            }
        }

        private void shortcutTextBox_TextChanged(object sender, EventArgs e)
        {
            var newShortcut = shortcutTextBox.Text;

            if (!string.IsNullOrEmpty(newShortcut) && newShortcut != snippetShortcut)
            {
                snippetShortcut = newShortcut;
                IsFormDirty = true;
            }

            RefreshPropertiesWindow();
        }

        private void snippet_CollectionChanged<T>(object sender, CollectionEventArgs<T> e) => IsFormDirty = true;

        #region Properties which interact with fields in the gui

        public CollectionWithEvents<AlternativeShortcut> SnippetAlternativeShortcuts
        {
            get => snippetAlternativeShortcuts;
            set
            {
                snippetAlternativeShortcuts.Clear();
                snippetAlternativeShortcuts.AddRange(value);
            }
        }

        public string SnippetAuthor
        {
            get => snippetAuthor;

            set
            {
                if (snippetAuthor != value)
                {
                    IsFormDirty = true;
                }
                snippetAuthor = value;
            }
        }

        public string SnippetCode
        {
            get => snippetCodeWindow.CodeText;

            set => snippetCodeWindow.CodeText = value;
        }

        public string SnippetDelimiter
        {
            get => snippetDelimiter;

            set
            {
                IsFormDirty = true;
                snippetDelimiter = value;

                CodeWindow.TextView.Properties[SnippetReplacementTagger.ReplacementDelimiter] = snippetDelimiter;

                if (CodeWindow.TextView.Properties.TryGetProperty<SnippetReplacementTagger>(SnippetReplacementTagger.TaggerInstance, out var tagger))
                {
                    tagger.UpdateSnippetReplacementAdornmentsAsync();
                }
            }
        }

        public string SnippetDescription
        {
            get => snippetDescription;

            set
            {
                if (snippetDescription != value)
                {
                    IsFormDirty = true;
                }
                snippetDescription = value;
            }
        }

        /// <summary>
        /// File name of the snippet
        /// </summary>
        public string SnippetFileName => snippetFile.FileName;

        public string SnippetHelpUrl
        {
            get => snippetHelpUrl;

            set
            {
                if (snippetHelpUrl != value)
                {
                    IsFormDirty = true;
                }
                snippetHelpUrl = value;
            }
        }

        public CollectionWithEvents<string> SnippetImports
        {
            get => snippetImports;

            set
            {
                //This wont be called from the editor properties window
                //so we need to figure out a different way to tell if it is in a dirty state
                snippetImports.Clear();
                snippetImports.AddRange(value);
            }
        }

        public CollectionWithEvents<string> SnippetKeywords
        {
            get => snippetKeywords;

            set
            {
                //TODO: check if the keywords have changed
                IsFormDirty = true;

                snippetKeywords.Clear();
                snippetKeywords.AddRange(value);
            }
        }

        public string SnippetKind
        {
            get => snippetKind;

            set
            {
                IsFormDirty = true;
                snippetKind = value;
            }
        }

        /// <summary>
        /// Get: Converts the snippet language from display to xml form and returns it
        /// Set: Take a string snippet language in xml form and ocnverts to display form then adds it to GUI
        /// </summary>
        public string SnippetLanguage
        {
            get
            {
                if (toolStripLanguageBox.SelectedIndex > -1)
                {
                    var langString = toolStripLanguageBox.SelectedItem.ToString();
                    if (LanguageMaps.LanguageMap.DisplayLanguageToXML.ContainsKey(langString))
                    {
                        return LanguageMaps.LanguageMap.DisplayLanguageToXML[langString];
                    }
                    else
                    {
                        return string.Empty;
                    }
                }
                else
                {
                    return string.Empty;
                }
            }

            set
            {
                var language = string.Empty;
                if (!string.IsNullOrEmpty(value) && LanguageMaps.LanguageMap.SnippetSchemaLanguageToDisplay.ContainsKey(value.ToLower()))
                {
                    language = LanguageMaps.LanguageMap.SnippetSchemaLanguageToDisplay[value.ToLower()];
                }
                else
                {
                    language = LanguageMaps.LanguageMap.ToDisplayForm(SnippetDesignerPackage.Instance.Settings.DefaultLanguage);
                }

                var index = toolStripLanguageBox.Items.IndexOf(language);
                if (index >= 0)
                {
                    toolStripLanguageBox.SelectedIndex = index;
                }
                else
                {
                    toolStripLanguageBox.SelectedIndex = 0; //select first
                }
            }
        }

        public CollectionWithEvents<string> SnippetReferences
        {
            get => snippetReferences;

            set
            {
                //This wont be called from the editor properties window
                //so we need to figure out a different way to tell if it is in a dirty state

                snippetReferences.Clear();
                snippetReferences.AddRange(value);
            }
        }

        public CollectionWithEvents<Literal> SnippetReplacements
        {
            get
            {
                var replacements = new CollectionWithEvents<Literal>();
                foreach (DataGridViewRow row in replacementGridView.Rows)
                {
                    if (row.IsNewRow)
                    {
                        continue;
                    }

                    var currId = (string)row.Cells[StringConstants.ColumnID].EditedFormattedValue;
                    currId = currId.Trim();
                    if (string.IsNullOrEmpty(currId))
                    {
                        continue;
                    }

                    var isObj = false;
                    if (((row.Cells[StringConstants.ColumnReplacementKind]).EditedFormattedValue as string).ToLower() ==
                        Resources.ReplacementObjectName.ToLower())
                    {
                        isObj = true;
                    }

                    var isEditable = (bool)(row.Cells[StringConstants.ColumnEditable]).EditedFormattedValue;
                    replacements.Add(new Literal((string)row.Cells[StringConstants.ColumnID].EditedFormattedValue,
                                                 (string)row.Cells[StringConstants.ColumnTooltip].EditedFormattedValue,
                                                 (string)row.Cells[StringConstants.ColumnDefault].EditedFormattedValue,
                                                 (string)row.Cells[StringConstants.ColumnFunction].EditedFormattedValue,
                                                 isObj,
                                                 isEditable,
                                                 (string)row.Cells[StringConstants.ColumnType].EditedFormattedValue
                                         ));
                }
                return replacements;
            }

            set
            {
                //literals and objects
                replacementGridView.Rows.Clear();
                foreach (var literal in value)
                {
                    var objOrLiteral = Resources.ReplacementLiteralName;
                    if (literal.Object)
                    {
                        objOrLiteral = Resources.ReplacementObjectName;
                    }

                    object[] row = { literal.ID, literal.ToolTip, literal.DefaultValue, objOrLiteral, literal.Type, literal.Function, literal.Editable };
                    var rowIndex = replacementGridView.Rows.Add(row);
                    if (!literal.Object)
                    {
                        SetOrDisableTypeField(false, rowIndex);
                    }
                }
            }
        }

        public string SnippetShortcut
        {
            get => shortcutTextBox.Text;

            set => shortcutTextBox.Text = value;
        }

        /// <summary>
        /// The active snippet title
        /// </summary>
        public string SnippetTitle
        {
            get => snippetTitle;

            set => snippetTitle = value;
        }

        /// <summary>
        /// Get the list of snippet titles form the codeWindowHost
        /// Set the list of items in the codeWindowHost
        /// </summary>
        public CollectionWithEvents<string> SnippetTitles
        {
            get
            {
                var titleArray = new string[toolStripSnippetTitles.Items.Count];
                toolStripSnippetTitles.Items.CopyTo(titleArray, 0);
                return new CollectionWithEvents<string>(titleArray);
            }
            set
            {
                toolStripSnippetTitles.Items.Clear();
                foreach (var title in value)
                {
                    toolStripSnippetTitles.Items.Add(title);
                }
                toolStripSnippetTitles.SelectedIndex = toolStripSnippetTitles.Items.IndexOf(ActiveSnippet.Title);
            }
        }

        public CollectionWithEvents<SnippetType> SnippetTypes
        {
            get => snippetTypes;

            set
            {
                //TODO: check if the snippettype have changed
                IsFormDirty = true;

                snippetTypes.Clear();
                snippetTypes.AddRange(value);
            }
        }

        #endregion Properties which interact with fields in the gui

        private void toolStripSnippetsTitles_SelectedIndexChanged(object sender, EventArgs e)
        {
            var snippetsBox = sender as ToolStripComboBox;
            if (snippetsBox != null)
            {
                var newTitle = snippetsBox.SelectedItem as string;
                if (!string.IsNullOrEmpty(newTitle) && newTitle != ActiveSnippet.Title)
                {
                    PushFieldsIntoActiveSnippet();

                    //foreach (Snippet sn in snippetFile.Snippets)
                    for (var i = 0; i < snippetFile.Snippets.Count; i++)
                    {
                        if (snippetFile.Snippets[i].Title.Equals(newTitle, StringComparison.InvariantCulture))
                        {
                            snippetIndex = i;
                            ActiveSnippet = snippetFile.Snippets[i];
                        }
                    }

                    PullFieldsFromActiveSnippet();

                    //clear and show all markers
                    //RefreshReplacementMarkers(false);

                    //not the best way to do this but since I dont know if we want to move the change current snippet to the
                    // porperties window this will have to do for now
                    // I am assuming this object is actually an instance of snippeteditor
                    var theEditor = this as SnippetEditor;
                    if (theEditor != null)
                    {
                        theEditor.RefreshPropertiesWindow();
                    }
                }
            }
        }

        private void toolStripSnippetTitles_TextUpdate(object sender, EventArgs e)
        {
            var snippetsBox = sender as ToolStripComboBox;
            Debug.WriteLine("Text Update " + snippetsBox.Text);

            if (snippetsBox != null)
            {
                var newTitle = snippetsBox.Text;
                if (!string.IsNullOrEmpty(newTitle) && newTitle != snippetTitle)
                {
                    snippetsBox.Items.Remove(snippetTitle);
                    snippetTitle = newTitle;
                    PushFieldsIntoActiveSnippet();
                    IsFormDirty = true;
                }
            }
        }

        #region Replacement Grid Events

        private void removeReplacementToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // If there is one row then its the default row which we cant delete
            if (replacementGridView.Rows.Count > 1)
            {
                if (replacementGridView.SelectedCells.Count > 0)
                {
                    var rowIndex = replacementGridView.SelectedCells[0].RowIndex;

                    // make sure we are not deleting that last row which is the default one
                    if (rowIndex < replacementGridView.Rows.Count - 1)
                    {
                        var rowToDelete = replacementGridView.Rows[rowIndex];
                        //update markers and remove the row
                        UpdateMarkersAfterDeletedGridViewRow(rowToDelete);
                        replacementGridView.Rows.Remove(rowToDelete);
                    }
                }
            }
        }

        private void replacementGridView_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            var grid = sender as DataGridView;
            if (grid != null)
            {
                if (grid.Columns[e.ColumnIndex].Name == StringConstants.ColumnID)
                {
                    previousIDValue = (string)grid.Rows[e.RowIndex].Cells[e.ColumnIndex].EditedFormattedValue;
                    return;
                }
            }
            previousIDValue = null;
        }

        private void replacementGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            var grid = sender as DataGridView;
            if (grid != null)
            {
                IsFormDirty = true;
                if (grid.Columns[e.ColumnIndex].Name == StringConstants.ColumnID)
                {
                    var newIdValue = (string)grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                    if (newIdValue == null)
                    {
                        //if null make it empty
                        newIdValue = string.Empty;
                    }

                    //make sure a change is being made
                    if (previousIDValue != null && newIdValue != previousIDValue)
                    {
                        //check if the change you made is valid if not tell the user and return to previous value
                        if (IsValidReplaceableText(newIdValue))
                        {
                            //build new replacement text
                            var newReplacement = TurnTextIntoReplacementSymbol(newIdValue);
                            var oldReplacement = TurnTextIntoReplacementSymbol(previousIDValue);

                            //replace all occurances of the oldReplacement with newReplacement
                            //set false so it allows us to override existing replacements
                            ReplaceAll(oldReplacement, newReplacement, false);
                            ReplaceAll(newIdValue, newReplacement, true);

                            IsFormDirty = true; //form is now dirty
                        }
                        else
                        {
                            AlertInvalidReplacement(newIdValue);

                            //set id cell back to the old value
                            grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = previousIDValue;
                        }
                    }
                }
                else if (grid.Columns[e.ColumnIndex].Name == StringConstants.ColumnReplacementKind)
                {
                    if ((string)grid.Rows[e.RowIndex].Cells[e.ColumnIndex].EditedFormattedValue == Resources.ReplacementLiteralName)
                    {
                        SetOrDisableTypeField(false, e.RowIndex);
                    }
                    else
                    {
                        SetOrDisableTypeField(true, e.RowIndex);
                    }
                }
            }
        }

        private void replacementGridView_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            var grid = sender as DataGridView;
            if (grid != null)
            {
                currentlySelectedId = grid.Rows[e.RowIndex].Cells[StringConstants.ColumnID].Value as string;
                RefreshReplacementMarkers();
            }
        }

        private void replacementGridView_RowsRemoved(object sender, DataGridViewRowCancelEventArgs e) => UpdateMarkersAfterDeletedGridViewRow(e.Row);

        private void snippetReplacementGrid_MouseDown(object sender, MouseEventArgs e)
        {
            var info = replacementGridView.HitTest(e.X, e.Y);
            if (info.RowIndex >= 0)
            {
                replacementGridView.Rows[info.RowIndex].Selected = true;
            }
        }

        #endregion Replacement Grid Events

        private string TurnReplacementSymbolIntoText(string text)
        {
            if (text.Length > 2
                && text[0].ToString() == SnippetDelimiter
                && text[text.Length - 1].ToString() == SnippetDelimiter)
            {
                return text.Substring(1, text.Length - 2);
            }
            return text;
        }

        private string TurnTextIntoReplacementSymbol(string text) => SnippetDelimiter + text + SnippetDelimiter;

        private void UpdateMarkersAfterDeletedGridViewRow(DataGridViewRow deletedRow)
        {
            if (deletedRow != null)
            {
                var deletedID = deletedRow.Cells[StringConstants.ColumnID].EditedFormattedValue as string;
                if (deletedID != null)
                {
                    //build new replacement text
                    var currentText = TurnTextIntoReplacementSymbol(deletedID);
                    ReplaceAll(currentText, deletedID, false);
                }
            }
        }
    }
}
