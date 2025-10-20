
namespace River.OneMoreAddIn.Commands
{
    using River.OneMoreAddIn.Models;
    using System.Collections.Generic;
    using System.Drawing;
    using System.Linq;
    using System.Threading.Tasks;
    using System.Xml.Linq;
    using Resx = Properties.Resources;

    internal class SWHAddNoteCommand : Command
    {

        public enum NoteType
        {
            Sidenote,
            EditingNote,
            OverallEditingNote,
            AdditionNote,
            InLineAdditionNote
        }
        protected class NoteTypeInfo
        {
            public NoteType noteType;
            public string name;
            public Color color;
            public bool bold;
            public bool italic;
            public bool underline;
            

            public NoteTypeInfo()
            {
                noteType = NoteType.Sidenote;
                name = "Sidenote";
                color = Color.Empty;
            }

            public NoteTypeInfo(NoteType noteType, string name, Color color, bool bold = false, bool italic = false, bool underline = false)
            {
                this.noteType = noteType;
                this.name = name;
                this.color = color;
                this.bold = bold;
                this.italic = italic;
                this.underline = underline;
            }
        }
        Dictionary<NoteType, NoteTypeInfo> noteTypeInfo;

        XNamespace ns;
        Page page;

        const int horizontalOffset = 30;

        public SWHAddNoteCommand()
        {
            BuildNoteTypeDictionary();
        }
        
        void BuildNoteTypeDictionary()
        {
            noteTypeInfo = new Dictionary<NoteType, NoteTypeInfo>();
            noteTypeInfo[NoteType.Sidenote] = new NoteTypeInfo();
            noteTypeInfo[NoteType.EditingNote] = new NoteTypeInfo(NoteType.EditingNote, "Editing note", Color.Red);
            noteTypeInfo[NoteType.OverallEditingNote] = new NoteTypeInfo(NoteType.OverallEditingNote, "OVERALL EDITING NOTE", Color.Red);
            noteTypeInfo[NoteType.AdditionNote] = new NoteTypeInfo(NoteType.AdditionNote, "Add note", Color.FromArgb(61, 174, 255), italic: true);
            noteTypeInfo[NoteType.InLineAdditionNote] = new NoteTypeInfo(NoteType.InLineAdditionNote, "In-line add note", Color.FromArgb(61, 174, 255));
        }

        public override async Task Execute(params object[] args)
        {
            try
            {
                await using var one = new OneNote(out page, out ns);

                if (!page.ConfirmBodyContext())
                {
                    ShowError(Resx.Error_BodyContext);
                    return;
                }

                var noteType = (NoteType)args[0];
                PageNamespace.Set(ns);

                if (noteType == NoteType.InLineAdditionNote)
                {
                    logger.WriteLineSWH("Inline NoteType not implented yet.");
                    return;
                }
                else
                {
                    bool overall = (noteType == NoteType.OverallEditingNote);
                    /// TODO: ask for quoting 
                    // if SelectionRange == cursorSize (look that one up), THEN ask if quote, otherwise no
                    Outline noteBox = CreateNoteBox(noteTypeInfo[noteType], overall/*, quote: DialogBox("Quote selection?")*/);
                    page.Root.Add(noteBox);
                }

                await one.Update(page);
                logger.WriteLineSWH($"Added {noteTypeInfo[noteType].name} {(false ? "with" : "without")} quoting.");

            }
            catch
            {
                logger.WriteLineSWH("Failed to execute: " + nameof(SWHAddNoteCommand));
            }
        }

        protected Outline CreateNoteBox(NoteTypeInfo noteType, bool overall = false, bool quote = true)
        {
            /// TODO: New noteBoxes shift the existing outlines downwards -- fix that
            Outline noteBox = SetupNoteBox(overall);

            Paragraph noteParagraph = noteBox.AddContent(noteType.name + ": ");


            if (quote)
            {
                noteParagraph.Add(GetQuote());
            }

            XElement text = new(ns + "T", "⋯");
            noteParagraph.Add(text);
            var editor = new PageEditor(page);
            editor.Deselect();
            text.SetAttributeValue("selected", "all");

            /// TODO: set paragraph color
            // Set paragraph style (color, bold, italic, blabla)
            // or all texts if I need, but do that through Paragraph functions, not here
            // also add that to Outline, like, JESUS

            return noteBox;
        }

        protected Outline SetupNoteBox(bool overall = false)
        {
            Outline.GetCurrentOutline().GetAllPositionalData(out int refX, out int refY, out int refWidth, out int refHeight);

            int x = refX + refWidth + horizontalOffset * 2;
            int y = refY;

            if (!overall)
            {
                /// TODO: Find position of selected text if it's not a general note
                y = refY; // obviously change when I figure out how to determine cursor position
            }

            Outline noteBox = null;

            // Add new note to an existing note outline, if one exists at the position we want to create at
            foreach (var outline in page.Root.Elements(ns + "Outline").Select(o => new Outline(o)))
            {
                if (outline.Overlap(x, y))
                {
                    noteBox = outline;
                    break;
                }
            }

            // If the new note's position is empty, create new outline for it
            if (noteBox == null)
            {
                noteBox = new Outline();
                x = refX + refWidth + horizontalOffset;
                noteBox.SetPosition(x, y);
            }

            return noteBox;
        }

        protected string GetQuoteEditor()
        {
            var editor = new PageEditor(page);
            return "\"" + editor.GetSelectedText() + "\" -> ";
        }


        // not using pageEditor for this because it doesn't keep new paragraphs
        protected string GetQuote()
        {
            Outline currentOutline = Outline.GetCurrentOutline();
            if (currentOutline is null)
            {
                // shouldn't happen
                logger.WriteLineSWH("No outline found to quote from.");
                return "";
            }
            string selectedText = currentOutline?.GetSelectedText();
            if (selectedText.IsNullOrEmpty())
            {
                ShowError("Can't quote without a selection!");
                return "";
            }
            return "\"" + selectedText + "\" -> ";
        }
    }
}
