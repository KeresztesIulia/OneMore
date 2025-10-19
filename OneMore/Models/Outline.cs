//************************************************************************************************
// Copyright © 2021 Steven M Cohn.  All rights reserved.
//************************************************************************************************

namespace River.OneMoreAddIn.Models
{
    using System.Linq;
    using System.Web.UI.WebControls;
    using System.Windows.Forms;
    using System.Xml.Linq;
    using System.Collections.Generic;


    /// <summary>
    /// Represents the Outline element of a OneNote page
    /// </summary>
    /// <remarks>
    /// OneMore page models are typically used in conjunction with the PageNamespace class.
    /// However, there are times when multiple pages and multiple namespaces may be needed
    /// so model classes also provide override constructors that accept explicit namespaces.
    /// </remarks>
    internal class Outline : XElement
    {

        protected class OEChildren : XElement
        {
            XNamespace ns;

            public OEChildren()
                : this(PageNamespace.Value)
            {

            }

            public OEChildren(XNamespace ns)
                : base(ns + "OEChildren")
            {
                this.ns = ns;
            }

            protected Paragraph AddContent(object content, int index = -1, bool isText = true)
            {
                try
                {
                    if (index == -1)
                    {
                        Paragraph newParagraph;
                        if (isText)
                        {
                            newParagraph = new Paragraph((string)content);

                        }
                        else
                        {
                            newParagraph = new Paragraph((XElement)content);
                        }
                        Add(newParagraph);
                        return newParagraph;
                    }
                    else
                    {
                        Paragraph paragraph = (Paragraph)Elements(ns + "OE").ElementAt(index);
                        if (isText)
                        {
                            XElement text = new XElement(ns + "T");
                            text.Value = (string)content; //new XCData(content)?
                            paragraph?.Add(text);
                        }
                        else
                        {
                            paragraph?.Add(content);
                        }
                        return paragraph;

                    }
                }
                catch
                {
                    return null;
                }

            }

            public Paragraph AddContent(string content, int index = -1)
            {
                return AddContent(content, index, true);
            }

            public Paragraph AddContent(XElement content, int index = -1)
            {
                return AddContent(content, index, false);
            }
        }

        private readonly XNamespace ns;
        private OEChildren oeChildren;

        #region constructors

        /// <summary>
        /// Instantiates a new empty outline with the predefined namespace
        /// </summary>
        /// <remarks>
        /// PageNamespace.Value must be set prior to using this constructor
        /// </remarks>
        public Outline()
            : this(PageNamespace.Value)
        {

        }


        /// <summary>
        /// Initializes a new empty outline with the given namespace
        /// </summary>
        /// <param name="ns">A namespace</param>
        public Outline(XNamespace ns)
            : base(ns + "Outline")
        {
            this.ns = ns;
            if (oeChildren == null) oeChildren = new OEChildren(ns);
            Add(oeChildren);
        }

        /// <summary>
        /// Initialize a new outline, adding the given content
        /// </summary>
        /// <param name="content"></param>
        /// <remarks>
        /// PageNamespace.Value must be set prior to using this constructor
        /// </remarks>
        public Outline(XElement content)
            : this(PageNamespace.Value, content)
        {
        }


        /// <summary>
        /// Initialize a new outline, adding the given content
        /// </summary>
        /// <param name="ns">A namespace</param>
        /// <param name="content">Content to add to the outline</param>
        public Outline(XNamespace ns, XElement content)
            : this(ns)
        {
            if (content.Name.LocalName == "Outline")
            {
                if (content.HasElements)
                {
                    Add(content.Elements());
                }
            }
            else
            {
                Add(content);
            }
        }

        #endregion

        #region Positional data
        /// <summary>
        /// Get the width of the outline
        /// </summary>
        /// <returns>An integer approximated the width</returns>
        public int GetWidth()
        {
            var size = Element(ns + "Size");
            if (size != null)
            {
                if (size.GetAttributeValue("width", out decimal width))
                {
                    return (int)width;
                }
            }

            return 0;
        }

        /// <summary>
        /// Get the width of the outline
        /// </summary>
        /// <returns>An integer approximated the width</returns>
        public int GetHeight()
        {
            var size = Element(ns + "Size");
            if (size != null)
            {
                if (size.GetAttributeValue("height", out decimal height))
                {
                    return (int)height;
                }
            }

            return 0;
        }

        /// <summary>
        /// Integer approximates of the width and height of the Outline
        /// </summary>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public void GetSize(out int width, out int height)
        {
            width = GetWidth();
            height = GetHeight();
        }


        /// <summary>
        /// Sets the size of the outline
        /// </summary>
        /// <param name="width">The width</param>
        /// <param name="height">The height, optional</param>
        public void SetSize(int width, int height = 0)
        {
            var size = Element(ns + "Size");
            if (size == null)
            {
                size = new XElement(ns + "Size", new XAttribute("width", $"{width}.0"));

                if (height > 0)
                {
                    size.Add(new XAttribute("height", $"{height}.0"));
                }

                AddFirst(size);
            }
            else
            {
                size.SetAttributeValue("width", $"{width}.0");

                if (height > 0)
                {
                    size.SetAttributeValue("height", $"{height}.0");
                }
            }
        }


        /// <summary>
        /// Get the x coordinate of the Outline's position
        /// </summary>
        /// <returns></returns>
        public int GetPositionX()
        {
            var position = Element(ns + "Position");
            if (position != null)
            {
                if (position.GetAttributeValue("x", out decimal x))
                {
                    return (int)x;
                }
            }
            return 0;
        }

        /// <summary>
        /// Get the y coordinate of the Outline's position
        /// </summary>
        /// <returns></returns>
        public int GetPositionY()
        {
            var position = Element(ns + "Position");
            if (position != null)
            {
                if (position.GetAttributeValue("y", out decimal y))
                {
                    return (int)y;
                }
            }
            return 0;
        }

        /// <summary>
        /// Get the position of the Outline
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        public void GetPosition(out int x, out int y)
        {
            x = GetPositionX();
            y = GetPositionY();
        }

        /// <summary>
        /// Sets the position of the outline
        /// </summary>
        /// <param name="x">The x coordinate of the position</param>
        /// <param name="y">The y coordinate of the position</param>
        public void SetPosition(int x = 0, int y = 0)
        {
            var position = Element(ns + "Position");
            if (position == null)
            {
                position = new XElement(ns + "Position", new XAttribute("x", $"{x}.0"));
                position.Add(new XAttribute("y", $"{y}.0"));
                

                AddFirst(position);
            }
            else
            {
                position.SetAttributeValue("x", $"{x}.0");
                position.SetAttributeValue("y", $"{y}.0");
                
            }
        }


        public void GetAllPositionalData(out int x, out int y, out int width, out int height)
        {
            GetPosition(out x, out y);
            GetSize(out width, out height);
        }
        
        public void SetAllPositionalData(int x, int y, int width, int height)
        {
            SetPosition(x, y);
            SetSize(width, height);
        }
        #region overlap
        public static bool Overlap(Outline outline1,  Outline outline2)
        {
            outline1.GetPosition(out int x, out int y);
            bool firstInSecond = Overlap(outline2, x, y);

            outline2.GetPosition(out x, out y);
            bool secondInFirst = Overlap(outline1, x, y);

            return firstInSecond || secondInFirst;
        }

        public static bool Overlap(Outline outline, int x, int y)
        {
            outline.GetPosition(out int compX, out int compY);
            outline.GetSize(out int width, out int height);

            return (compX <= x && x <= compX + width) && (compY <= y && y <= compY + height);
        }

        public bool Overlap(Outline outline)
        {
            return Overlap(this, outline);
        }

        public bool Overlap(int x, int y)
        {
            return Overlap(this, x, y);
        }
        #endregion

        #endregion



        /// <summary>
        /// Add text content to the end of the outline in a new paragraph without having to manually create the structure (OEChildren > OE > T)
        /// </summary>
        /// <param name="content"></param>
        /// <returns>The created paragraph (OE) that contains the added content</returns>
        public Paragraph AddContent(string content)
        {
            return oeChildren.AddContent(content);
        }

        /// <summary>
        /// Add XElement to the end of the outline without having to manually create the structure (OEChildren > OE)
        /// </summary>
        /// <param name="content"></param>
        /// <returns>The created paragraph (OE) that contains the added content</returns>
        public Paragraph AddContent(XElement content)
        {
            return oeChildren.AddContent(content);
        }

        /// <summary>
        /// Get all textnodes in the Outline that are selected.
        /// </summary>
        /// <returns></returns>
        public XElement[] GetSelectedTextNodes()
        {
            var textNodes = Descendants(ns + "T").Where(e => e.Attributes().Any(a => a.Name == "selected" && a.Value != "none"));
            return textNodes.ToArray();
        }

        /// <summary>
        /// Get selected text in the Outline as one string
        /// </summary>
        /// <returns></returns>
        public string GetSelectedText()
        {
            IEnumerable<string> texts = GetSelectedTextNodes().Select(e => e.Value);
            return string.Join("\n", texts);
        }

        /// <summary>
        /// Get the outline selected on the page
        /// </summary>
        /// <param name="page"></param>
        /// <returns></returns>
        public static Outline GetCurrentOutline()
        {
            try
            {
                using var one = new OneNote(out var page, out var ns);

                XElement outline = page.Root.Descendants(ns + "Outline").Where(o => o.Attributes().Any(a => a.Name == "selected" && a.Value != "none")).First();
                return new Outline(outline);

            }
            catch
            {
                return null;
            }

        }
    }
}
