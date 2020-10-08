using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Visio = Microsoft.Office.Interop.Visio;

namespace SupportTools_Visio.Actions
{
    class Visio_TableOfContents
    {
        const double cTOC_Initial_xLoc = 1;
        const double cTOC_Initial_yLoc = 8.0;
        const double cTOC_Link_Width = 2.5;
        const double cTOC_Link_Height = 0.125;

        const int    cTOC_Link_FontSize = 10;

        const int    cTOC_MaxItemsInColumn = 25;
        const double cTOC_Offset_Row = 0.25;
        const double cTOC_Offset_Column = 2.5;

        const double cTOC_Page_Initial_xLoc = 9.75;
        const double cTOC_PageLink_Initial_yLoc = 8.125;
        const double cTOC_PageLink_Width = 1.0;
        const double cTOC_PageLink_Height = 0.125;

        const int    cTOC_PageLink_FontSize = 6;


        public static void CreateTableOfContents()
        {
            Visio.Page pageTOC = CreateTOCPage();

            foreach (Visio.Page page in Globals.ThisAddIn.Application.ActiveDocument.Pages)
            {
                if (!page.Name.Equals("Table of Contents"))
                {
                    AddTOCLinkToPage(page);
                }
            }

            // Should drive this off a calculation based on page size, # of pages, etc..  Hack it for now.

            double xLoc = cTOC_Initial_xLoc;
            double yLoc = cTOC_Initial_yLoc;

            int columnCount = 0;

            foreach (Visio.Page page in Globals.ThisAddIn.Application.ActiveDocument.Pages)
            {
                if (!page.Name.Equals("Table of Contents"))
                {
                    AddPageLinkToTOCPage(pageTOC, page, xLoc, yLoc);
                    yLoc -= cTOC_Offset_Row;
                    columnCount++;

                    if (columnCount > cTOC_MaxItemsInColumn)
                    {
                        xLoc += cTOC_Offset_Column;
                        yLoc = cTOC_Initial_yLoc;
                        columnCount = 0;
                    }
                }
            }
        }

        private static void AddPageLinkToTOCPage(Visio.Page pageTOC, Visio.Page page, double xLoc, double yLoc)
        {
            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("AddPageLinkToTOCPage");

            Visio.Shape pageLinkShape = pageTOC.DrawRectangle(xLoc, yLoc, xLoc + cTOC_Link_Width, yLoc +cTOC_Link_Height);

            pageLinkShape.TextStyle = "Normal";
            pageLinkShape.LineStyle = "Text Only";
            pageLinkShape.FillStyle = "Text Only";
            pageLinkShape.Characters.Begin = 0;
            pageLinkShape.Characters.End = 0;
            pageLinkShape.Text = page.Name;
            pageLinkShape.Characters.set_CharProps((short)Visio.VisCellIndices.visCharacterSize, cTOC_Link_FontSize);

            Visio.Hyperlink hlink = pageLinkShape.Hyperlinks.Add();
            // hlink.Name = "do we need a name?";
            hlink.SubAddress = page.Name;

            Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);
        }

        private static void AddTOCLinkToPage(Visio.Page page)
        {
            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("AddTOCLinkToPage");

            // Clear out any existing link.

            foreach (Visio.Shape shape in page.Shapes)
            {
                if (shape.Text == "Table of Contents" || shape.Name == "TOCLink")
                {
                    shape.Delete();
                }
            }

            Visio.Shape tocShape = page.DrawRectangle(
                cTOC_Page_Initial_xLoc, cTOC_PageLink_Initial_yLoc, 
                cTOC_Page_Initial_xLoc + cTOC_PageLink_Width, cTOC_PageLink_Initial_yLoc + cTOC_PageLink_Height);

            tocShape.Name = "TOCLink";

            tocShape.Text = "Table of Contents";
            tocShape.TextStyle = "Normal";
            tocShape.LineStyle = "Text Only";
            tocShape.FillStyle = "Text Only";
            tocShape.Characters.Begin = 0;
            tocShape.Characters.End = 0;
            tocShape.Characters.set_CharProps((short)Visio.VisCellIndices.visCharacterSize, 6);

            Visio.Hyperlink hlink = tocShape.Hyperlinks.Add();
            // hlink.Name = "do we need a name?";
            hlink.SubAddress = "Table of Contents";

            Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);
        }

        private static Visio.Page CreateTOCPage()
        {
            Visio.Page page = null;

            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("GenerateTOCPage");

            try
            {
                page = Globals.ThisAddIn.Application.ActiveDocument.Pages["Table of Contents"];
                // We found a page, delete it.  Not much luck iterating across shapes and clearing page - See ClearPage()

                page.Delete(0);
                //ClearPage(page);
                // Need to delete all the stuff.
            }
            catch (Exception ex)
            {
                
                
            }

            page = Globals.ThisAddIn.Application.ActiveDocument.Pages.Add();

            page.Name = "Table of Contents";
            page.Background = 0;
            page.Index = 1;

            Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);
            
            return page;
        }

        private static void ClearPage(Visio.Page page)
        {
            System.Diagnostics.Debug.WriteLine(string.Format("Shapes on Page: {0}", page.Shapes.Count));

            // For some reason this deletes every other shape??  First time 32-16, 16-8, 8-4, etc.
            //foreach (Visio.Shape shape in page.Shapes)
            //{
            //    shape.Delete();
            //}

            try
            {
                for (int i = page.Shapes.Count - 1; i >= 0; i--)
                {
                    page.Shapes[i].Delete();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.ToString());
            }

            System.Diagnostics.Debug.WriteLine(string.Format("Shapes on Page: {0}", page.Shapes.Count));
            
        }

    }
}
