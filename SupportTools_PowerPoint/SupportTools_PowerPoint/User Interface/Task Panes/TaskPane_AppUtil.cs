using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;


namespace SupportTools_PowerPoint.User_Interface.Task_Panes
{
    public partial class TaskPane_AppUtil : UserControl
    {
        public TaskPane_AppUtil()
        {
            InitializeComponent();
        }

        #region Main Function Routines
        
        private void ListShapes()
        {
            foreach (Shape shp in Globals.ThisAddIn.Application.ActivePresentation.Slides[1].Shapes)
            {
                DisplayShapeInfo(shp);
            }
        }

        #endregion

        #region Event Handlers
        
        private void btnListShapes_Click(object sender, EventArgs e)
        {
            ListShapes();
        }

        #endregion

        private void DisplayShapeInfo(Shape shape)
        {
            string textFrame = "";
            string textFrame2 = "";
            string Id = "";
            string name = "";
            string dashStyle = "";
            string BackColor = "";
            string foreColor = "";


            try
            {
                Id = shape.Id.ToString();
            }
            catch(Exception ex)
            {
                Id = "<No Id>";
            }

            try
            {
                name = shape.Name;
            }
            catch(Exception ex)
            {
                name = "<No Name>";
            }

            try
            {
                dashStyle = shape.Line.DashStyle.ToString();
            }
            catch(Exception ex)
            {
                dashStyle = "<No DashStyle>";
            }

            try
            {
                BackColor = shape.Fill.BackColor.RGB.ToString();
            }
            catch(Exception ex)
            {
                BackColor = "<No BackColor>";
            }

            try
            {
                foreColor = shape.Fill.ForeColor.RGB.ToString();
            }
            catch(Exception ex)
            {
                foreColor = "<No ForeColor>";
            }

            try
            {
                if(shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    textFrame = shape.TextFrame.TextRange.Text;
                }
                else
                {
                    textFrame = "";
                }
            }
            catch(Exception ex)
            {
                Common.WriteToDebugWindow("ex shape.TextFrame.HasText");
            }

            try
            {
                if(shape.TextFrame2.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    textFrame2 = shape.TextFrame2.TextRange.Text;
                }
                else
                {
                    textFrame2 = "";
                }
            }
            catch(Exception ex)
            {
                Common.WriteToDebugWindow("ex shape.TextFrame2.HasText");
            }

            Common.WriteToDebugWindow(string.Format("Id:{0,-6}  Name:{1,-20}  ForeColor:{2,-10}  BackColor:{3,-10}   DashStyle:{4,-10}   TextFrame:{5,-20}   TextFrame2:{6,-20}",
                Id, name, foreColor, BackColor, dashStyle, textFrame, textFrame2));
        }

        private void UpdateShape()
        {
            Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            DocumentWindow window = Globals.ThisAddIn.Application.ActiveWindow;
            Slide slide = window.Selection.SlideRange[1];
            Shape shape = window.Selection.ShapeRange[1];
            
            shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(pnlForeColor.ForeColor);

            shape.TextFrame.TextRange.Text = txtTextFrame1.Text;
        }

        private void pnlForeColor_DoubleClick(object sender, EventArgs e)
        {
            colorDialog1.ShowDialog();
            pnlForeColor.ForeColor = colorDialog1.Color;
            pnlForeColor.BackColor = colorDialog1.Color;
        }

        private void btnListShapeInfo_Click(object sender, EventArgs e)
        {
            Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            DocumentWindow window = Globals.ThisAddIn.Application.ActiveWindow;
            Shape shape = window.Selection.ShapeRange[1];

            DisplayShapeInfo(shape);
        }

        private void btnUpdateTextFrame2_Click(object sender, EventArgs e)
        {
            Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            DocumentWindow window = Globals.ThisAddIn.Application.ActiveWindow;
            Slide slide = window.Selection.SlideRange[1];
            Shape shape = window.Selection.ShapeRange[1];

            shape.TextFrame2.TextRange.Text = txtTextFrame2.Text;
        }

        private void btnUpdateForeColor_Click(object sender, EventArgs e)
        {
            Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            DocumentWindow window = Globals.ThisAddIn.Application.ActiveWindow;
            Slide slide = window.Selection.SlideRange[1];
            Shape shape = window.Selection.ShapeRange[1];
            
            shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(pnlForeColor.ForeColor);
        }

        private void btnUpdateBackColor_Click(object sender, EventArgs e)
        {
            Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            DocumentWindow window = Globals.ThisAddIn.Application.ActiveWindow;
            Slide slide = window.Selection.SlideRange[1];
            Shape shape = window.Selection.ShapeRange[1];
            
            shape.Fill.BackColor.RGB = ColorTranslator.ToOle(pnlBackColor.ForeColor);

            shape.TextFrame.TextRange.Text = txtTextFrame1.Text;
        }

        private void btnUpdateTextFrame1_Click(object sender, EventArgs e)
        {
            Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            DocumentWindow window = Globals.ThisAddIn.Application.ActiveWindow;
            Slide slide = window.Selection.SlideRange[1];
            Shape shape = window.Selection.ShapeRange[1];

            shape.TextFrame.TextRange.Text = txtTextFrame1.Text;
        }

        private void pnlBackColor_DoubleClick(object sender, EventArgs e)
        {
            colorDialog1.ShowDialog();
            pnlBackColor.ForeColor = colorDialog1.Color;
            pnlBackColor.BackColor = colorDialog1.Color;
        }

        private void btnUpdateName_Click(object sender, EventArgs e)
        {
            DocumentWindow window = Globals.ThisAddIn.Application.ActiveWindow;
            Slide slide = window.Selection.SlideRange[1];
            Shape shape = window.Selection.ShapeRange[1];

            shape.Name = txtName.Text;
        }

        private void btnUpdateShape_Click(object sender, EventArgs e)
        {
            DocumentWindow window = Globals.ThisAddIn.Application.ActiveWindow;
            Slide slide = window.Selection.SlideRange[1];
            Shape shape = window.Selection.ShapeRange[1];

            shape.Name = txtName.Text;
            shape.TextFrame.TextRange.Text = txtTextFrame1.Text;
            shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(pnlForeColor.ForeColor);
        }

        private void TaskPane_AppUtil_Load(object sender, EventArgs e)
        {
            ucColorList1.PopulateListFromFile(@"C:\Temp\UnitedStatesMapConfig.xml");
            ucStateList1.PopulateListFromFile(@"C:\Temp\UnitedStatesMapConfig.xml");         
        }

        private void btnUpdateColor_Click(object sender, EventArgs e)
        {
            DocumentWindow window = Globals.ThisAddIn.Application.ActiveWindow;
            Slide slide = window.Selection.SlideRange[1];
            //Shape shape = window.Selection.ShapeRange[txtShapeName.Text];
            Shape shape = slide.Shapes[txtShapeName.Text];

            shape.Fill.ForeColor.RGB = int.Parse(txtColorValue.Text);
        }

        private void ucStateList1_ListElementsSelectionChanged_Event()
        {
            txtForeColor.Text = ucStateList1.ForeColor;
            txtShapeName.Text = ucStateList1.ShapeName;
            txtTextLine2.Text = ucStateList1.TextLine2;
        }

        private void ucColorList1_ColorsSelectionChanged_Event()
        {
            txtColorValue.Text = ucColorList1.ColorValue;
        }

        private void btnUpdateStates_Click(object sender, EventArgs e)
        {
            DocumentWindow window = Globals.ThisAddIn.Application.ActiveWindow;
            Slide slide = window.Selection.SlideRange[1];

            foreach (var state in ucStateList1.ListElements)
            {
                try
                {
                    Shape shape = slide.Shapes[state.ShapeName];

                    shape.Fill.ForeColor.RGB = int.Parse(state.ForeColor);
                }
                catch (Exception ex)
                {
                    
                }
            }
        }
    }
}
