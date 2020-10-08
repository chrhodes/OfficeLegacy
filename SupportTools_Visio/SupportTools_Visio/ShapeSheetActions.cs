using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Visio = Microsoft.Office.Interop.Visio;

namespace SupportTools_Visio
{
    public class ShapeSheetActions
    {
        public static void DrillDown(Visio.Application app, string doc, string page, string shape, string shapeu)
        {
            System.Windows.Forms.MessageBox.Show("TODO: Navigate Down");
        }

        public static void DrillUp(Visio.Application app, string doc, string page, string shape, string shapeu)
        {
            System.Windows.Forms.MessageBox.Show("TODO: Navigate Up");
        }

        public static void EditVisio(Visio.Application app, string doc, string page, string shape, string shapeu)
        {
            var frm = new User_Interface.Forms.frmEditVisio();
            frm.Show();
        }

        public static void Properties(Visio.Application app, string doc, string page, string shape, string shapeu)
        {
            System.Windows.Forms.MessageBox.Show("TODO: Navigate to Web Page");            
        }

        public static void RelatedProcess(Visio.Application app, string doc, string page, string shape, string shapeu)
        {
            var frm = new User_Interface.Forms.frmRelateShape();
            frm.Show();
        }

        public static void RelatedSystem(Visio.Application app, string doc, string page, string shape, string shapeu)
        {
            var frm = new User_Interface.Forms.frmRelateShape();
            frm.Show();
        }

        public static void RelatedInfrastructure(Visio.Application app, string doc, string page, string shape, string shapeu)
        {
            var frm = new User_Interface.Forms.frmRelateShape();
            frm.Show();
        }

        public static void Retrieve(Visio.Application app, string doc, string page, string shape, string shapeu)
        {
            var frm = new User_Interface.Forms.frmRetrieveShape();
            frm.Show();
        }

        public static void Validate(Visio.Application app, string doc, string page, string shape, string shapeu)
        {
            var frm = new User_Interface.Forms.frmValidate();
            frm.Show();
        }

    }
}
