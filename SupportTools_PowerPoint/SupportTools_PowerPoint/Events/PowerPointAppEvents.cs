using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SupportTools_PowerPoint.Events
{
    public class PowerPointAppEvents
    {
        private Microsoft.Office.Interop.PowerPoint.Application _PowerPointApplication;
        public Microsoft.Office.Interop.PowerPoint.Application PowerPointApplication
        {
            get
            {
                return _PowerPointApplication;
            }
            set
            {
                if (_PowerPointApplication != null)
                {
                    // Should remove all the event handlers;
                }

                _PowerPointApplication = value;

                if (_PowerPointApplication != null)
                {
                    _PowerPointApplication.AfterNewPresentation += _PowerPointApplication_AfterNewPresentation;
                    _PowerPointApplication.AfterPresentationOpen += _PowerPointApplication_AfterPresentationOpen;
                    _PowerPointApplication.ColorSchemeChanged += _PowerPointApplication_ColorSchemeChanged;
                    _PowerPointApplication.PresentationBeforeClose += _PowerPointApplication_PresentationBeforeClose;
                    _PowerPointApplication.PresentationBeforeSave += _PowerPointApplication_PresentationBeforeSave;
                    _PowerPointApplication.PresentationClose += _PowerPointApplication_PresentationClose;
                    _PowerPointApplication.PresentationCloseFinal += _PowerPointApplication_PresentationCloseFinal;
                    _PowerPointApplication.PresentationNewSlide += _PowerPointApplication_PresentationNewSlide;
                    _PowerPointApplication.PresentationOpen += _PowerPointApplication_PresentationOpen;
                    _PowerPointApplication.PresentationPrint += _PowerPointApplication_PresentationPrint;
                    _PowerPointApplication.PresentationSave += _PowerPointApplication_PresentationSave;
                    _PowerPointApplication.PresentationSync += _PowerPointApplication_PresentationSync;
                    _PowerPointApplication.ProtectedViewWindowActivate += _PowerPointApplication_ProtectedViewWindowActivate;
                    _PowerPointApplication.ProtectedViewWindowBeforeClose += _PowerPointApplication_ProtectedViewWindowBeforeClose;
                    _PowerPointApplication.ProtectedViewWindowBeforeEdit += _PowerPointApplication_ProtectedViewWindowBeforeEdit;
                    _PowerPointApplication.ProtectedViewWindowDeactivate += _PowerPointApplication_ProtectedViewWindowDeactivate;
                    _PowerPointApplication.ProtectedViewWindowOpen += _PowerPointApplication_ProtectedViewWindowOpen;
                    _PowerPointApplication.SlideSelectionChanged += _PowerPointApplication_SlideSelectionChanged;
                    _PowerPointApplication.SlideShowBegin += _PowerPointApplication_SlideShowBegin;
                    _PowerPointApplication.SlideShowEnd += _PowerPointApplication_SlideShowEnd;
                    _PowerPointApplication.SlideShowNextBuild += _PowerPointApplication_SlideShowNextBuild;
                    _PowerPointApplication.SlideShowNextClick += _PowerPointApplication_SlideShowNextClick;
                    _PowerPointApplication.SlideShowNextSlide += _PowerPointApplication_SlideShowNextSlide;
                    _PowerPointApplication.SlideShowOnNext += _PowerPointApplication_SlideShowOnNext;
                    _PowerPointApplication.SlideShowOnPrevious += _PowerPointApplication_SlideShowOnPrevious;
                    _PowerPointApplication.WindowActivate += _PowerPointApplication_WindowActivate;
                    //_PowerPointApplication.WindowBeforeRightClick += new Microsoft.Office.Interop.PowerPoint.EApplication_WindowBeforeRightClickEventHandler(_PowerPointApplication_WindowBeforeRightClick);
                    _PowerPointApplication.WindowDeactivate += _PowerPointApplication_WindowDeactivate;
                    _PowerPointApplication.WindowSelectionChange += _PowerPointApplication_WindowSelectionChange;
                }
            }
        }

        short WindowSelectionChange;
        void _PowerPointApplication_WindowSelectionChange(Microsoft.Office.Interop.PowerPoint.Selection Sel)
        {
            DisplayInWatchWindow(WindowSelectionChange++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WindowDeactivate;
        void _PowerPointApplication_WindowDeactivate(Microsoft.Office.Interop.PowerPoint.Presentation Pres, Microsoft.Office.Interop.PowerPoint.DocumentWindow Wn)
        {
            DisplayInWatchWindow(WindowDeactivate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WindowBeforeRightClick;
        void _PowerPointApplication_WindowBeforeRightClick(Microsoft.Office.Interop.PowerPoint.Presentation Pres, Microsoft.Office.Interop.PowerPoint.DocumentWindow Wn)
        {
            DisplayInWatchWindow(WindowBeforeRightClick++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WindowActivate;
        void _PowerPointApplication_WindowActivate(Microsoft.Office.Interop.PowerPoint.Presentation Pres, Microsoft.Office.Interop.PowerPoint.DocumentWindow Wn)
        {
            DisplayInWatchWindow(WindowActivate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SlideShowOnPrevious;
        void _PowerPointApplication_SlideShowOnPrevious(Microsoft.Office.Interop.PowerPoint.SlideShowWindow Wn)
        {
            DisplayInWatchWindow(SlideShowOnPrevious++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SlideShowOnNext;
        void _PowerPointApplication_SlideShowOnNext(Microsoft.Office.Interop.PowerPoint.SlideShowWindow Wn)
        {
            DisplayInWatchWindow(SlideShowOnNext++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SlideShowNextSlide;
        void _PowerPointApplication_SlideShowNextSlide(Microsoft.Office.Interop.PowerPoint.SlideShowWindow Wn)
        {
            DisplayInWatchWindow(SlideShowNextSlide++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SlideShowNextClick;
        void _PowerPointApplication_SlideShowNextClick(Microsoft.Office.Interop.PowerPoint.SlideShowWindow Wn, Microsoft.Office.Interop.PowerPoint.Effect nEffect)
        {
            DisplayInWatchWindow(SlideShowNextClick++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SlideShowNextBuild;
        void _PowerPointApplication_SlideShowNextBuild(Microsoft.Office.Interop.PowerPoint.SlideShowWindow Wn)
        {
            DisplayInWatchWindow(SlideShowNextBuild++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SlideShowEnd;
        void _PowerPointApplication_SlideShowEnd(Microsoft.Office.Interop.PowerPoint.Presentation Pres)
        {
            DisplayInWatchWindow(SlideShowEnd++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SlideShowBegin;
        void _PowerPointApplication_SlideShowBegin(Microsoft.Office.Interop.PowerPoint.SlideShowWindow Wn)
        {
            DisplayInWatchWindow(SlideShowBegin++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SlideSelectionChanged;
        void _PowerPointApplication_SlideSelectionChanged(Microsoft.Office.Interop.PowerPoint.SlideRange SldRange)
        {
            DisplayInWatchWindow(SlideSelectionChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProtectedViewWindowOpen;
        void _PowerPointApplication_ProtectedViewWindowOpen(Microsoft.Office.Interop.PowerPoint.ProtectedViewWindow ProtViewWindow)
        {
            DisplayInWatchWindow(ProtectedViewWindowOpen++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProtectedViewWindowDeactivate;
        void _PowerPointApplication_ProtectedViewWindowDeactivate(Microsoft.Office.Interop.PowerPoint.ProtectedViewWindow ProtViewWindow)
        {
            DisplayInWatchWindow(ProtectedViewWindowDeactivate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProtectedViewWindowBeforeEdit;
        void _PowerPointApplication_ProtectedViewWindowBeforeEdit(Microsoft.Office.Interop.PowerPoint.ProtectedViewWindow ProtViewWindow, ref bool Cancel)
        {
            DisplayInWatchWindow(ProtectedViewWindowBeforeEdit++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProtectedViewWindowBeforeClose;
        void _PowerPointApplication_ProtectedViewWindowBeforeClose(Microsoft.Office.Interop.PowerPoint.ProtectedViewWindow ProtViewWindow, Microsoft.Office.Interop.PowerPoint.PpProtectedViewCloseReason ProtectedViewCloseReason, ref bool Cancel)
        {
            DisplayInWatchWindow(ProtectedViewWindowBeforeClose++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProtectedViewWindowActivate;
        void _PowerPointApplication_ProtectedViewWindowActivate(Microsoft.Office.Interop.PowerPoint.ProtectedViewWindow ProtViewWindow)
        {
            DisplayInWatchWindow(ProtectedViewWindowActivate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short PresentationSync;
        void _PowerPointApplication_PresentationSync(Microsoft.Office.Interop.PowerPoint.Presentation Pres, Microsoft.Office.Core.MsoSyncEventType SyncEventType)
        {
            DisplayInWatchWindow(PresentationSync++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short PresentationSave;
        void _PowerPointApplication_PresentationSave(Microsoft.Office.Interop.PowerPoint.Presentation Pres)
        {
            DisplayInWatchWindow(PresentationSave++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short PresentationPrint;
        void _PowerPointApplication_PresentationPrint(Microsoft.Office.Interop.PowerPoint.Presentation Pres)
        {
            DisplayInWatchWindow(PresentationPrint++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short PresentationOpen;
        void _PowerPointApplication_PresentationOpen(Microsoft.Office.Interop.PowerPoint.Presentation Pres)
        {
            DisplayInWatchWindow(PresentationOpen++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short PresentationNewSlide;
        void _PowerPointApplication_PresentationNewSlide(Microsoft.Office.Interop.PowerPoint.Slide Sld)
        {
            DisplayInWatchWindow(PresentationNewSlide++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short PresentationCloseFinal;
        void _PowerPointApplication_PresentationCloseFinal(Microsoft.Office.Interop.PowerPoint.Presentation Pres)
        {
            DisplayInWatchWindow(PresentationCloseFinal++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short PresentationClose;
        void _PowerPointApplication_PresentationClose(Microsoft.Office.Interop.PowerPoint.Presentation Pres)
        {
            DisplayInWatchWindow(PresentationClose++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short PresentationBeforeSave;
        void _PowerPointApplication_PresentationBeforeSave(Microsoft.Office.Interop.PowerPoint.Presentation Pres, ref bool Cancel)
        {
            DisplayInWatchWindow(PresentationBeforeSave++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short PresentationBeforeClose;
        void _PowerPointApplication_PresentationBeforeClose(Microsoft.Office.Interop.PowerPoint.Presentation Pres, ref bool Cancel)
        {
            DisplayInWatchWindow(PresentationBeforeClose++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ColorSchemeChanged;
        void _PowerPointApplication_ColorSchemeChanged(Microsoft.Office.Interop.PowerPoint.SlideRange SldRange)
        {
            DisplayInWatchWindow(ColorSchemeChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short AfterPresentationOpen;
        void _PowerPointApplication_AfterPresentationOpen(Microsoft.Office.Interop.PowerPoint.Presentation Pres)
        {
            DisplayInWatchWindow(AfterPresentationOpen++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short AfterNewPresentation;
        void _PowerPointApplication_AfterNewPresentation(Microsoft.Office.Interop.PowerPoint.Presentation Pres)
        {
            DisplayInWatchWindow(AfterNewPresentation++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        //short AfterCalculate;
        //void _ExcelApplication_AfterCalculate()
        //{
        //    DisplayInWatchWindow(AfterCalculate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        //}

        private void DisplayInWatchWindow(short i, string outputLine)
        {
            if (Common.DisplayEvents)
            {
                AddinHelper.Common.WriteToWatchWindow(string.Format("{0}:{1}", outputLine, i));
            }
        }
    }
}
