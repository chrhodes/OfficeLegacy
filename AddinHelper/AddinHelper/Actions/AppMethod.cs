using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace AddinHelper
{
    public class AppMethod
    {
        public const string PROJECT_NAME = "AddinHelper";
        private Office.CommandBar _commandBar;
        private Office.CommandBarButton _commandBarButton;
        private Office.MsoButtonStyle _buttonStyle;
        private Office._CommandBarButtonEvents_ClickEventHandler _eventHandler;
        private string _name;
        private string _bitMapName = "";
        private Assembly _asmbly;
        private string _caption;
        private string _toolTipText;
        private string _description;

        private object missing = System.Reflection.Missing.Value;

        public AppMethod()
        {
        }

        public void Initialize()
        {
            // TODO: Add some validation that all has been set before

            switch (ButtonStyle)
            {
                case Office.MsoButtonStyle.msoButtonCaption:
                    AddButton(Caption, Description, ToolTipText);
                    break;

                case Office.MsoButtonStyle.msoButtonIcon:
                    AddButton(Asmbly, BitMapName, Description, ToolTipText);
                    break;

                case Office.MsoButtonStyle.msoButtonIconAndCaption:
                    AddButton(Asmbly, BitMapName, Caption, Description, ToolTipText);
                    break;

                default:
                    break;

            }
        }

        public void Delete()
        {
            CommandBarButton.Delete(missing);
        }

        public void Disable()
        {
            CommandBarButton.Enabled = false;
        }

        public void Enable()
        {
            CommandBarButton.Enabled = true;
        }

        public Office.CommandBar CommandBar
        {
            get
            {
                return _commandBar;
            }
            set
            {
                _commandBar = value;
            }
        }

        public Office.CommandBarButton CommandBarButton
        {
            get
            {
                return _commandBarButton;
            }
            set
            {
                _commandBarButton = value;
            }
        }

        public Office.MsoButtonStyle ButtonStyle
        {
            get
            {
                return _buttonStyle;
            }
            set
            {
                _buttonStyle = value;
            }
        }

        public Office._CommandBarButtonEvents_ClickEventHandler EventHandler
        {
            get
            {
                return _eventHandler;
            }
            set
            {
                _eventHandler = value;
            }
        }

        public string Name
        {
            get
            {
                return _name;
            }
            set
            {
                _name = value;
            }
        }

        public string BitMapName
        {
            get
            {
                return _bitMapName;
            }
            set
            {
                _bitMapName = value;
            }
        }

        public Assembly Asmbly
        {
            get
            {
                return _asmbly;
            }
            set
            {
                _asmbly = value;
            }
        }

        public string Caption
        {
            get
            {
                return _caption;
            }
            set
            {
                _caption = value;
            }
        }


        public string ToolTipText
        {
            get
            {
                return _toolTipText;
            }
            set
            {
                _toolTipText = value;
            }
        }

        public string Description
        {
            get
            {
                return _description;
            }
            set
            {
                _description = value;
            }
        }

        private void AddButton(
            string caption,
            string description,
            string toolTipText)
        {
            CommandBarButton = CommandBarHelper.AddButton(
                CommandBarButton,
                CommandBar,
                Name,
                caption,
                description,
                toolTipText,
                EventHandler);
        }

        private void AddButton(
            Assembly asmbly,
            string bitMapName,
            string description,
            string toolTipText)
        {
            try
            {
                CommandBarButton = CommandBarHelper.AddButton(
                    CommandBarButton,
                    CommandBar,
                    Name,
                    asmbly,
                    bitMapName,
                    description,
                    toolTipText,
                    EventHandler);
            }
            catch (Exception ex)
            {
                // TODO: Add Logging
                MessageBox.Show(ex.ToString());
                throw ex;
            }
        }

        private void AddButton(
            Assembly asmbly,
            string bitMapName,
            string caption,
            string description,
            string toolTipText)
        {
            try
            {
                CommandBarButton = CommandBarHelper.AddButton(
                    CommandBarButton,
                    CommandBar,
                    Name,
                    asmbly,
                    bitMapName,
                    caption,
                    description,
                    toolTipText,
                    EventHandler);
            }
            catch (Exception ex)
            {
                // TODO: Add Logging
                MessageBox.Show(ex.ToString());
                throw ex;
            }
        }
    }
}
