using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SupportTools_Visio
{
    class Common : AddinHelper.Common
    {
        public const int EMPTY_INT = 0;
        public const string EMPTY_STR = "''";

        public const int TABLE_ABSTRACTION = 1;
        public const int TABLE_ATTRIBUTE = 2;
        public const int TABLE_SHAPETYPE = 3;
        public const int TABLE_SHAPE = 4;
        public const int TABLE_SHAPETYPEATTRIBUTE = 5;
        public const int TABLE_ATTRIBUTEVALUE = 6;
        public const int TABLE_RELATION = 7;
        public const int TABLE_SDTABLE = 8;
        public const int TABLE_LOG = 9;
        public const int TABLE_CONSTRAINEDVALUES = 10;
        public const int TABLE_ARTIFACT = 11;
        public const int TABLE_CONTAINER = 12;
        public const int TABLE_SDLOGFUNCTION = 13;
        public const int TABLE_LIFECYCLE = 14;

        public const int SHAPETYPE_APPLICATION = 1;
        public const int SHAPETYPE_DATA = 2;
        public const int SHAPETYPE_BATCHFILE = 3;
        public const int SHAPETYPE_ACTOR = 4;
        public const int SHAPETYPE_FILEREPORT = 5;
        public const int SHAPETYPE_QUEUE = 6;
        public const int SHAPETYPE_JOB = 7;
        public const int SHAPETYPE_OBJECTRELATION = 8;
        public const int SHAPETYPE_SYSTEM = 9;
        public const int SHAPETYPE_SYSTEMRELATION = 10;
        public const int SHAPETYPE_SERVER = 11;
        public const int SHAPETYPE_IP = 12;
        public const int SHAPETYPE_DEVICERELATION = 13;
        public const int SHAPETYPE_DESKTOP = 14;
        public const int SHAPETYPE_MAINFRAME = 15;
        public const int SHAPETYPE_LAPTOP = 16;
        public const int SHAPETYPE_PDA = 17;
        public const int SHAPETYPE_TABLET = 18;
        public const int SHAPETYPE_PRINTER = 19;
        public const int SHAPETYPE_SAN = 20;
        public const int SHAPETYPE_FIREWALL = 21;
        public const int SHAPETYPE_LOADBALANCER = 22;
        public const int SHAPETYPE_ROUTER = 23;
        public const int SHAPETYPE_PIECE = 24;
        public const int SHAPETYPE_BUSINESS = 25;
        public const int SHAPETYPE_STEP = 26;
        public const int SHAPETYPE_DECISION = 27;
        public const int SHAPETYPE_BUSINESSRELATION = 28;
        public const int SHAPETYPE_STEPRELATION = 29;
        public const int SHAPETYPE_CROSSRELATION = 30;
        public const int SHAPETYPE_TIME = 31;
        public const int SHAPETYPE_START = 32;
        public const int SHAPETYPE_END = 33;
        public const int SHAPETYPE_SUBSTEP = 34;
        public const int SHAPETYPE_BROWSER = 35;
        public const int SHAPETYPE_SQLINSTANCE = 36;
        public const int SHAPETYPE_FOLDER = 37;
        public const int SHAPETYPE_FILE = 38;
        public const int SHAPETYPE_REPORT = 39;

        public const int ABSTRACTION_SYSTEM = 1;
        public const int ABSTRACTION_OBJECT = 2;
        public const int ABSTRACTION_DEVICE = 3;
        public const int ABSTRACTION_SYSTEMOBJECT = 4;
        public const int ABSTRACTION_IP = 5;
        public const int ABSTRACTION_DEVICEOBJECT = 6;
        public const int ABSTRACTION_ARTIFACT = 7;
        public const int ABSTRACTION_PIECE = 8;
        public const int ABSTRACTION_OBJECTPIECE = 9;
        public const int ABSTRACTION_BUSINESS = 10;
        public const int ABSTRACTION_STEP = 11;
        public const int ABSTRACTION_BUSINESSSTEP = 12;
        public const int ABSTRACTION_RELATION = 13;
        public const int ABSTRACTION_STEPOBJECT = 14;
        public const int ABSTRACTION_SUBSTEP = 15;

        public const int VISIO_ENABLED = 0;
        public const int VISIO_DISABLED = 1;

        public const int ACTIONITEM_DRILLDOWN = 0;
        public const int ACTIONITEM_DRILLUP = 1;
        public const int ACTIONITEM_EDITVISIO = 2;
        public const int ACTIONITEM_PROPERTIES = 4;
        public const int ACTIONITEM_BUSINESSPROCESS = 5;
        public const int ACTIONITEM_INFRASTRUCTURE = 6;
        public const int ACTIONITEM_SYSTEM = 7;
        public const int ACTIONITEM_RETRIEVE = 8;
        public const int ACTIONITEM_VALIDATESHAPE = 9;
        public const int ACTIONITEM_SHAPEDROP = 10;
        public const int ACTIONITEM_RECONNECT = 11;
        public const int ACTIONITEM_SUMMARIZE = 12;
        public const int ACTIONITEM_VALIDATESHEET = 13;
        public const int ACTIONITEM_SHOWRELATIONS = 14;

        public const int RELATIONTYPE_CONTAINER = 1;
        public const int RELATIONTYPE_CONTAINED = 2;
        public const int RELATIONTYPE_DIRECT = 3;

        public const int LINEFORMAT_LINEPATTERN_SOLID = 1;
        public const int LINEFORMAT_LINEPATTERN_SLASH = 2;
        public const int LINEFORMAT_LINEPATTERN_DOT = 3;
        public const string LINEFORMAT_LINEWEIGHT = "1.5 pt";
        public const int LINEFORMAT_ARROW_NONE = 0;
        public const int LINEFORMAT_ARROW_NORMAL = 1;

        new public const string PROJECT_NAME = "SupportTools_Visio";

        public static AddinHelper.Visio VisioHelper = new AddinHelper.Visio();
        public static Events.VisioAppEvents AppEvents;
        public static bool DisplayChattyEvents = false;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneConfig;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneHelp;

        public static Microsoft.Office.Tools.CustomTaskPane TaskPaneAppUtil;
    }
}
