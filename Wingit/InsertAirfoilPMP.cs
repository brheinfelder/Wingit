using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;
using WingIt;

namespace Wingit
{
    public class InsertAirfoilPMP
    {
        //Local Objects
        public IPropertyManagerPage2 swPropertyPage = null;
        InsertAirfoilPMPHandler handler = null;
        ISldWorks iSwApp = null;
        SwAddin userAddin = null;
        IPropertyManagerPageTab ppagetab1 = null;
        IPropertyManagerPageTab ppagetab2 = null;

        #region Property Manager Page Controls

        //Groups
        IPropertyManagerPageGroup group1;

        //Controls
        IPropertyManagerPageLabel NACABoxLabel;
        IPropertyManagerPageTextbox NACABox;
        IPropertyManagerPageLabel ChordLabel;
        IPropertyManagerPageNumberbox ChordBox;
        IPropertyManagerPageLabel TwistLabel;
        IPropertyManagerPageNumberbox TwistBox;
        IPropertyManagerPageLabel TwistLocLabel;
        IPropertyManagerPageNumberbox TwistLocBox;
        IPropertyManagerPageCheckbox MirrorCheck;

        //Control IDs
        public const int NACABoxLabelID = 0;
        public const int NACABoxID = 1;
        public const int ChordLabelID = 2;
        public const int ChordBoxID = 3;
        public const int TwistLabelID = 4;
        public const int TwistBoxID = 5;
        public const int TwistLocLabelID = 6;
        public const int TwistLocBoxID = 7;
        public const int MirrorCheckID = 8;

        #endregion

        public InsertAirfoilPMP(SwAddin addin)
        {
            userAddin = addin;
            if (userAddin != null)
            {
                iSwApp = (ISldWorks)userAddin.SwApp;
                CreatePropertyManagerPage();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("SwAddin not set.");
            }
        }


        protected void CreatePropertyManagerPage()
        {
            int errors = -1;
            int options = /*(int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_OkayButton |*/
                (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_CancelButton |
                (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_OkayButton;

            handler = new InsertAirfoilPMPHandler(userAddin, this);
            swPropertyPage = (IPropertyManagerPage2)iSwApp.CreatePropertyManagerPage("Insert Airfoil", options, handler, ref errors);
            if (swPropertyPage != null && errors == (int)swPropertyManagerPageStatus_e.swPropertyManagerPage_Okay)
            {
                try
                {
                    AddControls();
                }
                catch (Exception e)
                {
                    iSwApp.SendMsgToUser2(e.Message, 0, 0);
                }
            }
        }

        //Controls are displayed on the page top to bottom in the order 
        //in which they are added to the object.
        protected void AddControls()
        {
            short controlType = -1;
            short align = -1;
            int options = -1;

            /*
            //Plane Selection
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Selectionbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            Selectionbox = (IPropertyManagerPageSelectionbox)swPropertyPage.AddControl(SelectionboxID, controlType, "Airfoil Plane", align, options, "Airfoil Plane");

            swSelectType_e[] filters = new swSelectType_e[1];
            filters[0] = swSelectType_e.swSelSKETCHES;
            object filterobj = filters;

            Selectionbox.SingleEntityOnly = true;
            Selectionbox.AllowMultipleSelectOfSameEntity = false;
            Selectionbox.Height = 20;
            Selectionbox.SetSelectionFilters(filterobj);
            */

            //NACA Designation Label
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Label;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            NACABoxLabel = (IPropertyManagerPageLabel)swPropertyPage.AddControl(NACABoxLabelID, controlType, "NACA Designation", align, options, "NACA Designation");

            //NACA Designation Input Box
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Textbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            NACABox = (IPropertyManagerPageTextbox)swPropertyPage.AddControl(NACABoxID, controlType, "NACA Designation", align, options, "NACA Designation");

            NACABox.Style = (int)swPropMgrPageTextBoxStyle_e.swPropMgrPageTextBoxStyle_NotifyOnlyWhenFocusLost;

            //Airfoil Chord Label
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Label;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            ChordLabel = (IPropertyManagerPageLabel)swPropertyPage.AddControl(ChordLabelID, controlType, "Airfoil Chord", align, options, "Airfoil Chord");

            //Chord Box
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Numberbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            ChordBox = (IPropertyManagerPageNumberbox)swPropertyPage.AddControl(ChordBoxID, controlType, "Airfoil Chord", align, options, "Airfoil Chord");

            ChordBox.Style = (int)swPropMgrPageNumberBoxStyle_e.swPropMgrPageNumberBoxStyle_NoScrollArrows | (int)swPropMgrPageNumberBoxStyle_e.swPropMgrPageNumberBoxStyle_SuppressNotifyWhileTracking;
            ChordBox.SetRange2((int)swNumberboxUnitType_e.swNumberBox_Length, (double)0, Math.Pow(10, 10), true, (double)10, (double)20, (double)5);
            ChordBox.Value = 1;

            //Twist Label
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Label;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            TwistLabel = (IPropertyManagerPageLabel)swPropertyPage.AddControl(TwistLabelID, controlType, "Airfoil Twist", align, options, "Airfoil Twist");

            //Twist Box
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Numberbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            TwistBox = (IPropertyManagerPageNumberbox)swPropertyPage.AddControl(TwistBoxID, controlType, "Airfoil Twist Angle", align, options, "Airfoil Twist Angle");

            TwistBox.Style = (int)swPropMgrPageNumberBoxStyle_e.swPropMgrPageNumberBoxStyle_NoScrollArrows | (int)swPropMgrPageNumberBoxStyle_e.swPropMgrPageNumberBoxStyle_SuppressNotifyWhileTracking;
            TwistBox.SetRange2((int)swNumberboxUnitType_e.swNumberBox_Angle, (double)-180, (double)180, true, (double)10, (double)20, (double)5);

            //Twist Location Label
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Label;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            TwistLocLabel = (IPropertyManagerPageLabel)swPropertyPage.AddControl(TwistLabelID, controlType, "Airfoil Twist Location", align, options, "Airfoil Twist Location along Camber Line as a percent of the Chord");

            //Twist Location Box
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Numberbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            TwistLocBox = (IPropertyManagerPageNumberbox)swPropertyPage.AddControl(TwistLocBoxID, controlType, "Airfoil Twist Location", align, options, "Airfoil Twist Location along Camber Line as a percent of the Chord");

            TwistLocBox.Style = (int)swPropMgrPageNumberBoxStyle_e.swPropMgrPageNumberBoxStyle_NoScrollArrows | (int)swPropMgrPageNumberBoxStyle_e.swPropMgrPageNumberBoxStyle_SuppressNotifyWhileTracking;
            TwistLocBox.SetRange2((int)swNumberboxUnitType_e.swNumberBox_Percent, (double)0, (double)100, true, (double)10, (double)20, (double)5);

            //Mirror Checkbox
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Checkbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            MirrorCheck = (IPropertyManagerPageCheckbox)swPropertyPage.AddControl(MirrorCheckID, controlType, "Mirror Airfoil", align, options, "Mirror Airfoil");
        }

        public void Show(airfoil Airfoil)
        {
            if (swPropertyPage != null)
            {
                NACABox.Text = Airfoil.NACA.ToString();
                ChordBox.Value = Airfoil.chord;
                TwistBox.Value = Airfoil.twist;
                TwistLocBox.Value = Airfoil.twistloc;
                MirrorCheck.Checked = Airfoil.mirror;
                swPropertyPage.Show();
            }
        }
    }
}
