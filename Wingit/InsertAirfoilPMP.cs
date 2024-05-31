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

        #region Property Manager Page Controls

        //Groups
        IPropertyManagerPageGroup AirfoilType;
        IPropertyManagerPageGroup OptionsPage;

        //Controls
        public IPropertyManagerPageOption NACAAirfoil;
        public IPropertyManagerPageOption CustomAirfoil;
        public IPropertyManagerPageLabel NACABoxLabel;
        public IPropertyManagerPageTextbox NACABox;
        public IPropertyManagerPageLabel ChordLabel;
        public IPropertyManagerPageNumberbox ChordBox;
        public IPropertyManagerPageLabel TwistLabel;
        public IPropertyManagerPageNumberbox TwistBox;
        public IPropertyManagerPageLabel TwistLocLabel;
        public IPropertyManagerPageNumberbox TwistLocBox;
        public IPropertyManagerPageCheckbox MirrorCheck;
        public IPropertyManagerPageButton ImportAirfoil;
        public IPropertyManagerPageTextbox CustomAirfoilFileName;
        public IPropertyManagerPageCheckbox InvertCamber;

        //Group IDs
        public const int AirfoilTypeGroupID = 20;
        public const int OptionsPageID = 21;

        //Control IDs
        public const int NACAAirfoilID = 0;
        public const int CustomAirfoilID = 1;
        public const int NACABoxLabelID = 2;
        public const int NACABoxID = 3;
        public const int ChordLabelID = 4;
        public const int ChordBoxID = 5;
        public const int TwistLabelID = 6;
        public const int TwistBoxID = 7;
        public const int TwistLocLabelID = 8;
        public const int TwistLocBoxID = 9;
        public const int MirrorCheckID = 10;
        public const int ImportAirfoilID = 11;
        public const int CustomAirfoilFileNameID = 12;
        public const int InvertCamberID = 13;

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
            int options =
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

            //Add Groups
            options = (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Expanded |
                      (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Visible;

            AirfoilType = (IPropertyManagerPageGroup)swPropertyPage.AddGroupBox(AirfoilTypeGroupID, "Airfoil Type", options);
            OptionsPage = (IPropertyManagerPageGroup)swPropertyPage.AddGroupBox(AirfoilTypeGroupID, "Options", options);

            //NACA Airfoil Type Selector
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Option;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            NACAAirfoil = (IPropertyManagerPageOption)AirfoilType.AddControl(NACAAirfoilID, controlType, "NACA Airfoil", align, options, "NACA Airfoil");

            NACAAirfoil.Checked = true;
            NACAAirfoil.Style = (int)swPropMgrPageOptionStyle_e.swPropMgrPageOptionStyle_FirstInGroup;

            //Custom Airfoil Type Selector
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Option;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            CustomAirfoil = (IPropertyManagerPageOption)AirfoilType.AddControl(CustomAirfoilID, controlType, "Custom Airfoil", align, options, "Custom Airfoil");

            //Airfoil Message
            swPropertyPage.SetMessage3("Compatible with basic 4 and 5 digit NACA airfoils.",
                                            (int)swPropertyManagerPageMessageVisibility.swImportantMessageBox,
                                            (int)swPropertyManagerPageMessageExpanded.swMessageBoxExpand,
                                            "NACA Airfoil");

            //NACA Designation Label
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Label;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            NACABoxLabel = (IPropertyManagerPageLabel)AirfoilType.AddControl(NACABoxLabelID, controlType, "NACA Designation", align, options, "NACA Designation");

            //NACA Designation Input Box
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Textbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            NACABox = (IPropertyManagerPageTextbox)AirfoilType.AddControl(NACABoxID, controlType, "NACA Designation", align, options, "NACA Designation");

            NACABox.Style = (int)swPropMgrPageTextBoxStyle_e.swPropMgrPageTextBoxStyle_NotifyOnlyWhenFocusLost;

            //Import Airfoil Button
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Button;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            ImportAirfoil = (IPropertyManagerPageButton)AirfoilType.AddControl(ImportAirfoilID, controlType, "Import Airfoil", align, options, "Import Airfoil Data");

            ((IPropertyManagerPageControl)ImportAirfoil).Visible = false;

            //Custom Airfoil Filename Input Box
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Textbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = 0;

            CustomAirfoilFileName = (IPropertyManagerPageTextbox)AirfoilType.AddControl(CustomAirfoilFileNameID, controlType, "", align, options, "Airfoil File Name");

            CustomAirfoilFileName.Style = (int)swPropMgrPageTextBoxStyle_e.swPropMgrPageTextBoxStyle_ReadOnly;

            //Airfoil Chord Label
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Label;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            ChordLabel = (IPropertyManagerPageLabel)OptionsPage.AddControl(ChordLabelID, controlType, "Airfoil Chord", align, options, "Airfoil Chord");

            //Chord Box
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Numberbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            ChordBox = (IPropertyManagerPageNumberbox)OptionsPage.AddControl(ChordBoxID, controlType, "Airfoil Chord", align, options, "Airfoil Chord");

            ChordBox.Style = (int)swPropMgrPageNumberBoxStyle_e.swPropMgrPageNumberBoxStyle_NoScrollArrows | (int)swPropMgrPageNumberBoxStyle_e.swPropMgrPageNumberBoxStyle_SuppressNotifyWhileTracking;
            ChordBox.SetRange2((int)swNumberboxUnitType_e.swNumberBox_Length, (double)0, Math.Pow(10, 10), true, (double)10, (double)20, (double)5);
            ChordBox.Value = 1;

            //Twist Label
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Label;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            TwistLabel = (IPropertyManagerPageLabel)OptionsPage.AddControl(TwistLabelID, controlType, "Airfoil Twist", align, options, "Airfoil Twist");

            //Twist Box
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Numberbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            TwistBox = (IPropertyManagerPageNumberbox)OptionsPage.AddControl(TwistBoxID, controlType, "Airfoil Twist Angle", align, options, "Airfoil Twist Angle");

            TwistBox.Style = (int)swPropMgrPageNumberBoxStyle_e.swPropMgrPageNumberBoxStyle_NoScrollArrows | (int)swPropMgrPageNumberBoxStyle_e.swPropMgrPageNumberBoxStyle_SuppressNotifyWhileTracking;
            TwistBox.SetRange2((int)swNumberboxUnitType_e.swNumberBox_Angle, (double)-180, (double)180, true, (double)10, (double)20, (double)5);

            //Twist Location Label
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Label;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            TwistLocLabel = (IPropertyManagerPageLabel)OptionsPage.AddControl(TwistLabelID, controlType, "Airfoil Twist Location", align, options, "Airfoil Twist Location along Camber Line as a percent of the Chord");

            //Twist Location Box
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Numberbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            TwistLocBox = (IPropertyManagerPageNumberbox)OptionsPage.AddControl(TwistLocBoxID, controlType, "Airfoil Twist Location", align, options, "Airfoil Twist Location along Camber Line as a percent of the Chord");

            TwistLocBox.Style = (int)swPropMgrPageNumberBoxStyle_e.swPropMgrPageNumberBoxStyle_NoScrollArrows | (int)swPropMgrPageNumberBoxStyle_e.swPropMgrPageNumberBoxStyle_SuppressNotifyWhileTracking;
            TwistLocBox.SetRange2((int)swNumberboxUnitType_e.swNumberBox_Percent, (double)0, (double)100, true, (double)10, (double)20, (double)5);

            //Mirror Checkbox
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Checkbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            MirrorCheck = (IPropertyManagerPageCheckbox)OptionsPage.AddControl(MirrorCheckID, controlType, "Mirror Airfoil", align, options, "Mirror Airfoil");

            //Invert Camber Checkbox
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Checkbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            InvertCamber = (IPropertyManagerPageCheckbox)OptionsPage.AddControl(InvertCamberID, controlType, "Invert Camber", align, options, "Invert Camber");
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
                InvertCamber.Checked = Airfoil.invertcamber;
                CustomAirfoilFileName.Text = Airfoil.airfoilfilename;
                if (Airfoil.airfoiltype==airfoil.AirfoilType.Custom)
                {
                    CustomAirfoil.Checked = true;
                    NACAAirfoil.Checked = false;
                    ((IPropertyManagerPageControl)NACABoxLabel).Visible = false;
                    ((IPropertyManagerPageControl)NACABox).Visible = false;
                    ((IPropertyManagerPageControl)ImportAirfoil).Visible = true;
                    ((IPropertyManagerPageControl)CustomAirfoilFileName).Visible = true;
                }
                else
                {
                    CustomAirfoil.Checked = false;
                    NACAAirfoil.Checked = true;
                    ((IPropertyManagerPageControl)NACABoxLabel).Visible = true;
                    ((IPropertyManagerPageControl)NACABox).Visible = true;
                    ((IPropertyManagerPageControl)ImportAirfoil).Visible = false;
                    ((IPropertyManagerPageControl)CustomAirfoilFileName).Visible = false;
                }
                swPropertyPage.Show();
            }
        }
    }
}
