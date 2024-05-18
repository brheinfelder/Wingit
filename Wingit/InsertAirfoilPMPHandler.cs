using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;
using WingIt;

namespace Wingit
{
    public class InsertAirfoilPMPHandler : IPropertyManagerPage2Handler9
    {
        ISldWorks iSwApp;
        SwAddin userAddin;
        InsertAirfoilPMP ppage;

        public const int NACABoxLabelID = 0;
        public const int NACABoxID = 1;
        public const int ChordLabelID = 2;
        public const int ChordBoxID = 3;
        public const int TwistLabelID = 4;
        public const int TwistBoxID = 5;
        public const int TwistLocLabelID = 6;
        public const int TwistLocBoxID = 7;
        public const int MirrorCheckID = 8;

        public InsertAirfoilPMPHandler(SwAddin addin, InsertAirfoilPMP page)
        {
            userAddin = addin;
            iSwApp = (ISldWorks)userAddin.SwApp;
            ppage = page;
        }

        //Implement these methods from the interface
        public void AfterClose()
        {
            //This function must contain code, even if it does nothing, to prevent the
            //.NET runtime environment from doing garbage collection at the wrong time.
            int IndentSize;
            IndentSize = System.Diagnostics.Debug.IndentSize;
            System.Diagnostics.Debug.WriteLine(IndentSize);
        }

        public void OnCheckboxCheck(int id, bool status)
        {
            airfoil oldairfoil = userAddin.CurrentAirfoil;
            airfoil newairfoil = null;
            if (id==MirrorCheckID)
            {
                newairfoil = new airfoil(null, oldairfoil.NACA, oldairfoil.chord, oldairfoil.twist, oldairfoil.twistloc, status);
                userAddin.CurrentAirfoil = newairfoil;
                userAddin.RemoveAirfoil(oldairfoil);
                userAddin.AirfoilPMP.Show(newairfoil);
                userAddin.GenerateAirfoil(newairfoil);
            }
        }

        public void OnClose(int reason)
        {
            if(reason != (int)swPropertyManagerPageCloseReasons_e.swPropertyManagerPageClose_Okay)
            {
                userAddin.swModelDocExt.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
            }
        }

        public void OnComboboxEditChanged(int id, string text)
        {

        }

        public int OnActiveXControlCreated(int id, bool status)
        {
            return -1;
        }

        public void OnButtonPress(int id)
        {
            
        }

        public void OnComboboxSelectionChanged(int id, int item)
        {

        }

        public void OnGroupCheck(int id, bool status)
        {

        }

        public void OnGroupExpand(int id, bool status)
        {

        }

        public bool OnHelp()
        {
            string helppath;
            System.Windows.Forms.Form helpForm = new System.Windows.Forms.Form();

            // Specify a url path or a path to a chm file
            helppath = "https://github.com/brheinfelder/Wingit";
            //helppath = "C:\\Program Files\\SolidWorks Corp\\SOLIDWORKS (2)\\api\\apihelp.chm";

            System.Windows.Forms.Help.ShowHelp(helpForm, helppath);

            return true;
        }

        public void OnListboxSelectionChanged(int id, int item)
        {

        }

        public bool OnNextPage()
        {
            return true;
        }

        public void OnNumberboxChanged(int id, double val)
        {
            airfoil oldairfoil = userAddin.CurrentAirfoil;
            airfoil newairfoil = null;
            if (id==ChordBoxID)
            {
                newairfoil = new airfoil(null, oldairfoil.NACA, val, oldairfoil.twist, oldairfoil.twistloc, oldairfoil.mirror);
            }
            else if(id == TwistBoxID)
            {
                newairfoil = new airfoil(null, oldairfoil.NACA, oldairfoil.chord, val, oldairfoil.twistloc, oldairfoil.mirror);
            }
            else if(id == TwistLocBoxID)
            {
                newairfoil = new airfoil(null, oldairfoil.NACA, oldairfoil.chord, oldairfoil.twist, val, oldairfoil.mirror);
            }
            userAddin.CurrentAirfoil = newairfoil;
            userAddin.RemoveAirfoil(oldairfoil);
            userAddin.AirfoilPMP.Show(newairfoil);
            userAddin.GenerateAirfoil(newairfoil);
        }

        public void OnNumberBoxTrackingCompleted(int id, double val)
        {

        }

        public void OnOptionCheck(int id)
        {

        }

        public bool OnPreviousPage()
        {
            return true;
        }

        public void OnSelectionboxCalloutCreated(int id)
        {

        }

        public void OnSelectionboxCalloutDestroyed(int id)
        {

        }

        public void OnSelectionboxFocusChanged(int id)
        {

        }

        public void OnSelectionboxListChanged(int id, int item)
        {
            // When a user selects entities to populate the selection box, display a popup cursor.
            ppage.swPropertyPage.SetCursor((int)swPropertyManagerPageCursors_e.swPropertyManagerPageCursors_Advance);
        }

        public void OnTextboxChanged(int id, string text)
        {
            airfoil oldairfoil = userAddin.CurrentAirfoil;
            airfoil newairfoil = null;
            if(id== NACABoxID)
            {
                newairfoil = new airfoil(null, text, oldairfoil.chord, oldairfoil.twist, oldairfoil.twistloc, oldairfoil.mirror);
                userAddin.CurrentAirfoil = newairfoil;
                userAddin.RemoveAirfoil(oldairfoil);
                userAddin.AirfoilPMP.Show(newairfoil);
                userAddin.GenerateAirfoil(newairfoil);
            }
        }

        public void AfterActivation()
        {

        }

        public bool OnKeystroke(int Wparam, int Message, int Lparam, int Id)
        {
            return true;
        }

        public void OnPopupMenuItem(int Id)
        {

        }

        public void OnPopupMenuItemUpdate(int Id, ref int retval)
        {

        }

        public bool OnPreview()
        {
            return true;
        }

        public void OnSliderPositionChanged(int Id, double Value)
        {

        }

        public void OnSliderTrackingCompleted(int Id, double Value)
        {

        }

        public bool OnSubmitSelection(int Id, object Selection, int SelType, ref string ItemText)
        {
            return true;
        }

        public bool OnTabClicked(int Id)
        {
            return true;
        }

        public void OnUndo()
        {

        }

        public void OnWhatsNew()
        {

        }


        public void OnGainedFocus(int Id)
        {

        }

        public void OnListboxRMBUp(int Id, int PosX, int PosY)
        {

        }

        public void OnLostFocus(int Id)
        {

        }

        public void OnRedo()
        {

        }

        public int OnWindowFromHandleControlCreated(int Id, bool Status)
        {
            return 0;
        }


    }
}
