using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;
using System.IO;
using System.Windows.Forms;
using WingIt;

namespace Wingit
{
    public class InsertAirfoilPMPHandler : IPropertyManagerPage2Handler9
    {
        ISldWorks iSwApp;
        SwAddin userAddin;
        InsertAirfoilPMP ppage;

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
            airfoil newairfoil = oldairfoil;
            if (id==InsertAirfoilPMP.MirrorCheckID)
            {
                newairfoil.mirror = status;
                userAddin.CurrentAirfoil = newairfoil;
                userAddin.RemoveAirfoil(oldairfoil);
                userAddin.AirfoilPMP.Show(newairfoil);
                userAddin.GenerateAirfoil(newairfoil);
            }
            else if (id==InsertAirfoilPMP.InvertCamberID)
            {
                newairfoil.invertcamber = status;
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
            if(id == InsertAirfoilPMP.ImportAirfoilID)
            {
                string filePath;

                OpenFileDialog dlg = new OpenFileDialog();
                dlg.RestoreDirectory = true;
                dlg.Filter = "Airfoil Data File (*.txt;*.csv;*.dat)|*.txt;*.csv;*.dat";
                dlg.FilterIndex = 0;

                if(dlg.ShowDialog()==DialogResult.OK)
                {
                    filePath = dlg.FileName;
                    airfoil oldairfoil = userAddin.CurrentAirfoil;
                    airfoil newairfoil = oldairfoil;
                    newairfoil.airfoiltype = airfoil.AirfoilType.Custom;
                    newairfoil.airfoilpath = filePath;
                    newairfoil.airfoilfilename = dlg.SafeFileName;
                    userAddin.CurrentAirfoil = newairfoil;
                    userAddin.RemoveAirfoil(oldairfoil);
                    userAddin.AirfoilPMP.Show(newairfoil);
                    userAddin.GenerateAirfoil(newairfoil);
                }
            }
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
            airfoil newairfoil = oldairfoil;
            if (id== InsertAirfoilPMP.ChordBoxID)
            {
                newairfoil.chord = val;
            }
            else if(id == InsertAirfoilPMP.TwistBoxID)
            {
                newairfoil.twist = val;
            }
            else if(id == InsertAirfoilPMP.TwistLocBoxID)
            {
                newairfoil.twistloc = val;
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
            airfoil oldairfoil = userAddin.CurrentAirfoil;
            airfoil newairfoil = oldairfoil;
            if (id==0)
            {
                newairfoil.airfoiltype = airfoil.AirfoilType.NACA;
                ppage.swPropertyPage.SetMessage3("Compatible with basic 4 and 5 digit NACA airfoils.",
                                            (int)swPropertyManagerPageMessageVisibility.swImportantMessageBox,
                                            (int)swPropertyManagerPageMessageExpanded.swMessageBoxExpand,
                                            "NACA Airfoil");
                ((IPropertyManagerPageControl)ppage.NACABoxLabel).Visible = true;
                ((IPropertyManagerPageControl)ppage.NACABox).Visible = true;
                ((IPropertyManagerPageControl)ppage.ImportAirfoil).Visible = false;
                ((IPropertyManagerPageControl)ppage.CustomAirfoilFileName).Visible = false;
                userAddin.CurrentAirfoil = newairfoil;
                userAddin.RemoveAirfoil(oldairfoil);
                userAddin.AirfoilPMP.Show(newairfoil);
                userAddin.GenerateAirfoil(newairfoil);
            }
            else if(id==1)
            {
                newairfoil.airfoiltype = airfoil.AirfoilType.Custom;
                ppage.swPropertyPage.SetMessage3("Normalized airfoil data must be in Lednicer or Selig format.",
                                            (int)swPropertyManagerPageMessageVisibility.swImportantMessageBox,
                                            (int)swPropertyManagerPageMessageExpanded.swMessageBoxExpand,
                                            "Custom Airfoil");
                ((IPropertyManagerPageControl)ppage.NACABoxLabel).Visible = false;
                ((IPropertyManagerPageControl)ppage.NACABox).Visible = false;
                ((IPropertyManagerPageControl)ppage.ImportAirfoil).Visible = true;
                ((IPropertyManagerPageControl)ppage.CustomAirfoilFileName).Visible = true;
                userAddin.CurrentAirfoil = newairfoil;
                userAddin.RemoveAirfoil(oldairfoil);
                userAddin.AirfoilPMP.Show(newairfoil);
                userAddin.GenerateAirfoil(newairfoil);
            }
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
            airfoil newairfoil = oldairfoil;
            if(id == InsertAirfoilPMP.NACABoxID)
            {
                newairfoil.NACA = text;
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
