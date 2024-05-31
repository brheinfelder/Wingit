using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;
using System.Windows.Forms;
using WingIt;
using static WingIt.airfoil;

namespace Wingit
{
    public class WingBuilderPMP
    {
        //Local Objects
        public IPropertyManagerPage2 swPropertyPage = null;
        WingBuilderPMPHandler handler = null;
        ISldWorks iSwApp = null;
        SwAddin userAddin = null;

        public WingBuilderPMP(SwAddin addin)
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

            handler = new WingBuilderPMPHandler(userAddin, this);
            swPropertyPage = (IPropertyManagerPage2)iSwApp.CreatePropertyManagerPage("Build Wing", options, handler, ref errors);
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

        protected void AddControls()
        {
        }

        public void Show(airfoil Airfoil)
        {
            if (swPropertyPage != null)
            {
                swPropertyPage.Show();
            }
        }
    }
}
