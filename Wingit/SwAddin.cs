using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using SolidWorksTools;
using SolidWorksTools.File;
using System;
using System.Windows.Forms;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using WingIt;


namespace Wingit
{
    /// <summary>
    /// Summary description for Wingit.
    /// </summary>
    [Guid("ad576ee2-3b70-4edf-80e2-84b57e49502b"), ComVisible(true)]
    [SwAddin(
        Description = "Wing Builder Add-In for Solidworks",
        Title = "Wingit",
        LoadAtStartup = true
        )]
    public class SwAddin : ISwAddin
    {
        #region Local Variables
        ISldWorks iSwApp = null;
        ICommandManager iCmdMgr = null;
        public ModelDoc2 swModelDoc;
        public ModelDocExtension swModelDocExt;
        FeatureManager swFeatureManager = default(FeatureManager);
        int addinID = 0;
        BitmapHandler iBmp;
        int registerID;
        bool boolstatus;
        public airfoil CurrentAirfoil;
        public OpenFileDialog browser = new OpenFileDialog();

        public const int mainCmdGroupID = 5;
        public const int mainItemID1 = 0;
        public const int mainItemID2 = 1;
        public const int mainItemID3 = 2;
        public const int flyoutGroupID = 91;

        string[] mainIcons = new string[6];
        string[] icons = new string[6];

        #region Event Handler Variables
        Hashtable openDocs = new Hashtable();
        SolidWorks.Interop.sldworks.SldWorks SwEventPtr = null;
        #endregion

        #region Property Manager Variables
        public InsertAirfoilPMP AirfoilPMP = null;
        #endregion


        // Public Properties
        public ISldWorks SwApp
        {
            get { return iSwApp; }
        }
        public ICommandManager CmdMgr
        {
            get { return iCmdMgr; }
        }

        public Hashtable OpenDocs
        {
            get { return openDocs; }
        }

        #endregion

        #region SolidWorks Registration
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            #region Get Custom Attribute: SwAddinAttribute
            SwAddinAttribute SWattr = null;
            Type type = typeof(SwAddin);

            foreach (System.Attribute attr in type.GetCustomAttributes(false))
            {
                if (attr is SwAddinAttribute)
                {
                    SWattr = attr as SwAddinAttribute;
                    break;
                }
            }

            #endregion

            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
                addinkey.SetValue(null, 0);

                addinkey.SetValue("Description", SWattr.Description);
                addinkey.SetValue("Title", SWattr.Title);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                addinkey = hkcu.CreateSubKey(keyname);
                addinkey.SetValue(null, Convert.ToInt32(SWattr.LoadAtStartup), Microsoft.Win32.RegistryValueKind.DWord);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
                System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: SWattr is null.\n\"" + nl.Message + "\"");
            }

            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);

                System.Windows.Forms.MessageBox.Show("There was a problem registering the function: \n\"" + e.Message + "\"");
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                hklm.DeleteSubKey(keyname);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                hkcu.DeleteSubKey(keyname);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + nl.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + nl.Message + "\"");
            }
            catch (System.Exception e)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + e.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + e.Message + "\"");
            }
        }

        #endregion

        #region ISwAddin Implementation
        public SwAddin()
        {
        }

        public bool ConnectToSW(object ThisSW, int cookie)
        {
            iSwApp = (ISldWorks)ThisSW;
            addinID = cookie;

            //Setup callbacks
            iSwApp.SetAddinCallbackInfo(0, this, addinID);

            #region Setup the Command Manager
            iCmdMgr = iSwApp.GetCommandManager(cookie);
            AddCommandMgr();
            #endregion

            #region Setup the Event Handlers
            SwEventPtr = (SolidWorks.Interop.sldworks.SldWorks)iSwApp;
            openDocs = new Hashtable();
            AttachEventHandlers();
            #endregion

            #region Setup Sample Property Manager
            AddPMP();
            #endregion

            return true;
        }

        public bool DisconnectFromSW()
        {
            RemoveCommandMgr();
            RemovePMP();
            DetachEventHandlers();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(iCmdMgr);
            iCmdMgr = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(iSwApp);
            iSwApp = null;
            //The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers 
            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return true;
        }
        #endregion

        #region UI Methods
        public void AddCommandMgr()
        {
            ICommandGroup cmdGroup;
            if (iBmp == null)
                iBmp = new BitmapHandler();
            Assembly thisAssembly;
            int cmdIndex0, cmdIndex1;
            string Title = "WingIt", ToolTip = "WingIt Add-In";


            int[] docTypes = new int[] { (int)swDocumentTypes_e.swDocPART };

            thisAssembly = System.Reflection.Assembly.GetAssembly(this.GetType());


            int cmdGroupErr = 0;
            bool ignorePrevious = false;

            object registryIDs;
            //get the ID information stored in the registry
            bool getDataResult = iCmdMgr.GetGroupDataFromRegistry(mainCmdGroupID, out registryIDs);

            int[] knownIDs = new int[2] { mainItemID1, mainItemID2 };

            if (getDataResult)
            {
                if (!CompareIDs((int[])registryIDs, knownIDs)) //if the IDs don't match, reset the commandGroup
                {
                    ignorePrevious = true;
                }
            }

            cmdGroup = iCmdMgr.CreateCommandGroup2(mainCmdGroupID, Title, ToolTip, "WingIt for SOLIDWORKS", -1, ignorePrevious, ref cmdGroupErr);

            // Add bitmaps to your project and set them as embedded resources or provide a direct path to the bitmaps.
            icons[0] = iBmp.CreateFileFromResourceBitmap("Wingit.toolbar20x.png", thisAssembly);
            icons[1] = iBmp.CreateFileFromResourceBitmap("Wingit.toolbar32x.png", thisAssembly);
            icons[2] = iBmp.CreateFileFromResourceBitmap("Wingit.toolbar40x.png", thisAssembly);
            icons[3] = iBmp.CreateFileFromResourceBitmap("Wingit.toolbar64x.png", thisAssembly);
            icons[4] = iBmp.CreateFileFromResourceBitmap("Wingit.toolbar96x.png", thisAssembly);
            icons[5] = iBmp.CreateFileFromResourceBitmap("Wingit.toolbar128x.png", thisAssembly);

            mainIcons[0] = iBmp.CreateFileFromResourceBitmap("Wingit.mainicon_20.png", thisAssembly);
            mainIcons[1] = iBmp.CreateFileFromResourceBitmap("Wingit.mainicon_32.png", thisAssembly);
            mainIcons[2] = iBmp.CreateFileFromResourceBitmap("Wingit.mainicon_40.png", thisAssembly);
            mainIcons[3] = iBmp.CreateFileFromResourceBitmap("Wingit.mainicon_64.png", thisAssembly);
            mainIcons[4] = iBmp.CreateFileFromResourceBitmap("Wingit.mainicon_96.png", thisAssembly);
            mainIcons[5] = iBmp.CreateFileFromResourceBitmap("Wingit.mainicon_128.png", thisAssembly);

            cmdGroup.MainIconList = mainIcons;
            cmdGroup.IconList = icons;

            int menuToolbarOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);
            cmdIndex0 = cmdGroup.AddCommandItem2("Insert Airfoil", 1, "Insert an Airfoil into the Current Sketch", "Insert Airfoil", 2, "ShowAirfoilPMP", "EnablePMP", mainItemID1, menuToolbarOption);

            cmdGroup.HasToolbar = true;
            cmdGroup.HasMenu = true;
            cmdGroup.Activate();

            bool bResult;



            FlyoutGroup flyGroup = iCmdMgr.CreateFlyoutGroup2(flyoutGroupID, "Dynamic Flyout", "Flyout Tooltip", "Flyout Hint",
              cmdGroup.MainIconList, cmdGroup.IconList, "FlyoutCallback", "FlyoutEnable");


            flyGroup.AddCommandItem("FlyoutCommand 1", "test", 0, "FlyoutCommandItem1", "FlyoutEnableCommandItem1");

            flyGroup.FlyoutType = (int)swCommandFlyoutStyle_e.swCommandFlyoutStyle_Simple;


            foreach (int type in docTypes)
            {
                CommandTab cmdTab;

                cmdTab = iCmdMgr.GetCommandTab(type, Title);

                if (cmdTab != null & !getDataResult | ignorePrevious)//if tab exists, but we have ignored the registry info (or changed command group ID), re-create the tab.  Otherwise the ids won't matchup and the tab will be blank
                {
                    bool res = iCmdMgr.RemoveCommandTab(cmdTab);
                    cmdTab = null;
                }

                //if cmdTab is null, must be first load (possibly after reset), add the commands to the tabs
                if (cmdTab == null)
                {

                    cmdTab = iCmdMgr.AddCommandTab(type, Title);

                    CommandTabBox cmdBox = cmdTab.AddCommandTabBox();

                    int[] cmdIDs = new int[1];
                    int[] TextType = new int[1];

                    cmdIDs[0] = cmdGroup.get_CommandID(cmdIndex0);

                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    bResult = cmdBox.AddCommands(cmdIDs, TextType);
                }

            }

        }

        public void RemoveCommandMgr()
        {
            iBmp.Dispose();

            iCmdMgr.RemoveCommandGroup(mainCmdGroupID);
            iCmdMgr.RemoveFlyoutGroup(flyoutGroupID);
        }

        public bool CompareIDs(int[] storedIDs, int[] addinIDs)
        {
            List<int> storedList = new List<int>(storedIDs);
            List<int> addinList = new List<int>(addinIDs);

            addinList.Sort();
            storedList.Sort();

            if (addinList.Count != storedList.Count)
            {
                return false;
            }
            else
            {

                for (int i = 0; i < addinList.Count; i++)
                {
                    if (addinList[i] != storedList[i])
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        public Boolean AddPMP()
        {
            AirfoilPMP = new InsertAirfoilPMP(this);
            return true;
        }

        public Boolean RemovePMP()
        {
            AirfoilPMP = null;
            return true;
        }

        #endregion

        #region UI Callbacks
        public airfoil GenerateAirfoil(airfoil NewAirfoil)
        {
            //Document Setup
            swModelDoc = (ModelDoc2)SwApp.ActiveDoc;
            SketchManager swSketchMgr = swModelDoc.SketchManager;
            swModelDocExt = swModelDoc.Extension;
            swFeatureManager = (FeatureManager)swModelDoc.FeatureManager;
            airfoil currentairfoil = NewAirfoil;
            double chord = NewAirfoil.chord;
            int err = 0;
            int Resolution = 100;
            double[] x = new double[2 * Resolution + 1];
            double[] y = new double[2 * Resolution + 1];
            double[] z = new double[2 * Resolution + 1];

            if (NewAirfoil.NACA.Length==4)
            {
                (x, y, z) = NACA4(NewAirfoil);
            }
            else if(NewAirfoil.NACA.Length==5)
            {
                (x, y, z) = NACA5(NewAirfoil);
            }

            //Draw Airfoil
            double[] pointdata = new double[x.Length * 3];
            for (int i = 0; i < x.Length; i++)
            {
                pointdata[3 * i] = x[i] * chord;
                pointdata[3 * i + 1] = y[i] * chord;
                pointdata[3 * i + 2] = 0;
            }
            SketchSegment Airfoil;
            object statuses = null;
            Airfoil = (SketchSegment)swSketchMgr.CreateSpline3(pointdata, null, null, false, out statuses);
            swSketchMgr.CreateLine(x[0] * chord, y[0] * chord, 0, x[x.Length - 1] * chord, y[y.Length - 1] * chord, 0);
            Airfoil.Select4(true, null);
            currentairfoil.Selection = swModelDocExt.SaveSelection(out err);
            return currentairfoil;
        }

        public (double[] x, double[] y, double[] z) NACA4(airfoil NewAirfoil)
        {
            int Resolution = 100;
            double[] x = new double[2 * Resolution + 1];
            double[] y = new double[2 * Resolution + 1];
            double[] z = new double[2 * Resolution + 1];

            double[] camber = new double[Resolution + 1];
            double[] thickness = new double[Resolution + 1];

            double chord = NewAirfoil.chord;

            double a0 = 1.4845;
            double a1 = 0.6300;
            double a2 = 1.7580;
            double a3 = 1.4215;
            double a4 = 0.5075;

            int NACA = Int32.Parse(NewAirfoil.NACA);

            //Parse 4 Digit NACA Number
            double M = ((double)(NACA / 1000)) / 100; // Maximum value of camber line in hundreths of chord
            int rem = NACA % 1000;
            double P = ((double)(rem / 100)) / 10; // chordwise position of the maximum camber in tenths of the chord
            double XX = ((double)(NACA % 100)) / 100; // Maxium thickness in percent chord

            // Generate Chordwise Thickness and Camber Values
            for (int i = 0; i <= Resolution; i++)
            {
                x[i] = 1 - ((double)i / Resolution);
                thickness[i] = (XX) * (a0 * Math.Pow(x[i], 0.5) - a1 * x[i] - a2 * Math.Pow(x[i], 2) + a3 * Math.Pow(x[i], 3) - a4 * Math.Pow(x[i], 4));
                if (x[i] < P)
                {
                    camber[i] = (M / Math.Pow(P, 2)) * (2 * P * x[i] - Math.Pow(x[i], 2));
                }
                else
                {
                    camber[i] = (M / Math.Pow(1 - P, 2)) * (1 - 2 * P + 2 * P * x[i] - Math.Pow(x[i], 2));
                }
            }

            //Generate Airfoil Point Coordinates
            for (int i = 0; i <= Resolution; i++)
            {
                x[i] = 1 - ((double)i / Resolution);
                y[i] = camber[i] + thickness[i];
                z[i] = 0;
            }
            for (int i = Resolution; i >= 0; i--)
            {
                x[i + Resolution] = (double)i / Resolution;
                y[i + Resolution] = camber[Resolution - i] - thickness[Resolution - i];
                z[i] = 0;
            }

            //Mirror Airfoil
            if (NewAirfoil.mirror)
            {
                for (int i = 0; i < x.Length; i++)
                {
                    x[i] = x[i] * -1;
                }
            }

            //Twist Airfoil
            double xtwist = NewAirfoil.twistloc / 100;
            double ytwist = 0;
            if (xtwist < P)
            {
                ytwist = (M / Math.Pow(P, 2)) * (2 * P * xtwist - Math.Pow(xtwist, 2));
            }
            else
            {
                ytwist = (M / Math.Pow(1 - P, 2)) * (1 - 2 * P + 2 * P * xtwist - Math.Pow(xtwist, 2));
            }
            for (int i = 0; i < x.Length; i++)
            {
                if (NewAirfoil.mirror)
                {
                    (x[i], y[i]) = TwistPoint(-xtwist, ytwist, x[i], y[i], -NewAirfoil.twist);
                }
                else
                {
                    (x[i], y[i]) = TwistPoint(xtwist, ytwist, x[i], y[i], NewAirfoil.twist);
                }
            }
            return (x, y, z);
        }

        public (double[] x, double[] y, double[] z) NACA5(airfoil NewAirfoil)
        {
            int Resolution = 100;
            double[] x = new double[2 * Resolution + 1];
            double[] y = new double[2 * Resolution + 1];
            double[] z = new double[2 * Resolution + 1];

            double[] camber = new double[Resolution + 1];
            double[] thickness = new double[Resolution + 1];

            double a0 = 1.4845;
            double a1 = 0.6300;
            double a2 = 1.7580;
            double a3 = 1.4215;
            double a4 = 0.5075;

            double chord = NewAirfoil.chord;

            //r Dictionary
            Dictionary<int, double> rdictionary = new Dictionary<int, double>();
            rdictionary.Add(10, 0.0580);
            rdictionary.Add(20, 0.1260);
            rdictionary.Add(30, 0.2025);
            rdictionary.Add(40, 0.2900);
            rdictionary.Add(50, 0.3910);
            rdictionary.Add(21, 0.1300);
            rdictionary.Add(31, 0.2170);
            rdictionary.Add(41, 0.3180);
            rdictionary.Add(51, 0.4410);

            //k1 Dictionary
            Dictionary<int, double> k1dictionary = new Dictionary<int, double>();
            k1dictionary.Add(10, 361.400);
            k1dictionary.Add(20, 51.640);
            k1dictionary.Add(30, 15.957);
            k1dictionary.Add(40, 6.643);
            k1dictionary.Add(50, 3.230);
            k1dictionary.Add(21, 51.990);
            k1dictionary.Add(31, 15.793);
            k1dictionary.Add(41, 6.520);
            k1dictionary.Add(51, 3.191);

            //k1/k2 Dictionary
            Dictionary<int, double> k1k2dictionary = new Dictionary<int, double>();
            k1k2dictionary.Add(21, 0.000764);
            k1k2dictionary.Add(31, 0.00677);
            k1k2dictionary.Add(41, 0.0303);
            k1k2dictionary.Add(51, 0.1355);

            int NACA = Int32.Parse(NewAirfoil.NACA);

            //Parse 5 Digit NACA Number
            int L = NACA / 10000;
            int P = (NACA / 1000) % 10;
            int Q = (NACA / 100) % 10;
            double XX = ((double)(NACA % 100)) / 100;

            //Choose r and k1 values
            int digits = 10 * P + Q;
            double r=0;
            double k1=0;
            double k1k2=0;
            if (Q==0)
            {
                r = rdictionary[digits];
                k1 = k1dictionary[digits];
            }
            else if (Q==1)
            {
                r = rdictionary[digits];
                k1 = k1dictionary[digits];
                k1k2 = k1k2dictionary[digits];
            }

            //Generate Thickness and Camber
            for (int i = 0; i <= Resolution; i++)
            {
                x[i] = 1 - ((double)i / Resolution);
                thickness[i] = (XX) * (a0 * Math.Pow(x[i], 0.5) - a1 * x[i] - a2 * Math.Pow(x[i], 2) + a3 * Math.Pow(x[i], 3) - a4 * Math.Pow(x[i], 4));
                if (x[i]<r)
                {
                    if(Q==0)
                    {
                        camber[i] = (k1 / 6) * (Math.Pow(x[i], 3) - 3 * r * Math.Pow(x[i], 2) + Math.Pow(r, 2) * (3 - r) * x[i]);
                    }
                    else if(Q==1)
                    {
                        camber[i] = (k1 / 6) * (Math.Pow((x[i] - r), 3) - k1k2 * Math.Pow((1 - r), 3) * x[i] - Math.Pow(r, 3) * x[i] + Math.Pow(r, 3));
                    }
                }
                else if (x[i]>=r)
                {
                    if (Q == 0)
                    {
                        camber[i] = (k1 / 6) * Math.Pow(r, 3) * (1 - x[i]);
                    }
                    else if (Q == 1)
                    {
                        camber[i] = (k1 / 6) * (k1k2 * Math.Pow((x[i] - r), 3) - k1k2 * Math.Pow((1 - r), 3) * x[i] - Math.Pow(r, 3) * x[i] + Math.Pow(r, 3));
                    }
                }
            }

            //Generate Airfoil Point Coordinates
            for (int i = 0; i <= Resolution; i++)
            {
                x[i] = 1 - ((double)i / Resolution);
                y[i] = camber[i] + thickness[i];
                z[i] = 0;
            }
            for (int i = Resolution; i >= 0; i--)
            {
                x[i + Resolution] = (double)i / Resolution;
                y[i + Resolution] = camber[Resolution - i] - thickness[Resolution - i];
                z[i] = 0;
            }

            //Mirror Airfoil
            if (NewAirfoil.mirror)
            {
                for (int i = 0; i < x.Length; i++)
                {
                    x[i] = x[i] * -1;
                }
            }

            //Twist Airfoil
            double xtwist = NewAirfoil.twistloc / 100;
            double ytwist = 0;
            if (xtwist < r)
            {
                if (Q == 0)
                {
                    ytwist = (k1 / 6) * (xtwist * xtwist * xtwist - 3 * r * xtwist * xtwist + r * r * (3 - r) * xtwist);
                }
                else if (Q == 1)
                {
                    ytwist = (k1 / 6) * ((xtwist - r) * (xtwist - r) * (xtwist - r) - k1k2 * (1 - r) * (1 - r) * (1 - r) * xtwist - r * r * r * xtwist + r * r * r);
                }
            }
            else if (xtwist >= r)
            {
                if (Q == 0)
                {
                    ytwist = (k1 / 6) * r * r * r * (1 - xtwist);
                }
                else if (Q == 1)
                {
                    ytwist = (k1 / 6) * (k1k2 * (xtwist - r) * (xtwist - r) * (xtwist - r) - k1k2 * (1 - r) * (1 - r) * (1 - r) * xtwist - r * r * r * xtwist + r * r * r);
                }
            }
            for (int i = 0; i < x.Length; i++)
            {
                if (NewAirfoil.mirror)
                {
                    (x[i], y[i]) = TwistPoint(-xtwist, ytwist, x[i], y[i], -NewAirfoil.twist);
                }
                else
                {
                    (x[i], y[i]) = TwistPoint(xtwist, ytwist, x[i], y[i], NewAirfoil.twist);
                }
            }

            return (x, y, z);
        }

        public (double x, double y) TwistPoint(double xorigin, double yorigin, double x, double y, double angle)
        {
            double xcoord = x - xorigin;
            double ycoord = y - yorigin;
            double magnitude = Math.Sqrt(xcoord * xcoord + ycoord * ycoord);
            double currentangle = Math.Atan2(ycoord, xcoord);
            double newangle = currentangle - angle;
            y = Math.Sin(newangle) * magnitude + yorigin;
            x = Math.Cos(newangle) * magnitude + xorigin;
            return (x, y);
        }

        public void RemoveAirfoil(airfoil Airfoil)
        {
            swModelDoc = (ModelDoc2)SwApp.ActiveDoc;
            swModelDocExt = swModelDoc.Extension;
            swModelDocExt.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
        }

        public void PopupCallbackFunction()
        {
            bool bRet;

            bRet = iSwApp.ShowThirdPartyPopupMenu(registerID, 500, 500);
        }

        public int PopupEnable()
        {
            if (iSwApp.ActiveDoc == null)
                return 0;
            else
                return 1;
        }

        public void TestCallback()
        {
            Debug.Print("Test Callback, CSharp");
        }

        public int EnableTest()
        {
            if (iSwApp.ActiveDoc == null)
                return 0;
            else
                return 1;
        }

        public void ShowAirfoilPMP()
        {
            if (AirfoilPMP != null)
            {
                airfoil NACA0012 = new airfoil(null, "0012", 1, 0, 0, false);
                AirfoilPMP.Show(NACA0012);
                CurrentAirfoil = GenerateAirfoil(NACA0012);
            }
        }

        public int EnablePMP()
        {
            if (iSwApp.ActiveDoc != null)
                return 1;
            else
                return 0;
        }

        public void FlyoutCallback()
        {
            FlyoutGroup flyGroup = iCmdMgr.GetFlyoutGroup(flyoutGroupID);
            flyGroup.RemoveAllCommandItems();

            flyGroup.AddCommandItem(System.DateTime.Now.ToLongTimeString(), "test", 0, "FlyoutCommandItem1", "FlyoutEnableCommandItem1");

        }
        public int FlyoutEnable()
        {
            return 1;
        }

        public void FlyoutCommandItem1()
        {
            iSwApp.SendMsgToUser("Flyout command 1");
        }

        public int FlyoutEnableCommandItem1()
        {
            return 1;
        }
        #endregion

        #region Event Methods
        public bool AttachEventHandlers()
        {
            AttachSwEvents();
            //Listen for events on all currently open docs
            AttachEventsToAllDocuments();
            return true;
        }

        private bool AttachSwEvents()
        {
            try
            {
                SwEventPtr.ActiveDocChangeNotify += new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
                SwEventPtr.DocumentLoadNotify2 += new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
                SwEventPtr.FileNewNotify2 += new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
                SwEventPtr.ActiveModelDocChangeNotify += new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
                SwEventPtr.FileOpenPostNotify += new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }



        private bool DetachSwEvents()
        {
            try
            {
                SwEventPtr.ActiveDocChangeNotify -= new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
                SwEventPtr.DocumentLoadNotify2 -= new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
                SwEventPtr.FileNewNotify2 -= new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
                SwEventPtr.ActiveModelDocChangeNotify -= new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
                SwEventPtr.FileOpenPostNotify -= new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }

        }

        public void AttachEventsToAllDocuments()
        {
            ModelDoc2 modDoc = (ModelDoc2)iSwApp.GetFirstDocument();
            while (modDoc != null)
            {
                if (!openDocs.Contains(modDoc))
                {
                    AttachModelDocEventHandler(modDoc);
                }
                else if (openDocs.Contains(modDoc))
                {
                    bool connected = false;
                    DocumentEventHandler docHandler = (DocumentEventHandler)openDocs[modDoc];
                    if (docHandler != null)
                    {
                        connected = docHandler.ConnectModelViews();
                    }
                }

                modDoc = (ModelDoc2)modDoc.GetNext();
            }
        }

        public bool AttachModelDocEventHandler(ModelDoc2 modDoc)
        {
            if (modDoc == null)
                return false;

            DocumentEventHandler docHandler = null;

            if (!openDocs.Contains(modDoc))
            {
                switch (modDoc.GetType())
                {
                    case (int)swDocumentTypes_e.swDocPART:
                        {
                            docHandler = new PartEventHandler(modDoc, this);
                            break;
                        }
                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        {
                            docHandler = new AssemblyEventHandler(modDoc, this);
                            break;
                        }
                    case (int)swDocumentTypes_e.swDocDRAWING:
                        {
                            docHandler = new DrawingEventHandler(modDoc, this);
                            break;
                        }
                    default:
                        {
                            return false; //Unsupported document type
                        }
                }
                docHandler.AttachEventHandlers();
                openDocs.Add(modDoc, docHandler);
            }
            return true;
        }

        public bool DetachModelEventHandler(ModelDoc2 modDoc)
        {
            DocumentEventHandler docHandler;
            docHandler = (DocumentEventHandler)openDocs[modDoc];
            openDocs.Remove(modDoc);
            modDoc = null;
            docHandler = null;
            return true;
        }

        public bool DetachEventHandlers()
        {
            DetachSwEvents();

            //Close events on all currently open docs
            DocumentEventHandler docHandler;
            int numKeys = openDocs.Count;
            object[] keys = new Object[numKeys];

            //Remove all document event handlers
            openDocs.Keys.CopyTo(keys, 0);
            foreach (ModelDoc2 key in keys)
            {
                docHandler = (DocumentEventHandler)openDocs[key];
                docHandler.DetachEventHandlers(); //This also removes the pair from the hash
                docHandler = null;
            }
            return true;
        }
        #endregion

        #region Event Handlers
        //Events
        public int OnDocChange()
        {
            return 0;
        }

        public int OnDocLoad(string docTitle, string docPath)
        {
            return 0;
        }

        int FileOpenPostNotify(string FileName)
        {
            AttachEventsToAllDocuments();
            return 0;
        }

        public int OnFileNew(object newDoc, int docType, string templateName)
        {
            AttachEventsToAllDocuments();
            return 0;
        }

        public int OnModelChange()
        {
            return 0;
        }

        #endregion
    }

}
