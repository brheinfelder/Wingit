using SolidWorks.Interop.sldworks;
using System;

namespace WingIt
{
    public class airfoil
    {
        public enum AirfoilType
        {
            NACA,
            Custom
        }

        public SelectionSet Selection; //Selection Set including all airfoil geometry
        public AirfoilType airfoiltype;
        public string NACA; // Airfoil NACA Designation
        public string airfoilpath;
        public string airfoilfilename;
        public double chord; // Airfoil Chord
        public double twist; // Airfoil Twist
        public double twistloc; //Location of fixed twist point as %
        public bool mirror; // Mirror Airfoil?
        public bool invertcamber;

        public airfoil(AirfoilType foiltype, SelectionSet SelectionSet, string NACADesignation, string filepath, double AirfoilChord, double AirfoilTwist, double AirfoilTwistLoc, bool AirfoilMirror)
        {
            airfoiltype = foiltype;
            Selection = SelectionSet;
            NACA = NACADesignation;
            airfoilpath = filepath;
            chord = AirfoilChord;
            twist = AirfoilTwist;
            twistloc = AirfoilTwistLoc;
            mirror = AirfoilMirror;
        }
    }
}
