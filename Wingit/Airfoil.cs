using SolidWorks.Interop.sldworks;
using System;

namespace WingIt
{
    public class airfoil
    {
        public SelectionSet Selection; //Selection Set including all airfoil geometry
        public string NACA; // Airfoil NACA Designation
        public double chord; // Airfoil Chord
        public double twist; // Airfoil Twist
        public double twistloc; //Location of fixed twist point as %
        public bool mirror; // Mirror Airfoil?

        public airfoil(SelectionSet SelectionSet, string NACADesignation, double AirfoilChord, double AirfoilTwist, double AirfoilTwistLoc, bool AirfoilMirror)
        {
            Selection = SelectionSet;
            NACA = NACADesignation;
            chord = AirfoilChord;
            twist = AirfoilTwist;
            twistloc = AirfoilTwistLoc;
            mirror = AirfoilMirror;
        }
    }
}
