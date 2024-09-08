#r "C:\Program Files\Rhino 7\System\RhinoCommon.dll"
#r "C:\Program Files\Rhino 7\System\RhinoWindows.dll"

#r "C:\Program Files\Rhino 7\Plug-ins\Grasshopper\Grasshopper.dll"

#r "C:\Program Files\Rhino 7\System\Eto.dll"
#r "C:\Program Files\Rhino 7\System\Eto.Wpf.dll"

#r "E:\Projet\Rhino\Libx\Out\net48\Libx.Node.Task.dll"
#r "E:\Projet\Rhino\Libx\Out\net48\Libx.Grasshopper.Task.gha"

#r "System.Windows.Forms"

#load "./1 InstantMode.cs"
#nullable enable


using System;
using System.Collections;
using System.Collections.Generic;

using Libx.Task;
using Libx.Task.Grasshopper;



#region ACTIVATION


[Input] bool Enabled;
[Input] bool ShowConfig; 

if (Enabled) { _instantMode.Start (); }
else { _Unload (); }

if (ShowConfig && Enabled) {
    _instantMode.ShowConfigDialog ();
}

[OnUnLoad] public static void _Unload ()
{
    _instantMode.Stop ();
    _instantMode.HideConfigDialog ();
}


#endregion


static InstantMode _instantMode = new ();

static IEnumerable <Sequence> GetDefaultMapAlias () => new Sequence[]
{
    new Sequence ("L"    , "!_Polyline"),
    new Sequence ("L,L"  , "!_Curve"),
    new Sequence ("L+C"  , "!_Circle"),
    new Sequence ("L+A"  , "!_Arc"),
    new Sequence ("C+A"  , "!_Arc"),
    new Sequence ("L+H"  , "!_Polygon"),
    new Sequence ("L+R"  , "!_Rectangle"),
    new Sequence ("L,E"  , "!_Extend "),

    new Sequence ("M" , "_Move"),

    // "<" ">"
    new Sequence ("OemBackslash"              , "!_Isolate"),
    new Sequence ("OemBackslash,OemBackslash" , "!_Unisolate"),

    new Sequence ("P"         , "'_CPlane "),
    new Sequence ("P,P"       , "'_CPlane _Object"),
    new Sequence ("P,P,P"     , "'_CPlane _3Point "),

    new Sequence ("J"         , "!_Join"),
    new Sequence ("J,J"       , "!_Explode"),

    new Sequence ("NumPad7"   , "'_IsoMetric _NW"),
    new Sequence ("NumPad9"   , "'_IsoMetric _NE"),
    new Sequence ("NumPad1"   , "'_IsoMetric _SW"),
    new Sequence ("NumPad3"   , "'_IsoMetric _SE"),

    new Sequence ("NumPad8"   , "'_SetView _World _Top"),
    new Sequence ("NumPad2"   , "'_SetView _World _Bottom"),
    new Sequence ("NumPad4"   , "'_SetView _World _Left"),
    new Sequence ("NumPad6"   , "'_SetView _World _Right"),
    new Sequence ("NumPad5"   , "! _-ViewportProperties _enter _Projection _Toggle _enter"),

    new Sequence ("Decimal"   , "'_Zoom _Selected"),

    new Sequence ("D+C"   , "! _DupEdge"),
    new Sequence ("D+E"   , "! _DupEdge"),
    new Sequence ("D+F"   , "! _DupBorder"),
    new Sequence ("D+S"   , "! _DupBorder"),

    new Sequence ("E+C"   , "! _ExtrudeCrv"),
    new Sequence ("E+S"   , "! _ExtrudeSrf"),

    new Sequence ("O+L"   , "! _Offset"),
    new Sequence ("O+C"   , "! _Offset"),
    new Sequence ("O+S"   , "! _OffsetSrf"),
    new Sequence ("O+S,S" , "! _VariableOffsetSrf"),

    new Sequence ("B+C"   , "!_CurveBoolean"),
    new Sequence ("B+S"   , "_BooleanUnion"),
    new Sequence ("B+S,S" , "_BooleanDifference"),

    new Sequence ("B,B"   , "_BooleanUnion"),

    new Sequence ("F"     , "! _PlanarSrf"),
};



class InstantMode
{
    public InstantMode ()
    {
        _config = new Config();
        _observer = new (_config);
        InitialiseDefault ();
    }

    readonly Config _config;
    readonly InstantModeObserver _observer;

    public void InitialiseDefault ()
    {
        _config.InitialiseDefault (GetDefaultMapAlias (), 300);
    }

    public void Start () { _observer.Start (); }
    public void Stop () { _observer.Stop (); }

    public void ShowConfigDialog ()
    {
        _config.ShowConfigDialog (_observer);
    }
    public void HideConfigDialog ()
    {
        _config.HideConfigDialog ();
    }
}
