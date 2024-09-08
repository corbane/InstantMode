


using System;
using System.Runtime.InteropServices;

using RH = Rhino;
using RP = Rhino.PlugIns;
using RC = Rhino.Commands;


[assembly: Guid("ED076A41-553E-4BA7-88F3-16680982533F")]

namespace Libx.InstantMode;



public class InstantMode : RP.PlugIn
{

    #nullable disable
    public static InstantMode Instance { get; private set; }
    #nullable enable

    public override RP.PlugInLoadTime LoadTime => RP.PlugInLoadTime.AtStartup;


    public InstantMode ()
    {
        Instance = this;
        _config = new Config();
        _observer = new (_config);
        RH.RhinoApp.Initialized += _OnAppInitialized;
        RH.RhinoApp.Closing += _OnAppClosing;
    }
    
    protected override RP.LoadReturnCode OnLoad (ref string errorMessage)
    {
        _config.Read (Settings);
        return base.OnLoad (ref errorMessage);
    }
    
    void _OnAppInitialized (object sender, EventArgs e) { Start (); }

    void _OnAppClosing (object sender, EventArgs e) { Stop (); }


    readonly Config _config;
    readonly InstantModeObserver _observer;

    public void Start () { _observer.Start (); }
    public void Stop () { _observer.Stop (); }

    public void ShowConfigDialog ()
    {
        _config.ShowConfigDialog (_observer);
    }
}


public class StartCommand : RC.Command
{
    public override string EnglishName => "StartInstantMode";

    protected override RC.Result RunCommand (RH.RhinoDoc doc, RC.RunMode mode)
    {
        InstantMode.Instance.Start ();
        return RC.Result.Success;
    }
}


public class StopCommand : RC.Command
{
    public override string EnglishName => "StopInstantMode";

    protected override RC.Result RunCommand (RH.RhinoDoc doc, RC.RunMode mode)
    {
        InstantMode.Instance.Stop ();
        return RC.Result.Success;
    }
}


public class ShowOptionsCommand : RC.Command
{
    public override string EnglishName => "ShowInstantModeOptions";

    protected override RC.Result RunCommand (RH.RhinoDoc doc, RC.RunMode mode)
    {
        InstantMode.Instance.ShowConfigDialog ();
        return RC.Result.Success;
    }
}
