
using System;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Timers;
using System.Reflection;
using System.Xml;
using System.Xml.Linq;

using RH = Rhino;
using RC = Rhino.Commands;
using RUI = Rhino.UI;


using SF = System.Windows.Forms;

using EF = Eto.Forms;
using ED = Eto.Drawing;


#if CSX==false
namespace Libx.InstantMode;
#endif

using Keys = System.Windows.Forms.Keys;



#region CONFIG


class Config : INotifyPropertyChanged
{

    /*/
        TODO: Trier les séquences pour effectuer des recherches binaire.
    /*/

    ObservableCollection <Sequence> _sequences = new ();
    public ObservableCollection <Sequence> Sequences => _sequences;

    public bool SequenceExists (Sequence seq)
    {
        foreach (var sequence in _sequences)
        {
            if (seq.Equals (sequence)) {
                return true;
            }
        }
        return false;
    }

    public Sequence? Find (Sequence seq)
    {
        foreach (var sequence in _sequences)
        {
            if (seq.Equals (sequence)) {
                return sequence;
            }
        }
        return null;
    }

    int _delay = 300;
    public int Delay
    {
        get => _delay;
        set {
            if (value == _delay) return;
            _delay = value;
            Emit ();
        }
    }


    #region IO

    public void ExportToXml (string filepath)
    {
        var root = new XElement ("ROOT",
            new XElement (nameof (Sequences),
                from seq in _sequences
                select new XElement ("Item",
                    new XElement (nameof (seq.AliasString), seq.AliasString),
                    new XElement (nameof (seq.CommandString), seq.CommandString)
                )
            ),
            new XElement (nameof (Delay), Delay)
        );

        var doc = new XDocument (root);
        try {
            doc.Save (filepath);
        } catch (Exception e) {
            DBG.Log (e.ToString ());
        }
    }

    public void importXml (string filepath)
    {
        XDocument doc;
        try {
            doc = XDocument.Load (filepath);
        } catch (Exception e) {
            DBG.Log (e.ToString ());
            return;
        }

        if (doc.Root is not XElement root)
            return;

        if (root.Element (nameof (this.Sequences)) is XElement xsequences)
        {
            _sequences.Clear ();
            foreach (var item in xsequences.Elements ())
            {
                var a = item.Element (nameof (Sequence.AliasString));
                if (a is null) continue;

                var alias = a.Value;
                if (string.IsNullOrWhiteSpace (alias)) continue;

                var c = item.Element (nameof (Sequence.CommandString));
                _sequences.Add (new Sequence (
                    alias,
                    c == null ? "" : (c.Value ?? "")
                ));
            }
        }

        if (root.Element (nameof (this.Delay)) is XElement xdelay)
        {
            if (int.TryParse (xdelay.Value, out var delay))
                Delay = delay;
        }
    }

    public void Save (RH.PersistentSettings settings)
    {
        settings.SetStringDictionary (
            nameof (Sequences),
            (from s in _sequences select new KeyValuePair <string, string> (s.AliasString, s.CommandString)).ToArray ()
        );

        settings.SetInteger (nameof (Delay), _delay);
    }

    public void Read (RH.PersistentSettings settings)
    {
        var pairs = settings.GetStringDictionary (nameof (Sequences), null);
        if (pairs is null) {
            _sequences = new ObservableCollection <Sequence> ();
        } else {
            _sequences = new ObservableCollection <Sequence> (
                from p in pairs select new Sequence (p.Key, p.Value)
            );
        }

        _delay = settings.GetInteger (nameof (Delay), 300);

        Emit (nameof (Delay));
        Emit (nameof (Sequences));
    }

    public void InitialiseDefault (IEnumerable <Sequence> sequences, int delay)
    {
        _sequences = new ObservableCollection<Sequence> (sequences);
        _delay = delay;
        Emit (nameof (Delay));
        Emit (nameof (Sequences));
    }

    #endregion

    #region Events

    public event PropertyChangedEventHandler? PropertyChanged;
    void Emit ([CallerMemberName] string? name = null) { PropertyChanged?.Invoke (this, new (name)); }
    
    #endregion
    
    #region Form

    static ConfigDialog? _form = null;

    public void ShowConfigDialog (InstantModeObserver observer)
    {
        if (_form != null && _form.IsDisposed == false)
            return;

        _form = new ConfigDialog (this, observer) {
            Location = RUI.RhinoEtoApp.MainWindow.Location,
            Owner = RUI.RhinoEtoApp.MainWindow
        };
        _form.Closed += (sender, e) => { _form = null; };
        _form.Show ();
    }

    public void HideConfigDialog ()
    {
        if (_form == null)
            return;

        if (_form.IsDisposed) {
            _form = null;
            return;
        }

        _form.Close ();
        _form.Dispose ();
        _form = null;
    }

    #endregion
}


class ConfigDialog : EF.Form
{
    readonly Config _config;
    readonly InstantModeObserver _observer;

    readonly EF.GridView _commandsView;
    readonly EF.NumericStepper _nDelay;
    readonly EF.Button _btnAppend;
    readonly EF.Button _btnRemove;
    readonly EF.Button _btnCancel;

    public ConfigDialog (Config config, InstantModeObserver observer)
    {
        _config = config;
        _observer = observer;

        _commandsView = _CreateCommandsView ();
        _nDelay       = new EF.NumericStepper { MinValue = 10, MaxValue = 1000, Value = config.Delay };
        _btnAppend    = new EF.Button { Text = "Append" };
        _btnRemove    = new EF.Button { Text = "Remove", Enabled = false };
        _btnCancel    = new EF.Button { Text = "Cancel" };

        var header = new EF.StackLayout {

        };

        var footer = new EF.StackLayout {
            Spacing = 6,
            Padding = 6,
            Orientation = EF.Orientation.Horizontal,
            Items = {
                new EF.Label { Text = "Delay" },
                _nDelay,
                null,
                _btnAppend,
                _btnRemove
            }
        };

        Menu = new EF.MenuBar {
            Items = {
                new EF.ButtonMenuItem (new EF.Command (_OnBtnImport) { MenuText = "Import..." }),
                new EF.ButtonMenuItem (new EF.Command (_OnBtnExport) { MenuText = "Export..." }),
            }
        };

        Content = new EF.StackLayout
        {
            Orientation = EF.Orientation.Vertical,
            Items = {
                new EF.StackLayoutItem (_commandsView, EF.HorizontalAlignment.Stretch, expand: true),
                new EF.StackLayoutItem (footer, EF.HorizontalAlignment.Stretch)
            }
        };

        _nDelay.ValueBinding.Bind (config, nameof (config.Delay), EF.DualBindingMode.TwoWay);

        _commandsView.SelectionChanged += _OnGridSelectionChange;
        _commandsView.MouseDoubleClick += _OnItemDoubleClick;
        _btnAppend.Click += _OnBtnAppend;
        _btnRemove.Click += _OnBtnRemove;
    }

    EF.GridView _CreateCommandsView ()
    {
        var grid = new EF.GridView ();
        grid.ShowHeader = true;
        grid.DataStore = _config.Sequences;

        grid.Columns.Add (new EF.GridColumn {
            HeaderText = "Sequence",
            Editable = false,
            DataCell = new EF.TextBoxCell (nameof (Sequence.AliasString))
        });

        grid.Columns.Add (new EF.GridColumn {
            HeaderText = "Command",
            Editable = false,
            DataCell = new EF.TextBoxCell (nameof (Sequence.CommandString))
        });

        return grid;
    }

    // Affiche la boîte de dialogue d'édition.
    bool _ShowRecorder (Sequence seq)
    {
        var diag = new RecorderDialog(_observer, _config, seq) {
            Location = Location,
            Width = Width,
            Height = 400,
        };
        return diag.ShowModal (this);
    }

    void _OnBtnExport (object sender, EventArgs e)
    {
        var diag = new EF.SaveFileDialog {
            Filters = { "xml|*.xml" }
        };

        if (diag.ShowDialog (RUI.RhinoEtoApp.MainWindow) != EF.DialogResult.Ok)
            return;

        _config.ExportToXml (diag.FileName);
    }

    void _OnBtnImport (object sender, EventArgs e)
    {
        var diag = new EF.OpenFileDialog {
            Filters = { "xml|*.xml" }
        };

        if (diag.ShowDialog (RUI.RhinoEtoApp.MainWindow) != EF.DialogResult.Ok)
            return;

        _config.importXml (diag.FileName);
    }

    // Active le bouton Remove.
    void _OnGridSelectionChange (object sender, EventArgs e)
    {
        _btnRemove.Enabled = _commandsView.SelectedItems.ToArray ().Length != 0;
    }

    // Affiche la boîte de dialogue d'édition pour l'alias de commande sélectionné.
    void _OnItemDoubleClick (object sender, EF.MouseEventArgs e)
    {
        if (_commandsView.SelectedItem is Sequence seq)
        {
            var copy = new Sequence (seq.AliasString, seq.CommandString);
            if (_ShowRecorder (copy)) {
                seq.AliasString = copy.AliasString;
                seq.CommandString = copy.CommandString;
                _commandsView.Invalidate ();
            }
        }
    }

    // Affiche la boîte de dialogue d'édition avec un nouvel alias de commande.
    void _OnBtnAppend (object sender, EventArgs e)
    {
        var seq = new Sequence ("", "");
        if (_ShowRecorder (seq) && seq.IsValid) {
            _config.Sequences.Add (seq);
        }
    }

    // Supprime un alias de commande.
    void _OnBtnRemove (object sender, EventArgs e)
    {
        if (_commandsView.SelectedItem is Sequence seq) {
            _config.Sequences.Remove (seq);
        }
    }

}


class RecorderDialog : EF.Dialog <bool>, IRunner
{
    const int PADDING = 6;

    readonly InstantModeObserver _observer;
    readonly Config _config;
    readonly Sequence _sequence;
    readonly EF.TextBox _txtShortcut;
    readonly EF.TextArea _txtCommand;
    readonly EF.Button _btnOk;
    readonly EF.Button _btnCancel;

    public RecorderDialog (InstantModeObserver observer, Config config, Sequence sequence)
    {
        _observer = observer;
        _config   = config;
        _sequence = sequence;

        _txtShortcut = new EF.TextBox { Text = sequence.AliasString, ReadOnly = true };
        _txtCommand  = new EF.TextArea { Text = sequence.CommandString };
        _btnOk       = new EF.Button { Text = "OK" };
        _btnCancel   = new EF.Button { Text = "Cancel" };

        var footer = new EF.StackLayout
        {
            Orientation = EF.Orientation.Horizontal,
            Items = { null, _btnOk, _btnCancel }
        };

        Padding = new ED.Padding (PADDING);
        Content = new EF.StackLayout
        {
            Orientation = EF.Orientation.Vertical,
            Spacing = PADDING,
            Items = {
                new EF.StackLayoutItem (_txtShortcut, EF.HorizontalAlignment.Stretch),
                new EF.StackLayoutItem (_txtCommand, EF.HorizontalAlignment.Stretch, expand: true),
                new EF.StackLayoutItem (footer, EF.HorizontalAlignment.Stretch),
            }
        };

        _txtShortcut.GotFocus   += _lblShortcut_GotFocus;
        _txtShortcut.LostFocus  += _lblShortcut_LostFocus;
        _btnOk.Click            += _btnOk_Click;
        _btnCancel.Click        += _btnCancel_Click;
    }

    ED.Color _activeColor = ED.Colors.AliceBlue;
    ED.Color _transparentColor = ED.Colors.Transparent;
    ED.Color _invalidColor = ED.Colors.PaleVioletRed;

    void _SetShortcut (string shortcut)
    {
        var tmp = new Sequence (shortcut, "");
        var invalid = _sequence.Equals (tmp) == false && _config.SequenceExists (tmp);
        _txtShortcut.Text = shortcut;
        _txtShortcut.BackgroundColor = invalid ? _invalidColor : _activeColor;
        _btnOk.Enabled = invalid == false;
    }


    // Mets à jour le raccourci clavier associé et quitte.
    void _btnOk_Click (object sender, EventArgs e)
    {
        _sequence.AliasString = _txtShortcut.Text;
        _sequence.CommandString = _txtCommand.Text;
        Close (true);
    }

    // Quitte
    void _btnCancel_Click (object sender, EventArgs e)
    {
        Close (false);
    }

    protected override void OnClosed (EventArgs e)
    {
        _StopRecorder ();
        base.OnClosed (e);
    }


    void _lblShortcut_GotFocus (object sender, EventArgs e)
    {
        _StartRecorder ();
        _txtShortcut.BackgroundColor = _activeColor;
    }

    void _lblShortcut_LostFocus (object sender, EventArgs e)
    {
        _StopRecorder ();
        _txtShortcut.BackgroundColor = _transparentColor;
    }

    // Active les entrées du clavier
    void _StartRecorder ()
    {
        _observer.SetTargets (this.NativeHandle, this);
    }

    // Désactive les entrées du clavier
    void _StopRecorder ()
    {
        _observer.RestoreDefaultTargets ();
    }

    // Fonction exécuter lorsqu'un raccourci clavier est capté.
    public void Run (Snapshot snapshot)
    {
        // L'utilisateur du timer dans Debouncer impose de changer de thread.
        EF.Application.Instance.Invoke (() => {
            if (_txtShortcut.HasFocus)
                _SetShortcut (snapshot.ToString ());
        });
    }
}


#endregion


#region OBSERVER


class InstantModeObserver
{
    readonly Config _config;

    public InstantModeObserver (Config config)
    {
        _config = config;
        _keyboardListener = new ();
    }

    const int INSTANT_MODE = 0;
    const int CMDLINE_MODE = 1;

    readonly KeyboardHook _keyboardListener;


    public void Start ()
    {
        if (_keyboardListener.IsEnabled)
            return;

        DBG.Log ("Activation du plugin");

        RC.Command.BeginCommand += _OnBeginCommand;
        RC.Command.EndCommand   += _OnEndCommand;
        RH.RhinoApp.Idle        += _OnIdle_SetFocusToMainWindow;

        KeyboardHook.OnKeyDown   += _OnKeyDown;
        KeyboardHook.OnLongPress += _OnLongPress;
        KeyboardHook.OnKeyUp     += _OnKeyUp;
        _keyboardListener.EnableEvents ();
    }

    public void Stop ()
    {
        if (_keyboardListener.IsEnabled == false)
            return;

        DBG.Log ("Désactivation du plugin");

        RC.Command.BeginCommand -= _OnBeginCommand;
        RC.Command.EndCommand   -= _OnEndCommand;
        RH.RhinoApp.Idle        -= _OnIdle_SetFocusToMainWindow;

        _keyboardListener.DisableEvents ();
        KeyboardHook.OnKeyDown   -= _OnKeyDown;
        KeyboardHook.OnLongPress -= _OnLongPress;
        KeyboardHook.OnKeyUp     -= _OnKeyUp;
    }


    public static bool _IsApplicationKey (Keys key)
    {
        if (
            // [L/R]ShiftKey, [L/R]ControlKey, [L/R]Menu
            Keys.LShiftKey <= key && key <= Keys.RMenu ||
            Keys.F1 <= key && key <= Keys.F24 ||
            key == Keys.Shift ||
            key == Keys.Control ||
            key == Keys.Alt
        ) return true;
        return false;
    }

    public static bool _IsInstantKey (Keys key)
    {
        if (
            Keys.A <= key && key <= Keys.Z ||
            // 0...9 * + Separator - . /
            Keys.NumPad0 <= key && key <= Keys.Divide ||
            key == Keys.OemBackslash
        ) return true;

        return false;
    }


    int _mode = INSTANT_MODE;

    void _SetInstantMode (int mode)
    {
        _mode = mode;
        RH.RhinoApp.WriteLine ("Toggle Mode to " + (_mode == INSTANT_MODE ? "instant commands" : "command line"));
        
        // Je suppose qu'il y a des événements internes à Rhino qui remettent le focus sur la ligne de commande.
        // if (_mode == INSTANT_MODE) {
        //
        //     // Lors de l'événement _OnEndCommand, ceci n'a aucun effet.
        //     RH.RhinoApp.SetFocusToMainWindow (); 
        //
        //     // Malheureusement la ligne suivante ne marche pas mieux.
        //     // RH.RhinoApp.InvokeOnUiThread (() => { RH.RhinoApp.SetFocusToMainWindow (); });
        // }
        // C'est moche mais j'ai dû mettre en place _OnIdle_SetFocusToMainWindow
    }

    void _OnIdle_SetFocusToMainWindow (object sender, EventArgs e)
    {
        // Permet de remettre le focus sur l'application après l'exécution d'une commande.
        if (_mode == INSTANT_MODE)
        {
            // Handle de la fenêtre active.
            var h = UnsafeNativeMethods.GetForegroundWindow ();
            
            // Récupère le titre de la fenêtre active.
            var sb = new StringBuilder (8); // 8 == "Command".Length + 1;
            UnsafeNativeMethods.GetWindowText ( h, sb, 8 );

            // Si la fenêtre active et la ligne de commande alors mettre le focus sur le viewport.
            if (sb.ToString () == "Command")
                RH.RhinoApp.SetFocusToMainWindow ();
        }
    }


    void _OnBeginCommand (object ender, RC.CommandEventArgs e)
    {
        _mode = CMDLINE_MODE;
    }

    void _OnEndCommand (object ender, RC.CommandEventArgs e)
    {
        // Vérifie qu'il ne s'agit pas d'une sous commande telle que _SetView.
        if (RC.Command.GetCommandStack ().Length > 1)
            return;

        RH.RhinoApp.Write ("_OnEndCommand ");
        _SetInstantMode (INSTANT_MODE);
    }


    /*/
        Permet d'outrepasser le comportement par défaut en exécutant la fonction en IRunner.Run
        Utilisé par la fenêtre de configuration afin de lire et d'écrire les séquences de touches.
    /*/

    IRunner? Runner = null;

    public void SetTargets (IntPtr windowHandle, IRunner runner)
    {
        DBG.Log ("Set custom window handle");
        _keyboardListener.SetHandleContraint (windowHandle);
        Runner = runner;
    }

    public void RestoreDefaultTargets ()
    {
        DBG.Log ("Set default window handle");
        _keyboardListener.SetDefaultHandleContraint ();
        Runner = null;
    }


    readonly Snapshot _snapshot = new ();
    readonly Debouncer _debouncer = new ();
    
    void _OnKeyDown (Keys key, ref bool handled)
    {
        if (Runner is not null && _IsInstantKey (key)) {

            // Cela m'ennuie de devoir tester cette condition à chaque fois
            // même quand la boîte de dialogue n'est pas présente
            // ce qui sera souvent le cas.

            // Ajoute toutes les touches enfoncées à l'instantané.
            _snapshot.AddKey (key);
            
            // handled = true;
            return;
        }

        if (RC.Command.InCommand () || RC.Command.InScriptRunnerCommand ()) {

            // DBG.Log ("in cmd");

        } else if (_IsApplicationKey (key)) {

            _mode = CMDLINE_MODE;

        } else if (key == Keys.Space) {

            // Bascule d'un mode à l'autre.
            _SetInstantMode (_mode == INSTANT_MODE ? CMDLINE_MODE : INSTANT_MODE);

            // Évite de répéter la commande lorsque l'on passe du mode instantané au mode ligne de commande.
            handled = true; 

        } else if (_mode == INSTANT_MODE && _IsInstantKey (key)) {

            // Empêche l'écriture dans la ligne de commande.
            handled = true;

            // Ajoute toutes les touches enfoncées à l'instantané.
            _snapshot.AddKey (key);

        } else {

            DBG.Log ("NOTHING: " + key.ToString ());

        }
    }

    void _OnLongPress (Keys key, ref bool handled)
    {
        if (_mode == INSTANT_MODE)
            handled = true; // Empêche l'écriture sur la ligne de commande.
    }

    void _OnKeyUp (Keys key, ref bool handled)
    {
        // DBG.Log ("--UP-- " + key.ToString ());

        if (_mode == INSTANT_MODE) {

            if (_IsInstantKey (key)) {
                _snapshot.NextCombination ();
                _debouncer.Debounce (_HandleSnapshots, _config.Delay);
            }

        } else if (_IsApplicationKey (key)) {

            _SetInstantMode (INSTANT_MODE);

        }
    }


    void _HandleSnapshots ()
    {
        // DBG.Log (_snapshot.ToString ());

        if (Runner is not null) {

            Runner.Run (_snapshot);
             
        } else {

            foreach (var sequence in _config.Sequences)
            {
                if (_snapshot.EqualsSequence (sequence)) {
                    sequence.Command ();
                    break;
                }
            }

        }

        _snapshot.Clear ();
    }

}


interface IRunner
{
    void Run (Snapshot snapshot);
}


/*/
    OnKeyDown
        Add to the snapshot all the pressed keys:
        if (_mode == INSTANT_MODE && _IsInstantKey (key))
            _snapshots.AddKey (key);
        
    OnKeyUp
        Memorize the pressed key combination and move to the next combination:
        if (_mode == INSTANT_MODE && _IsInstantKey (key))
            _snapshots.NextCombination ();

            After a delay, execute the key sequences.
            _debouncer.Debounce (_HandleSnapshots);
/*/
class Snapshot
{
    byte _seq = 0;
    readonly byte[] _count = new byte[5];
    readonly Keys[][] _combinations = new Keys[5][];

    public Snapshot ()
    {
        for (var i = 0 ; i < 5 ; i++ ) {
            _combinations[i] = new Keys[5];
            for (var j = 0 ; j < 5 ; j++ )
                _combinations[i][j] = Keys.None;
        }
    }

    /// <summary>
    /// Passe à la prochaine combinaison uniquement si la combinaison actuelle possède des touches enfoncer.
    /// </summary>
    public void NextCombination ()
    {
        if (_count[_seq] > 0 && _seq < 4)
            _seq++;
    }

    /// <summary>
    /// Ajoute une touche enfoncée à la combinaison actuelle.
    /// </summary>
    public void AddKey (Keys key)
    {
        if (_count[_seq] < 4) {
            _combinations[_seq][_count[_seq]++] = key;
        }
    }

    /// <summary>
    /// Mais à zéro toutes les combinaisons.
    /// </summary>
    public void Clear ()
    {
        for (var i = 0 ; i <= _seq ; i++ )
        {
            for (var j = 0 ; j <= _count[i] ; j++ )
                _combinations[i][j] = Keys.None;
            _count[i] = 0;
        }
        _seq = 0;
    }

    public bool IsEmpty ()
    {
        return _seq == 0 && _count[0] == 0;
    }


    public bool EqualsSequence (Sequence sequence)
    {
        if (_seq != sequence.Length)
            return false;

        var i = 0;
        foreach (var combination in sequence)
        {
            if (combination.Count != _count[i])
                return false;
        
            var keys = combination._keys;
            var j = 0;
            for ( ; j < _count[i] ; j++)
            {
                if (System.Array.IndexOf (_combinations[i], keys[j]) < 0)
                // if (keys[j] != _combinations[i][j]) 
                    return false;
            }
            if (j != _count[i])
                return false;

            i++;
        }
        return i == _seq;
    }


    public override string ToString ()
    {
        var s = new StringBuilder (5*5);

        for (var i = 0 ; i < _seq ; i++ )
        {
            for (var j = 0 ; j < _count[i] ; j++ )
                s.Append (_combinations[i][j].ToString () + "+");

            if (s.Length > 0)
                s[s.Length-1] = ',';
        }

        return s.Length > 0 ? s.ToString (0, s.Length-1) : "";
    }
}


#endregion


#region DEBOUNCER


class Debouncer
{
    private readonly Timer timer;
    private Action? action;

    public Debouncer()
    {
        timer = new Timer(300);
        timer.Elapsed += OnTimedEvent;
        timer.AutoReset = false;
    }

    public void Debounce(Action action, int interval)
    {
        this.action = action;
        timer.Stop();
        timer.Interval = interval;
        timer.Start();
    }

    private void OnTimedEvent(Object source, ElapsedEventArgs e)
    {
        action?.Invoke();
    }
}


#endregion


#region COMBINATION SEQUENCE


/// <summary>
/// Describes key or key combination sequences.
/// e.g. Control+Z,Z, 'A,B,C' 'Alt+R,S', 'Shift+R,Alt+K'
/// </summary>
class Sequence : IEnumerable <Combination>
{

    public Sequence (string text, string command)
    {
        AliasString = text;
        CommandString = command;
    }


    public bool IsValid
        => string.IsNullOrWhiteSpace (ToString ()) == false
        && string.IsNullOrWhiteSpace (CommandString) == false;


    private Combination[] _elements;

    /// <summary> Nombre de combinaisons dans la séquence. </summary>
    public int Length => _elements.Length;

    public string AliasString
    {
        get => ToString ();
        set {
            _elements = (
                from comb in (value ?? "").Split (',')
                where string.IsNullOrWhiteSpace (comb) == false
                select new Combination (comb)
            ).ToArray ();
        }
    }

    public string CommandString { get; set; }

    public void Command ()
    {
        // L'utilisation du Debouncer et de Systeme.Timer impose de revenir sur le tread principal.
        RH.RhinoApp.InvokeOnUiThread (() => {
            RH.RhinoApp.RunScript (CommandString, echo: false);
            RH.RhinoApp.SetFocusToMainWindow ();
        });
    }


    public override string ToString () => string.Join <Combination> (",", _elements);


    public IEnumerator <Combination> GetEnumerator () => _elements.Cast <Combination> ().GetEnumerator ();

    IEnumerator IEnumerable.GetEnumerator () => GetEnumerator ();


    protected bool Equals (Sequence other)
    {
        return _elements.Length == other._elements.Length
            && _elements.SequenceEqual (other._elements);
    }

    public override bool Equals (object obj)
    {
        if( ReferenceEquals (this, obj) ) return true;
        return obj.GetType () == GetType () && Equals ((Sequence) obj);
    }

    public override int GetHashCode ()
    {
        unchecked
        {
            return (_elements.Length + 13) ^
                    ((_elements.Length != 0
                            ? _elements[0].GetHashCode () ^ _elements[_elements.Length - 1].GetHashCode ()
                            : 0) * 397);
        }
    }
}


class Combination : IEnumerable <Keys>
{
    public readonly Keys[] _keys;

    public Combination (string chord)
    {
        var keys = (
            from p in (chord ?? "").Split ('+')
            where string.IsNullOrWhiteSpace (p) == false
            select Enum.Parse (typeof (Keys), p)
        ).Cast <Keys> ();

        _keys = (
            from k in keys ?? Array.Empty <Keys> ()
            orderby k
            select _NormalizeKey (k)
        ).ToArray ();
    }

    public int Count => _keys.Length;


    public IEnumerator <Keys> GetEnumerator () => _keys.Cast <Keys> ().GetEnumerator ();

    IEnumerator IEnumerable.GetEnumerator () => GetEnumerator ();


    public override string ToString () => string.Join ("+", _keys);

    protected bool Equals (Combination other)
    {
        return _keys.Length == other._keys.Length
            && _keys.SequenceEqual (other._keys);
    }

    public override bool Equals (object obj)
    {
        if( ReferenceEquals (null, obj) ) return false;
        if( ReferenceEquals (this, obj) ) return true;
        if( obj.GetType () != GetType () ) return false;
        return Equals ((Combination) obj);
    }

    public override int GetHashCode ()
    {
        unchecked
        {
            return (_keys.Length + 13)
                    ^ ((_keys.Length != 0
                            ? (int) _keys[0] ^ (int) _keys[_keys.Length - 1]
                            : 0
                    ) * 397);
        }
    }


    static Keys _NormalizeKey (Keys key)
    {
        if( (key & Keys.LControlKey) == Keys.LControlKey ||
            (key & Keys.RControlKey) == Keys.RControlKey ) return Keys.Control;

        if( (key & Keys.LMenu) == Keys.LMenu ||
            (key & Keys.RMenu) == Keys.RMenu ) return Keys.Alt;

        if( (key & Keys.LShiftKey) == Keys.LShiftKey ||
            (key & Keys.RShiftKey) == Keys.RShiftKey ) return Keys.Shift;

        return key;
    }
}


#endregion



#region HOOK

/*/

     Example:
     - https://github.com/mcneel/rhino-developer-samples/blob/6/rhinocommon/cs/SampleCsWinForms/WinHooks.cs
     - https://github.com/mcneel/rhino-developer-samples/blob/6/rhinocommon/cs/SampleCsWinForms/Forms/SampleCsModelessTabFix.cs

/*/


delegate IntPtr HookHandler ( int nCode, IntPtr wParam, IntPtr lParam );

enum WindowsHookType
{
    Keyboard = 2,
    Mouse = 7,
    LowLevelKeyboard = 13,
    LowLevelMouse = 14
}

class Hook
{
    //protected abstract IntPtr HookFunc (int nCode, IntPtr wParam, IntPtr lParam);

    public Hook (WindowsHookType hookType ) {
        HookType = (int) hookType;
    }

    public IntPtr NextHook ( int nCode, IntPtr wParam, IntPtr lParam) {
        return UnsafeNativeMethods.CallNextHookEx (m_hook, nCode, wParam, lParam);
    }

    public bool IsEnabled => m_hook != IntPtr.Zero;

    readonly int HookType;
    IntPtr m_hook = IntPtr.Zero;
    HookHandler? m_winproc;

    public void Install (HookHandler proc)
    {
        if( m_hook != IntPtr.Zero ) return;
        m_hook = AttachToSystem ( proc );

        // RhinoApp.Closing +=  (object sender, EventArgs e) => {
        //      UnInstall ();
        // };
    }

    public IntPtr AttachToApplication ( HookHandler proc )
    {
        m_winproc = proc;
        return UnsafeNativeMethods.SetWindowsHookEx (
            HookType,
            m_winproc,
            IntPtr.Zero,
            UnsafeNativeMethods.GetCurrentThreadId ()
        );
    }
    
    public IntPtr AttachToSystem ( HookHandler proc )
    {
        m_winproc = proc;
        return UnsafeNativeMethods.SetWindowsHookEx (
            HookType,
            m_winproc,
            Process.GetCurrentProcess().MainModule.BaseAddress,
            0
        );
    }

    public void UnInstall ()
    {
        if( m_hook == IntPtr.Zero ) return;
        UnsafeNativeMethods.UnhookWindowsHookEx (m_hook);
        m_hook = IntPtr.Zero;
    }
}

static class UnsafeNativeMethods
{
        [DllImport ("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public  static extern IntPtr SetWindowsHookEx ( int idHook, HookHandler lpfn, IntPtr hMod, uint dwThreadId );

        [DllImport ("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs (UnmanagedType.Bool)]
        public static extern bool UnhookWindowsHookEx ( IntPtr hhk );

        [DllImport ("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr CallNextHookEx ( IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam );

        [DllImport ("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr GetModuleHandle (string lpModuleName);

        [DllImport ("user32.dll")]
        public static extern IntPtr GetForegroundWindow ();

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();

        [DllImport ("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetWindowText (IntPtr hWnd, StringBuilder lpString, int nMaxCount);

}


#endregion


#region KEYBOARD

class KeyboardState
{
    readonly bool[]  _keyStates = new bool[255];

    public void SetDown (int key)
    {
        if (key < 256)
            _keyStates[(int)key] = true;
    }

    public void SetUp (int key)
    {
        if (key < 256)
            _keyStates[key] = false;
    }

    public bool IsDown (Keys key)
    {
        bool IsDownRaw (Keys k) => _keyStates[(int)k];

        if ((int) key < 256)
            return IsDownRaw (key);

        return key switch
        {
            Keys.Alt     => IsDownRaw (Keys.LMenu) || IsDownRaw (Keys.RMenu),
            Keys.Shift   => IsDownRaw (Keys.LShiftKey) || IsDownRaw (Keys.RShiftKey),
            Keys.Control => IsDownRaw (Keys.LControlKey) || IsDownRaw (Keys.RControlKey),
            _ => false,
        };
    }

}


/// <summary>
/// Class for intercepting low level keyboard hooks
/// </summary>
sealed class KeyboardHook
{
    public KeyboardHook ()
    {
        m_hook = new Hook (WindowsHookType.LowLevelKeyboard);
    }

    readonly Hook m_hook;

    const double DelayDoublePress = 500;

    const double DelayLongPress = 1000;

    public static ModifierTypes Modifiers { get; private set; } = 0;

    public delegate void HookFuncArgs (Keys key, ref bool handled);

    public static event HookFuncArgs? OnKeyDown;
    public static event HookFuncArgs? OnKeyUp;
    public static event HookFuncArgs? OnDoublePress;
    public static event HookFuncArgs? OnLongPress;

    [Flags]
    public enum ModifierTypes
    {
        None    = 0,
        Control = 1,
        Shift   = 2,
        Alt     = 4
    }


    public bool IsEnabled => m_hook.IsEnabled;

    public bool EnableEvents ()
    {
        if ( m_hook.IsEnabled ) return false;
        SetDefaultHandleContraint ();
        m_hook.Install (HookFunc);
        return true;
    }

    public bool DisableEvents ()
    {
        if ( m_hook.IsEnabled == false ) return false;
        m_hook.UnInstall ();
        return true;
    }


    #region Keys

    // Losely based on http://www.pinvoke.net/default.aspx/Enums/VK.html

    //                 Left      |      Right
    // Shift   : 0xA0 (1010 0000) 0xA1 (1010 0001)
    // Control : 0xA2 (1010 0010) 0xA3 (1010 0011)
    // Menu    : 0xA4 (1010 0100) 0xA5 (1010 0101)
    //           0xA6 (1010 0110) 0xA7 (1010 0111)
    //                            0xFE (1111 1110)

    const int LShift   = (int) Keys.LShiftKey;   // = 0xA0;
    const int LControl = (int) Keys.LControlKey; // = 0xA2;
    const int LMenu    = (int) Keys.LMenu;       // = 0xA4;

    const int KeyMask = (int) Keys.KeyCode; // = 0xFFFF;

    const int Shift   = (int) Keys.Shift;   // = 0x10000;
    const int Control = (int) Keys.Control; // = 0x20000;
    const int Alt     = (int) Keys.Alt;     // = 0x40000;

    public int GetKeyCode (string skey)
    {
        skey = skey switch
        {
            "+"    => "Add",
            "-"    => "Subtract",
            "/"    => "Divide",
            "*"    => "Multiply",
            "0"    => "NumPad0",
            "1"    => "NumPad1",
            "2"    => "NumPad2",
            "3"    => "NumPad3",
            "4"    => "NumPad4",
            "5"    => "NumPad5",
            "6"    => "NumPad6",
            "7"    => "NumPad7",
            "8"    => "NumPad8",
            "9"    => "NumPad9",
            "Ctrl" => "Control",
            "Alt"  => "Menu",
            "Maj"  => "Shift",
            _      => skey
        };

        var key = (int) Enum.Parse (typeof (Keys), skey);
        return (key & 0xFE) switch
        {
            LControl => Control,
            LMenu    => Alt,
            LShift   => Shift,
            _        => key
        };
    }

    #endregion


    public readonly KeyboardState State = new ();

    IntPtr _winHandle;

    public void SetDefaultHandleContraint ()
    {
        _winHandle = Rhino.RhinoApp.MainWindowHandle ();
    }
    public void SetHandleContraint (IntPtr h)
    {
        _winHandle = h;
    }

    IntPtr _lastMsg = IntPtr.Zero;
    int    _lastTime;
    int    _lastLongTime;
    int    _lastKey;

    IntPtr HookFunc ( int nCode, IntPtr wParam, IntPtr lParam )
    {
        if ( nCode != 0 ) return CallNext ();

        // Si la fenêtre active n'est pas rhino.
        var h = UnsafeNativeMethods.GetForegroundWindow ();
        if ( _winHandle != h )
        {
            // Si fenêtre active n'est pas le menu contextuel de la ligne de commande de Rhino.
            // https://discourse.mcneel.com/t/get-the-drop-down-list-handle-of-the-command-line/74629/3
            var sb = new StringBuilder (21); // 21 == "CRhinoUiPopUpListWnd".Length + 1;
            UnsafeNativeMethods.GetWindowText ( h, sb, 21 );
            if (sb.ToString () != "CRhinoUiPopUpListWnd") {
                Modifiers = 0;
                // DBG.Log ("On rejette l'entrée");
                return CallNext ();
            }
        }

        var infos = (KBDLLHOOKSTRUCT) Marshal.PtrToStructure (lParam, typeof (KBDLLHOOKSTRUCT));

        var handled = false;
        int delta;

        ModifierTypes modifier = 0;
        switch( (Keys) infos.VkCode )
        {
        case Keys.LShiftKey: 
        case Keys.RShiftKey:
            modifier = ModifierTypes.Shift;
            break;

        case Keys.LControlKey: 
        case Keys.RControlKey:
            modifier = ModifierTypes.Control;
            break;

        case Keys.Alt  : 
        case Keys.Menu : 
        case Keys.LMenu: 
        case Keys.RMenu:
            modifier = ModifierTypes.Alt;
            break;
        }

        const int Keydown    = 0x100;
        const int SysKeydown = 0x104;
        const int Keyup      = 0x101;
        const int SysKeyup   = 0x105;
        var msg = (int)wParam;

        if (_lastKey == infos.VkCode && _lastMsg == wParam && (Modifiers & modifier) == 0)
        {

            delta = infos.Time - _lastLongTime;
            if ( _lastLongTime != 0 && delta > DelayLongPress) {

                // DBG.Log ("LONG PRESS");
                _lastLongTime = 0;
                OnLongPress?.Invoke ((Keys) infos.VkCode, ref handled);

            } else if (State.IsDown ((Keys)infos.VkCode)) {

                // Supprime les répétitions de touches
                handled = true;

            } else {

                return CallNext ();

            }
        }
        else if (msg == Keydown || msg == SysKeydown)
        {
            State.SetDown (infos.VkCode);

            OnKeyDown?.Invoke ((Keys) infos.VkCode, ref handled);

            delta = infos.Time - _lastTime;
            if( _lastKey == infos.VkCode && delta < DelayDoublePress )
            OnDoublePress?.Invoke ((Keys) infos.VkCode, ref handled);

            _lastMsg      = wParam;
            _lastKey      = infos.VkCode;
            _lastLongTime = _lastTime = infos.Time;

            Modifiers |= modifier;
        }
        else if (msg == Keyup || msg == SysKeyup)
        {
            _lastMsg      = wParam;
            _lastLongTime = 0;

            Modifiers &= ~modifier;

            OnKeyUp?.Invoke ((Keys) infos.VkCode, ref handled);

            State.SetUp (infos.VkCode);
        }

        return handled ? (IntPtr) 1 : CallNext ();

        IntPtr CallNext () => m_hook.NextHook (nCode, wParam, lParam);
    }

    // https://learn.microsoft.com/fr-fr/windows/win32/api/winuser/ns-winuser-kbdllhookstruct
    [StructLayout (LayoutKind.Sequential)]
    readonly struct KBDLLHOOKSTRUCT
    {
        /// <summary>
        /// Code de clé virtuelle. Le code doit être une valeur comprise entre 1 et 254.
        /// </summary>
        public readonly int    VkCode;
        public readonly int    ScanCode;
        public readonly int    Flags;
        public readonly int    Time;
        public readonly IntPtr dwExtraInfo;
    }
}


#endregion



#region DBG


static class DBG
{
    
    static void _Emit (string group, MethodBase mT, object? message = null)
    {
        RH.RhinoApp.WriteLine ($"[{group} {mT.DeclaringType.Name}.{mT.Name}] {message}");
    }

    [Conditional("DEBUG")]
    public static void Log (params object?[] messages)
    {
        _Emit ("",
            new StackTrace().GetFrame (1).GetMethod(),
            string.Join (" ", from o in messages select o == null ? "null" : ""+o)
        );
    }
}


#endregion
