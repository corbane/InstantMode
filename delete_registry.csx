#nullable enable

using System;
using Microsoft.Win32;


var todo = new Dictionary <string, string> ();
RegistryKey HKEY_USERS = Registry.Users;

foreach (string profil in HKEY_USERS.GetSubKeyNames())
{
    var pluginsKey = HKEY_USERS.OpenSubKey (profil + "\\SOFTWARE\\McNeel\\Rhinoceros\\7.0\\Plug-Ins");
    if (pluginsKey == null) continue;

    foreach (string id in pluginsKey.GetSubKeyNames())
    {
        if (pluginsKey.OpenSubKey (id) is not RegistryKey idKey ||
            idKey.GetValue ("EnglishName") is not string name ||
            name != "Libx.InstantMode"
        )   continue;

        todo.Add (name, profil + "\\SOFTWARE\\McNeel\\Rhinoceros\\7.0\\Plug-Ins\\" + id);
        idKey.Close ();
    }

    pluginsKey.Close ();
}

if (todo.Count == 0)
{
    Console.WriteLine ("No plugins found.");
    return;
}

foreach (var item in todo)
    Console.WriteLine (item.Key);

Console.WriteLine ("Do you delete Keys [y|n]?");
var consoleKeyInfo = Console.ReadKey ();
Console.WriteLine ();
if (consoleKeyInfo.KeyChar == 'y' || consoleKeyInfo.KeyChar == 'Y')
{
    foreach (var item in todo)
    {
        HKEY_USERS.DeleteSubKeyTree (item.Value);
        Console.WriteLine ("DELETE " +  item.Value);
    }
}