using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace ArgumentParser
{
  public class ArgParse
  {
    Dictionary<string, string> argumentDict = new Dictionary<string, string>();
    List<ArgItem> argumentList = new List<ArgItem>();

    public enum ArgParseType
    {
      Long = 0,
      Int = 1,
      DateTime = 2,
      String = 3,
      Boolean = 4,
      File = 5,
      Pid = 6, //process id
      Url = 7,
      Choice = 8
    };

    public ArgParse(params ArgItem[] _argClass)
    {
      foreach (ArgItem arg in _argClass)
      {
        argumentList.Add(arg);
      }
    }

    public ArgParse(List<ArgItem> _argList)
    {
      argumentList = _argList;
    }

    public void parse(string[] args)
    {
      string prev = null;
      foreach (string arg in args)
      {
        if (arg == "-h" || arg == "--help")
        {
          showHelp("");
          return;
        }

        if (arg.StartsWith("--") || arg.StartsWith("-"))
        {
          ArgItem item = getArgItemBySwitch(arg);

          if (item == null)
          {
            showHelp(string.Format("Unknown switch: {0}", arg));
            return;
          }

          prev = item.longName;
          if (argumentDict.ContainsKey(item.longName))
          {
            showHelp(string.Format("duplicate switch: {0}", item.longName));
            return;
          }

          // set the initial value to the default value specified
          argumentDict.Add(item.longName, item.defaultValue);
        }
        else
        {
          if (prev != null) // prev will be null on the first iteration of this loop
          {
            argumentDict[prev] = arg;
          }
        }
      }

      // sets the defaults if a value does not exist
      foreach (ArgItem argClass in argumentList)
      {
        if (argClass.defaultValue != null && argClass.defaultValue.Length > 0)
        {
          if (!argumentDict.ContainsKey(argClass.longName))
          {
            argumentDict.Add(argClass.longName, argClass.defaultValue);
          }
        }
      }

      // check is required
      foreach (ArgItem argClass in argumentList)
      {
        if (argClass.required)
        {
          if (!argumentDict.ContainsKey(argClass.longName))
          {
            showHelp(string.Format("Field {0} is required but is missing.", argClass.longName));
            return;
          }
        }
      }

      // check is valid type
      foreach (ArgItem argClass in argumentList)
      {
        if (argumentDict.ContainsKey(argClass.longName))
        {
          string dataType = argClass.dataType.ToString();
          string dataValue = argumentDict[argClass.longName];

          if (argClass.required || dataValue.Length > 0)
          {
            try
            {
              switch (argClass.argType)
              {
                case ArgParseType.Int:
                  Int32.Parse(dataValue);
                  break;
                case ArgParseType.Long:
                  Int64.Parse(dataValue);
                  break;
                case ArgParseType.DateTime:
                  DateTime.Parse(dataValue);
                  break;
                case ArgParseType.Boolean:
                  bool.Parse(dataValue);
                  break;
                case ArgParseType.File:
                  if (!File.Exists(dataValue))
                  {
                    throw new Exception("File does not exist");
                  }
                  break;
                case ArgParseType.Pid:
                  Int32 pid = Int32.Parse(dataValue);
                  Process proc = Process.GetProcessById(pid);
                  if (proc == null)
                  {
                    throw new Exception("Process does not exist");
                  }
                  break;
                case ArgParseType.Url:
                  try
                  {
                    Uri u = null;
                    Uri.TryCreate(dataValue, UriKind.Absolute, out u);
                    if (u == null)
                    {
                      throw new Exception("Invalid URL.");
                    }
                  }
                  catch (Exception)
                  {
                    throw new Exception("Invalid URL.");
                  }
                  break;
                case ArgParseType.Choice:
                  int foundAt = Array.IndexOf(argClass.validValues, dataValue);
                  if (foundAt == -1)
                  {
                    string validCsv = string.Join(",", argClass.validValues);
                    throw new Exception(string.Format("Invalid option.  The only valid options are : {0}", validCsv));
                  }
                  break;
                default:
                  break;
              }
            }
            catch (Exception e)
            {
              showHelp(string.Format("{0} - {1}", argClass.longName, e.Message));
              break;
            }
          }
        }
      }
    }

    public void showHelp(string msg)
    {
      bool hasOptional = false;
      if (msg.Length > 0)
      {
        Console.WriteLine("Error: {0}\n", msg);
      }

      Console.WriteLine("Usage: {0} <options>", AppDomain.CurrentDomain.FriendlyName);
      Console.WriteLine("\tRequired");
      foreach (ArgItem argClass in argumentList)
      {
        if (argClass.required == true)
        {
          Console.WriteLine("\t\t-{0}, --{1}\t{2}", argClass.shortName, argClass.longName, argClass.helpText);
        }
        else
        {
          hasOptional = true;
        }
      }

      if (hasOptional)
      {
        Console.WriteLine("\n\tOptional");
        foreach (ArgItem argClass in argumentList)
        {
          if (argClass.required == false)
          {
            Console.WriteLine("\t\t-{0}, --{1}\t{2}", argClass.shortName, argClass.longName, argClass.helpText);
          }
        }
      }

      if (Debugger.IsAttached)
      {
        Console.ReadLine();
      }

      Environment.Exit(0);
    }

    private ArgItem getArgItemBySwitch(string arg)
    {
      if (arg.StartsWith("--"))
      {
        foreach (ArgItem item in argumentList)
        {
          if (item.longName == arg.Substring(2))
          {
            return item;
          }
        }
      }
      else
      {
        foreach (ArgItem item in argumentList)
        {
          if (item.shortName == arg.Substring(1))
          {
            return item;
          }
        }
      }

      return null;
    }

    public T Get<T>(string key)
    {
      if (!argumentDict.ContainsKey(key))
        return default(T);

      try
      {
        return (T)Convert.ChangeType(argumentDict[key], typeof(T));
      }
      catch (Exception)
      {
        return default(T);
      }
    }
  }

  public class ArgItem
  {
    public string longName { get; set; } //double dash option --domain
    public string shortName { get; set; } //single dash option -d
    public bool required { get; set; }
    public string helpText { get; set; }
    public string defaultValue { get; set; }
    public Type dataType { get; set; }
    public ArgParse.ArgParseType argType { get; set; }
    public string[] validValues { get; set; }

    public ArgItem(string _longName, string _shortName, bool _required)
    {
      init(_longName, _shortName, _required, "", "", ArgParse.ArgParseType.String, new string[] { });
    }

    public ArgItem(string _longName, string _shortName, bool _required, string _helpText)
    {
      init(_longName, _shortName, _required, _helpText, "", ArgParse.ArgParseType.String, new string[] { });
    }
    public ArgItem(string _longName, string _shortName, bool _required, string _helpText, string _defaultValue)
    {
      init(_longName, _shortName, _required, _helpText, _defaultValue, ArgParse.ArgParseType.String, new string[] { });
    }

    public ArgItem(string _longName, string _shortName, bool _required, ArgParse.ArgParseType _dataType)
    {
      init(_longName, _shortName, _required, "", "", _dataType, new string[] { });
    }

    public ArgItem(string _longName, string _shortName, bool _required, string _helpText, ArgParse.ArgParseType _dataType)
    {
      init(_longName, _shortName, _required, _helpText, "", _dataType, new string[] { });
    }

    public ArgItem(string _longName, string _shortName, bool _required, string _helpText, string _defaultValue, ArgParse.ArgParseType _dataType)
    {
      init(_longName, _shortName, _required, _helpText, _defaultValue, _dataType, new string[] { });
    }

    public ArgItem(string _longName, string _shortName, bool _required, string _helpText, string _defaultValue, ArgParse.ArgParseType _dataType, string[] _validValues)
    {
      init(_longName, _shortName, _required, _helpText, _defaultValue, _dataType, _validValues);
    }

    private void init(string _longName, string _shortName, bool _required, string _helpText, string _defaultValue, ArgParse.ArgParseType _dataType, string[] _validValues)
    {
      longName = _longName;
      shortName = _shortName;
      helpText = _helpText;
      defaultValue = _defaultValue;
      required = _required;
      argType = _dataType;
      dataType = _dataType.ToType();
      validValues = _validValues;
    }
  }

  public static class ArgParseExtensions
  {
    public static Type ToType(this ArgParse.ArgParseType argType)
    {
      switch (argType)
      {
        case ArgParse.ArgParseType.DateTime:
          return DateTime.MaxValue.GetType();
        case ArgParse.ArgParseType.Int:
          return int.MaxValue.GetType();
        case ArgParse.ArgParseType.Long:
          return long.MaxValue.GetType();
        case ArgParse.ArgParseType.Boolean:
          return true.GetType();
        default:
          return "".GetType();
      }
    }
  }
}