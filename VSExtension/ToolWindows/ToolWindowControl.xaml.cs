﻿using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using System;
using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;

namespace VSExtension.ToolWindows
{
  public partial class ToolWindowControl : UserControl
  {
    public ToolWindowControl()
    {
      this.InitializeComponent();
    }

    [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions", Justification = "Sample code")]
    private int NumOfObjects(CodeElements codeElements)
    {
      int numOfObj = 0;
      CodeElement codeElement;

      ThreadHelper.ThrowIfNotOnUIThread();

      for (int i = 1; i <= codeElements.Count; ++i)
      {
        codeElement = codeElements.Item(i);
        if (codeElement.Kind == vsCMElement.vsCMElementFunction)
        {
          ++numOfObj;
        }
      }
      return numOfObj;
    }

    private void AddRows(int count)
    {
      for (int i = 0; i < count; ++i)
      {
        InfoTable.Items.Add(new { Function = "0", Lines = "0", LinesWithoutComments = "0", KeyWords = "0" });
      }
    }
    private void RemoveRows(int count)
    {
      for (int i = 0; i < count; ++i)
      {
        InfoTable.Items.RemoveAt(0);
      }
    }

    private int NumOfLines(TextPoint startPoint, TextPoint endPoint)
    {

      ThreadHelper.ThrowIfNotOnUIThread();

      return endPoint.Line - startPoint.Line + 1;
    }

    private int Strings(string input)
    {
      int res = 0;
      for (int i = 0; i < input.Length; ++i)
      {
        if (input[i] == '\n') { ++res; }
      }
      return res;
    }

    private void RemoveSingleLineComment(ref string func)
    {
      Regex reg = new Regex(@"//(.*\\\n)*.*\n");
      func = reg.Replace(func, "\n");
    }

    private void RemoveMultipleLinesComment(ref string func)
    {
      int count;
      string pattern = @"/\*(.*?\n)*?.*?\*/";
      Match match = Regex.Match(func, pattern);
      while (match.Success)
      {
        count = Strings(match.Value);
        string newString = new string('\n', count);
        Regex reg = new Regex(pattern);
        func = reg.Replace(func, newString, 1);
        match = match.NextMatch();
      }
    }

    private void RemoveSingleSymbolQuote(ref string func)
    {
      string newString = "";
      Match match = Regex.Match(func, @"('.*?')|('.*?\n)");
      if (match.Success)
      {
        if (match.Value[match.Value.Length - 1] != '\n')
        {
          newString += " lavax ";
        }
        else
        {
          newString += " lavax\n";
        }
        Regex reg = new Regex(@"('.*?')|('.*?\n)");
        func = reg.Replace(func, newString, 1);
      }
    }

    private void RemoveMultipleSymbolsQuote(ref string func)
    {
      string newString = "";
      int count = 0;
      Match match = Regex.Match(func, @"(""(.*?\\\r\n)*?.*?"")|(""(.*?\\\r\n)+?.*?\n)|(""(.*?\\\r\n)*?.*?\n)");
      if (match.Success)
      {
        count = Strings(match.Value);
        for (int i = 0; i < count; ++i)
        {
          newString += "lavax\n";
        }
        if (match.Value[match.Value.Length - 1] != '\n')
        {
          newString += " lavax ";
        }
        Regex reg = new Regex(@"(""(.*?\\\r\n)*?.*?"")|(""(.*?\\\r\n)+?.*?\n)|(""(.*?\\\r\n)*?.*?\n)");
        func = reg.Replace(func, newString, 1);
      }
    }

    private void ConvertToLf(ref string func)
    {
      Regex cr = new Regex(@"\r");
      func = cr.Replace(func, "\n");

      Regex lflf = new Regex(@"\n\n");
      func = lflf.Replace(func, "\n");
    }

    private int NumOfCodeLines(TextPoint begin, TextPoint end, ref string func)
    {
      int delta = 0;

      ThreadHelper.ThrowIfNotOnUIThread();

      func = begin.CreateEditPoint().GetLines(begin.Line, end.Line + 1);
      func += '\n';
      Regex reg = new Regex(@"\\'");
      func = reg.Replace(func, "");

      ConvertToLf(ref func);

      for (int i = 0; i < func.Length - 2; ++i)
      {
        if (func[i] == '/' && func[i + 1] == '/')
        {
          RemoveSingleLineComment(ref func);
          MessageBox.Show(func);
        }
        else if (func[i] == '/' && func[i + 1] == '*')
        {
          RemoveMultipleLinesComment(ref func);
        }
        else if (func[i] == '"')
        {
          RemoveMultipleSymbolsQuote(ref func);
        }
        else if (func[i] == '\'')
        {
          RemoveSingleSymbolQuote(ref func);
        }
      }

      bool emptyStr = true;
      int duplicate = 0;
      for (int i = 0; i < func.Length; ++i)
      {
        if (func[i] == '\n')
        {
          if (emptyStr)
          {
            ++delta;
          }
          if (duplicate > 1)
          {
            delta -= duplicate - 1;
          }
          emptyStr = true;
          duplicate = 0;
          continue;
        }
        if (func[i] != '\t' && func[i] != '\r' && func[i] != ' ')
        {
          emptyStr = false;
        }
        if (i > func.Length - 7)
        {
          continue;
        }
        if (func[i] == 'l' && func[i + 1] == 'a' && func[i + 2] == 'v' && func[i + 3] == '.' && func[i + 4] == 'a' && func[i + 5] == 'x')
        {
          ++duplicate;
        }
      }
      return end.Line - begin.Line + 1 - delta;
    }

    private int NumOfKeyWords(string func)
    {
      if (func == null)
      {
        return 0;
      }
      int res = 0;
      for (int i = 0; i < KeyWords.Length; ++i)
      {
        string pattern = "";
        pattern += KeyWords[i] + "\\W";
        res += Regex.Matches(func, @pattern).Count;
        pattern = "(_";
        pattern += KeyWords[i] + "\\W" + ")|(\\w" + KeyWords[i] + "\\W)";
        res -= Regex.Matches(func, @pattern).Count;
      }
      return res;
    }

    private void Update(object sender, RoutedEventArgs e)
    {
      int lines;
      int linesWithoutComments;
      int keyWords;
      int numOfObjects;
      string name;
      string func = null;

      ThreadHelper.ThrowIfNotOnUIThread();

      DTE2 dte = VSExtensionPackage.GetGlobalService(typeof(DTE)) as DTE2;
      FileCodeModel fileCodeModel = dte.ActiveDocument.ProjectItem.FileCodeModel;
      CodeElements codeElements = fileCodeModel.CodeElements;
      numOfObjects = NumOfObjects(codeElements);
      if (InfoTable.Items.Count < numOfObjects)
      {
        AddRows(numOfObjects - InfoTable.Items.Count);
      }
      else if (InfoTable.Items.Count > numOfObjects)
      {
        RemoveRows(InfoTable.Items.Count - numOfObjects);
      }
      for (int i = 1, j = 0; i <= codeElements.Count; ++i)
      {
        CodeElement codeElement = codeElements.Item(i);
        if (codeElement.Kind == vsCMElement.vsCMElementFunction)
        {
          func = null;
          CodeFunction codeFunc = codeElements.Item(i) as CodeFunction;
          TextPoint startPoint = codeFunc.GetStartPoint(vsCMPart.vsCMPartHeader);
          TextPoint endPoint = codeFunc.GetEndPoint(vsCMPart.vsCMPartBodyWithDelimiter);
          name = codeFunc.FullName;
          lines = NumOfLines(startPoint, endPoint);
          linesWithoutComments = NumOfCodeLines(startPoint, endPoint, ref func);
          keyWords = NumOfKeyWords(func);
          InfoTable.Items[j] = new { Function = name, Lines = lines.ToString(), LinesWithoutComments = linesWithoutComments.ToString(), KeyWords = keyWords.ToString() };
          ++j;
        }
      }
    }

    private void Resize(object sender, RoutedEventArgs e)
    {
      InfoTable.Columns[0].Width = MyToolWindow.ActualWidth / 4;
      InfoTable.Columns[1].Width = MyToolWindow.ActualWidth / 4;
      InfoTable.Columns[2].Width = MyToolWindow.ActualWidth / 4;
      InfoTable.Columns[3].Width = MyToolWindow.ActualWidth / 4;
      if (MyToolWindow.ActualHeight - 70 > 0)
      {
        InfoTable.Height = MyToolWindow.ActualHeight - 70;
      }
    }

    string[] KeyWords =
    {
      "alignas",
      "alignof",
      "and",
      "and_eq",
      "asm",
      "atomic_cancel",
      "atomic_commit",
      "atomic_noexcept",
      "auto",
      "bitand",
      "bitor",
      "bool",
      "break",
      "case",
      "catch",
      "char",
      "char8_t",
      "char16_t",
      "char32_t",
      "class",
      "compl",
      "concept",
      "const",
      "consteval",
      "constexpr",
      "constinit",
      "const_cast",
      "continue",
      "co_await",
      "co_return",
      "co_yield",
      "decltype",
      "default",
      "delete",
      "do",
      "double",
      "dynamic_cast",
      "else",
      "enum",
      "explicit",
      "export",
      "false",
      "float",
      "for",
      "friend",
      "goto",
      "if",
      "inline",
      "int",
      "long",
      "mutable",
      "namespace",
      "new",
      "noexcept",
      "not",
      "not_eq",
      "nullptr",
      "operator",
      "or",
      "or_eq",
      "private",
      "protected",
      "public",
      "reflexpr",
      "register",
      "reinterpret_cast",
      "requires",
      "restrict",
      "return",
      "short",
      "signed",
      "sizeof",
      "static",
      "static_assert",
      "static_cast",
      "struct",
      "switch",
      "synchronized",
      "template",
      "this",
      "thread_local",
      "throw",
      "true",
      "try",
      "typedef",
      "typeid",
      "typename",
      "typeof",
      "typeof_unqual",
      "union",
      "unsigned",
      "using",
      "virtual",
      "void",
      "volatile",
      "wchar_t",
      "while",
      "xor",
      "xor_eq",
      "final",
      "override",
      "transaction_safe",
      "transaction_safe_dynamic",
      "import",
      "module",
      "_Alignas",
      "_Alignof",
      "_Atomic",
      "_BitInt",
      "_Bool",
      "_Complex",
      "_Decimal128",
      "_Decimal32",
      "_Decimal64",
      "_Generic",
      "_Imaginary",
      "_Noreturn",
      "_Static_assert",
      "_Thread_local"
    };
  }
}