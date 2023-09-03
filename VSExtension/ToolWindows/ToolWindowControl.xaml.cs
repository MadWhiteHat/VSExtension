using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
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
    [SuppressMessage("StyleCop.CSharp.NamingRules", "SA1300:ElementMustBeginWithUpperCaseLetter", Justification = "Default event handler naming pattern")]
    private int NumOfObjects(CodeElements codeElements)
    {
      int numOfObj = 0;
      CodeElement codeElement = null;

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

    private void AddRow(int count)
    {
      for (int i = 0; i < count; ++i)
      {
        InfoTable.Items.Add(new { Function = "0", Lines = "0", LinesWithoutComments = "0", KeyWords = "0" });
      }
    }
    private void RemoveRow(int count)
    {
      for (int i = 0; i < count; ++i)
      {
        InfoTable.Items.RemoveAt(0);
      }
    }

    private int NumOfLines(CodeFunction codeFunction)
    {

      ThreadHelper.ThrowIfNotOnUIThread();

      return codeFunction.GetEndPoint(vsCMPart.vsCMPartBodyWithDelimiter).Line - codeFunction.GetStartPoint(vsCMPart.vsCMPartHeader).Line + 1;
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

    private int SingleCommentTample(ref string func)
    {
      string newString = "";
      int count = 0;
      Match match = Regex.Match(func, @"//(.*\\\r\n)*.*\n");
      if (match.Success)
      {
        count = Strings(match.Value);
        for (int i = 0; i < count; ++i)
        {
          newString += " lav.ax\n";
          Regex reg = new Regex(@"//(.*\\\r\n)*.*\n");
          func = reg.Replace(func, newString, 1);
        }
      }
      return count;
    }

    private int MultipleCommentTample(ref string func)
    {
      string newString = "";
      int count = 0;
      Match match = Regex.Match(func, @"/\*(.*?\n)*?.*?\*/");
      if (match.Success)
      {
        count = Strings(match.Value);
        for (int i = 0; i < count; ++i)
        {
          newString += " lav.ax\n";
        }
        newString += " lav.ax ";
        Regex reg = new Regex(@"/\*(.*?\n)*?.*?\*/");
        func = reg.Replace(func, newString, 1);
      }
      return count + 1;
    }

    private void SingleQuoteTample(ref string func)
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

    private void MultipleQuoteTample(ref string func)
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

    private int NumOfCodeLines(TextPoint begin, TextPoint end, ref string func)
    {
      int delta = 0;

      ThreadHelper.ThrowIfNotOnUIThread();

      func = begin.CreateEditPoint().GetLines(begin.Line, end.Line + 1);
      func += '\n';
      Regex reg = new Regex(@"\\""");
      func = reg.Replace(func, "");
      reg = new Regex(@"\\'");
      func = reg.Replace(func, "");
      for (int i = 0; i < func.Length - 2; ++i)
      {
        if (func[i] == '/' && func[i + 1] == '/')
        {
          delta += SingleCommentTample(ref func);
        }
        else if (func[i] == '/' && func[i + 1] == '*')
        {
          delta += MultipleCommentTample(ref func);
        }
        else if (func[i] == '"')
        {
          MultipleQuoteTample(ref func);
        }
        else if (func[i] == '\'')
        {
          SingleQuoteTample(ref func);
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
            delta -= (duplicate - 1);
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
        pattern = null;
        pattern = "(_";
        pattern += KeyWords[i] + "\\W" + ")|(\\w" + KeyWords[i] + "\\W)";
        res -= Regex.Matches(func, @pattern).Count;
      }
      return res;
    }

    private void Update(object sender, RoutedEventArgs e)
    {
      int lines = 0;
      int linesWithoutComments = 0;
      int keyWords = 0;
      int numOfObjects = 0;
      string name = null;
      string func = null;

      ThreadHelper.ThrowIfNotOnUIThread();

      DTE2 dte = VSExtensionPackage.GetGlobalService(typeof(DTE)) as DTE2;
      FileCodeModel fileCodeModel = dte.ActiveDocument.ProjectItem.FileCodeModel;
      CodeElements codeElements = fileCodeModel.CodeElements;
      numOfObjects = NumOfObjects(codeElements);
      if (InfoTable.Items.Count < numOfObjects)
      {
        AddRow(numOfObjects - InfoTable.Items.Count);
      }
      else if (InfoTable.Items.Count > numOfObjects)
      {
        RemoveRow(InfoTable.Items.Count - numOfObjects);
      }
      for (int i = 1, j = 0; i <= codeElements.Count; ++i)
      {
        CodeElement codeElement = codeElements.Item(i);
        if (codeElement.Kind == vsCMElement.vsCMElementFunction)
        {
          func = null;
          var codeElem = codeElements.Item(i) as CodeFunction;
          name = codeElem.FullName;
          lines = NumOfLines(codeElem);
          linesWithoutComments = NumOfCodeLines(codeElem.GetStartPoint(vsCMPart.vsCMPartHeader), codeElem.GetEndPoint(vsCMPart.vsCMPartBodyWithDelimiter), ref func);
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