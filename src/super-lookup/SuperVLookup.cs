using ExcelDna.Integration;
using ExcelDna.Registration;

namespace super_vlookup;

public class AddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        ExcelRegistration
            .GetExcelFunctions()
            .ProcessParamsRegistrations()
            .RegisterFunctions();
    }

    public void AutoClose()
    {
    }
}

public static class SuperVLookup
{
    [ExcelFunction(Description = "sum2333")]
    public static double myfunc(params object[] pms) => pms.Length;

    [ExcelFunction(Description = "type")]
    public static string mytype(object a) => a.GetType().ToString();
    
    /// <summary>
    /// 类似vlookup，但根据多个条件查找（比如查找同时满足『第一列值是AAA、第二列值包含bbb』的行的第三列的内容）。
    /// </summary>
    /// <param name="range">查找范围</param>
    /// <param name="resultIndex">如果找到匹配行，要返回结果的列</param>
    /// <param name="dftValue">如果找不到匹配行，返回的默认值</param>
    /// <param name="lookupArgs">匹配条件。每个匹配条件由3个值构成：匹配列、匹配值、匹配模式（0：完全相等；1：字符串包含）。</param>
    /// <returns>匹配到的内容或默认值</returns>
    [ExcelFunction(Description = "super vlookup")]
    public static object supervlookup(object[,] range, int resultIndex, object dftValue, params object[] lookupArgs)
    {
        var fns = new List<Func<int, bool>>();
        for (int i = 0; i < lookupArgs.Length; i+= 3)
        {
            if (i + 2 >= lookupArgs.Length) break;
            var col = Convert.ToInt32(lookupArgs[i]) - 1;
            var vl = lookupArgs[i + 1];
            var matchMode = (MatchMode)Convert.ToInt32(lookupArgs[i + 2]);
            fns.Add((row) => matchMode.IsMatch(range[row, col], vl));
        }

        for (int row = 0; row < range.GetLength(0); row++)
        {
            var fire = fns.All(f => f(row));
            if (fire) return range[row, resultIndex-1];
        }
        
        return dftValue;
    }

    private static bool IsMatch(this MatchMode matchMode, object a, object b)
    {
        if (matchMode == MatchMode.Equal) return a.Equals(b);
        if (matchMode == MatchMode.Contains) return a is string sa && b is string sb && sa.Contains(sb);
        return false;
    }
    
    private enum MatchMode
    {
        Equal = 0,
        Contains = 1,
    }
}