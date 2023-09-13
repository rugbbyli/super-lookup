using ExcelDna.Integration;

namespace super_vlookup;

public static class SuperVLookup
{
    [ExcelFunction]
    public static double sum(double a, double b) => a + b;
    
    /// <summary>
    /// super vlookup
    /// </summary>
    /// <param name="range">查找范围</param>
    /// <param name="resultIndex">如果找到匹配行，要返回结果的列</param>
    /// <param name="dftValue">如果找不到匹配行，返回的默认值</param>
    /// <param name="lookupArgs">匹配条件。每个匹配条件由3个值构成：匹配列、匹配值、匹配模式（0：完全相等；1：字符串包含）。</param>
    /// <returns>匹配到的内容或默认值</returns>
    [ExcelFunction(Description = "super vlookup")]
    public static object supervlookup(object[][] range, int resultIndex, object dftValue, object[] lookupArgs)
    {
        var fns = new List<Func<object[], bool>>();
        for (int i = 0; i < lookupArgs.Length; i+= 3)
        {
            if (i + 2 >= lookupArgs.Length) break;
            var col = (int)lookupArgs[i];
            var vl = lookupArgs[i + 1];
            var matchMode = (MatchMode)lookupArgs[i + 2];
            fns.Add(row => matchMode.IsMatch(row[col], vl));
        }

        for (int row = 0; row < range.Length; row++)
        {
            
            var fire = fns.All(f => f(range[row]));
            if (fire) return range[row][resultIndex];
        }

        return dftValue;
    }

    private static bool IsMatch(this MatchMode matchMode, object a, object b)
    {
        if (matchMode == MatchMode.Equal) return a == b;
        if (matchMode == MatchMode.Contains) return a is string sa && b is string sb && sa.Contains(sb);
        return false;
    }
    
    private enum MatchMode
    {
        Equal = 0,
        Contains = 1,
    }
}