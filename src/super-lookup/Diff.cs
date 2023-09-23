using ExcelDna.Integration;
using ExcelDna.Registration;

namespace super_vlookup;

public static class Diff
{
    [ExcelFunction]
    public static object myrange(int row, int col)
    {
        return Enumerable.Range(1, row).ToArray2D(col, i => Enumerable.Range((i - 1) * col + 1, col).Cast<object>().ToArray());
    }
    
    [ExcelFunction]
    public static object[] except(object[] set1, object[] set2)
    {
        return set1.Except(set2).ToArray();
    }

    [ExcelFunction]
    public static object diff_by_sum(
        // [ExcelArgument(AllowReference = true)] ExcelReference set1, 
        // int key1,
        // int by1,
        // [ExcelArgument(AllowReference = true)] ExcelReference set2,
        // int key2,
        // int by2
        object[] key1, object[] by1,
        object[] key2, object[] by2
        )
    {
        var list1 = key1.Select((k, i) => (key: k, value: by1[i])).GroupBy(k => k.key).ToDictionary(i => i.Key, i => i.Sum(v => v.value is double d ? d : 0));
        var list2 = key2.Select((k, i) => (key: k, value: by2[i])).GroupBy(k => k.key).ToDictionary(i => i.Key, i => i.Sum(v => v.value is double d ? d : 0));
        return list1.Select(i => (i.Key, value: list2.GetValueOrDefault(i.Key, 0) - i.Value))
            .Where(i => Math.Abs(i.value) > 0.01).ToArray2D(2, i => new[] { i.Key, i.value });
        // var diff = list1.Where(i => Math.Abs(i.Value - list2.GetValueOrDefault(i.Key, 0)) > 0.01).ToArray();
        // return string.Join(",",
        //     diff.Select(i => i.Key).Append(diff.Sum(i => i.Value)));
    }

    private static object[,] ToArray2D<T>(this IEnumerable<T> input, int colSize, Func<T, object[]> splitor)
    {
        var data = input.ToArray();
        var arr = new object [data.Length, colSize];
        for (int i = 0; i < data.Length; i++)
        {
            var inr = splitor(data[i]);
            for (int j = 0; j < inr.Length; j++)
            {
                arr[i, j] = inr[j];
            }
        }

        return arr;
    }
}