enum MatchMode {
    Equals = 0,
    Contains = 1,
}

interface MatchFunc {
    Index: number
    Value: any
    Mode: MatchMode
}

/**
 * @customfunction
 * @param [args] 
 * @returns 
 */
export function test(args: any) {
    return typeof(args)
}

/**
 * @customfunction
 * @param range 
 * @param value 
 * @param index 
 * @returns 
 */
export function vlookup(range: any[][], value: any, index: number) {
    for(const row of range) {
        if(row[0] == value) return row[index-1]
    }
    throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable, "no result found")
}

/**
 * 高级vlookup。
 * @customfunction
 * @param range 查找范围
 * @param resultIndex 如果找到匹配行，要返回结果的列
 * @param defaultValue 如果找不到匹配行，返回的默认值
 * @param [lookupArgs] 匹配条件。每个匹配条件由3个值构成：匹配列、匹配值、匹配模式（0：完全相等；1：字符串包含）。
 * @returns 匹配到的内容或默认值
 */
export function super_vlookup(range: any[][], resultIndex: number, defaultValue: any, lookupArgs: any[]) {
    const args = unpackLookupArgs(lookupArgs)
    if(args.length == 0) return 
    for(let r = 0; r < range.length; r++) {
        const row = range[r]
        const matches = args.filter(([index, value, mode]) => {
            return isMatch(row[index-1], value, mode)
        })
        if(matches.length == args.length) {
            return row[resultIndex-1]
        }
    }
    return defaultValue
}

// console.log(super_vlookup([[1,2,3],[4,5,6],[7,8,9]], 3, 0, 1, 4, 0, 2, 5, 1))

function isMatch(value: any, match: any, mode: MatchMode) {
    if(mode == MatchMode.Equals) {
        return value == match
    }
    if(mode == MatchMode.Contains) {
        return String(value).includes(String(match))
    }
    return false
}

function unpackLookupArgs(lookupArgs) {
    const args = []
    let last = lookupArgs
    while (last.length >= 3) {
        args.push(last.slice(0, 3))
        last = last.slice(3)
    }
    return args
}