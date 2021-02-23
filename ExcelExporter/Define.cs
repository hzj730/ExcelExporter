using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelExporter
{
    public class Define
    {
#if DEBUG
        public static readonly string ExcelPath = "../../excel/";
#else
        public static readonly string ExcelPath = "../excel/";
#endif
        public static readonly string OutputLuaPath = ExcelPath + "lua/";
        public static readonly int StartColumIndex = 1; // 表内容默认从第二例起有效，第一列特殊标记用
        public static readonly int StartRowIndex = 5; // 表内容第六行起，前5行为保留用特殊作用

        // lua 结构字段定义导出模版
        public static string TABLE_FIELD_TEMPLATE =
@"-- 该文件自动生成，请不要随修改
_G.TableDefaultNum = 0
_G.TableDefaultStr = """"
_G.TableDefaultTable = {{}}

-- 所有lua表的字段定义
_G.TableDefine = {{
{0}}}

return _G.TableDefine
";

        // 单个表字段申明模版，配合TABLE_FIELD_TEMPLATE使用
        public static string SINGLE_TABLE_FIELD_TEMPLATE =
@"{0} = {{
    meta = {{
{1}    }},
    file = '{2}',
}},
";

        // 转出Lua时替换一些统一值，节约内存
        public static readonly string DefaultLuaNum = "TableDefaultNum";
        public static readonly string DefaultLuaStr = "TableDefaultStr";
        public static readonly string DefaultLuaTable = "TableDefaultTable";

        public static readonly int DefaultJsonNum = 0;
        public static readonly string DefaultJsonStr = string.Empty;
        public static readonly string DefaultJsonTable = "[]";

        public static string TABLE_DATA_TEMP =
@"-- 该文件自动生成，请不要随修改
local t =
{{
{0}}}
_G.TableData.{1} = t";
    }
}
