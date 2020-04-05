﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelExporter
{
    public class Define
    {
        public static readonly string ExcelPath = "../../excel/";
        public static readonly string OutputLuaPath = ExcelPath + "lua/";

        public static string TABLE_FIELD_TEMP =
@"-- 该文件自动生成，请不要随修改
local fields = {}
fields.DefaultNum = 0
fields.DefaultStr = """"
fields.DefaultTable = {}

_G.TableField = fields
return _G.TableField
";
        // 转出Lua时替换一些统一值，节约内存
        public static readonly string DefaultNum = "TableField.DefaultNum";
        public static readonly string DefaultStr = "TableField.DefaultStr";
        public static readonly string DefaultTable = "TableField.DefaultTable";

        public static string TABLE_DATA_TEMP =
@"-- 该文件自动生成，请不要随修改
local t =
{{
{0}
}}
_G.TableData.{1} = t";
    }
}
