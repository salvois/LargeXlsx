/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020-2023 Salvatore ISAJA. All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice,
this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
this list of conditions and the following disclaimer in the documentation
and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED THE COPYRIGHT HOLDER ``AS IS'' AND ANY EXPRESS
OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES
OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN
NO EVENT SHALL THE COPYRIGHT HOLDER BE LIABLE FOR ANY DIRECT,
INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/
namespace LargeXlsx
{
    public class XlsxSheetProtection
    {
        public string Password { get; }
        public bool Sheet { get; }
        public bool Objects { get; }
        public bool Scenarios { get; }
        public bool FormatCells { get; }
        public bool FormatColumns { get; }
        public bool FormatRows { get; }
        public bool InsertColumns { get; }
        public bool InsertRows { get; }
        public bool InsertHyperlinks { get; }
        public bool DeleteColumns { get; }
        public bool DeleteRows { get; }
        public bool SelectLockedCells { get; }
        public bool Sort { get; }
        public bool AutoFilter { get; }
        public bool PivotTables { get; }
        public bool SelectUnlockedCells { get; }

        public XlsxSheetProtection(
            string password,
            bool sheet = true,
            bool objects = true,
            bool scenarios = true,
            bool formatCells = true,
            bool formatColumns = true,
            bool formatRows = true,
            bool insertColumns = true,
            bool insertRows = true,
            bool insertHyperlinks = true,
            bool deleteColumns = true,
            bool deleteRows = true,
            bool selectLockedCells = false,
            bool sort = true,
            bool autoFilter = true,
            bool pivotTables = true,
            bool selectUnlockedCells = false)
        {
            Password = password;
            Sheet = sheet;
            Objects = objects;
            Scenarios = scenarios;
            FormatCells = formatCells;
            FormatColumns = formatColumns;
            FormatRows = formatRows;
            InsertColumns = insertColumns;
            InsertRows = insertRows;
            InsertHyperlinks = insertHyperlinks;
            DeleteColumns = deleteColumns;
            DeleteRows = deleteRows;
            SelectLockedCells = selectLockedCells;
            Sort = sort;
            AutoFilter = autoFilter;
            PivotTables = pivotTables;
            SelectUnlockedCells = selectUnlockedCells;
        }
    }
}