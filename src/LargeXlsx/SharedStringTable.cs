/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020-2022 Salvatore ISAJA. All rights reserved.

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
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using SharpCompress.Writers.Zip;

namespace LargeXlsx
{
    internal class SharedStringTable
    {
        private readonly Dictionary<string, int> _stringItems;
        private int _nextStringId;

        public SharedStringTable()
        {
            _stringItems = new Dictionary<string, int>();
            _nextStringId = 0;
        }

        public int ResolveStringId(string s)
        {
            if (!_stringItems.TryGetValue(s, out var id))
            {
                id = _nextStringId++;
                _stringItems.Add(s, id);
            }
            return id;
        }

        public void Save(ZipWriter zipWriter)
        {
            using (var stream = zipWriter.WriteToStream("xl/sharedStrings.xml", new ZipWriterEntryOptions()))
            using (var streamWriter = new StreamWriter(stream, Encoding.UTF8))
            {
                streamWriter.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                                   + "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
                foreach (var si in _stringItems.OrderBy(s => s.Value))
                    streamWriter.Write("<si><t>{0}</t></si>", Util.EscapeXmlText(si.Key));
                streamWriter.Write("</sst>");
            }
        }
    }
}