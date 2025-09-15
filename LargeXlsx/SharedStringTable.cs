/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020-2025 Salvatore ISAJA. All rights reserved.

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
using System.Linq;
using System.Threading.Tasks;

namespace LargeXlsx;

internal class SharedStringTable(bool skipInvalidCharacters)
{
    private readonly Dictionary<string, int> _stringItems = new();
    private int _nextStringId = 0;

    public int ResolveStringId(string s)
    {
        if (!_stringItems.TryGetValue(s, out var id))
        {
            id = _nextStringId++;
            _stringItems.Add(s, id);
        }
        return id;
    }

    public async Task Save(IZipWriter zipWriter, CustomWriter customWriter)
    {
#if NETCOREAPP2_1_OR_GREATER
        await using var stream = zipWriter.CreateEntry("xl/sharedStrings.xml");
#else
        using var stream = zipWriter.CreateEntry("xl/sharedStrings.xml");
#endif
        customWriter.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"u8
                            + "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"u8);
        foreach (var si in _stringItems.OrderBy(s => s.Value))
        {
            // <si><t xml:space="preserve">{0}</t></si>
            await customWriter
                .Append("<si><t"u8)
                .AddSpacePreserveIfNeeded(si.Key)
                .Append(">"u8)
                .AppendEscapedXmlText(si.Key, skipInvalidCharacters)
                .Append("</t></si>\n"u8)
                .TryFlushToAsync(stream).ConfigureAwait(false);
        }
        await customWriter.Append("</sst>\n"u8).FlushToAsync(stream).ConfigureAwait(false);
    }
}