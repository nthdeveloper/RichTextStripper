using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;

namespace NthDeveloper.Rtf
{
    /// <summary>
    /// Rich Text Stripper
    /// </summary>
    /// <remarks>
    /// Original version is translated from Python by Chris Benard (https://chrisbenard.net/2014/08/20/extract-text-from-rtf-in-c-net/)
    /// Improved by nthdeveloper (Mustafa Kok)
    /// Non-unicode encoding support by font table processing
    /// Some performance improvements
    /// </remarks>
    public static class RichTextStripper
    {
        private struct StackEntry
        {
            public int NumberOfCharactersToSkip;
            public bool Ignorable;

            public StackEntry(int numberOfCharactersToSkip, bool ignorable)
            {
                NumberOfCharactersToSkip = numberOfCharactersToSkip;
                Ignorable = ignorable;
            }
        }

        private static readonly Regex RtfRegex = new Regex(@"\\([a-z]{1,32})(-?\d{1,10})?[ ]?|\\'([0-9a-f]{2})|\\([^a-z])|([{}])|[\r\n]+|(.)", RegexOptions.Singleline | RegexOptions.IgnoreCase);

        private static readonly HashSet<string> DestinationTags = new HashSet<string>
    {
        "aftncn","aftnsep","aftnsepc","annotation","atnauthor","atndate","atnicn","atnid",
        "atnparent","atnref","atntime","atrfend","atrfstart","author","background",
        "bkmkend","bkmkstart","blipuid","buptim","category","colorschememapping",
        "colortbl","comment","company","creatim","datafield","datastore","defchp","defpap",
        "do","doccomm","docvar","dptxbxtext","ebcend","ebcstart","factoidname","falt",
        "fchars","ffdeftext","ffentrymcr","ffexitmcr","ffformat","ffhelptext","ffl",
        "ffname","ffstattext","field","file","filetbl","fldinst","fldrslt","fldtype",
        "fname","fontemb","fontfile","fonttbl","footer","footerf","footerl","footerr",
        "footnote","formfield","ftncn","ftnsep","ftnsepc","g","generator","gridtbl",
        "header","headerf","headerl","headerr","hl","hlfr","hlinkbase","hlloc","hlsrc",
        "hsv","htmltag","info","keycode","keywords","latentstyles","lchars","levelnumbers",
        "leveltext","lfolevel","linkval","list","listlevel","listname","listoverride",
        "listoverridetable","listpicture","liststylename","listtable","listtext",
        "lsdlockedexcept","macc","maccPr","mailmerge","maln","malnScr","manager","margPr",
        "mbar","mbarPr","mbaseJc","mbegChr","mborderBox","mborderBoxPr","mbox","mboxPr",
        "mchr","mcount","mctrlPr","md","mdeg","mdegHide","mden","mdiff","mdPr","me",
        "mendChr","meqArr","meqArrPr","mf","mfName","mfPr","mfunc","mfuncPr","mgroupChr",
        "mgroupChrPr","mgrow","mhideBot","mhideLeft","mhideRight","mhideTop","mhtmltag",
        "mlim","mlimloc","mlimlow","mlimlowPr","mlimupp","mlimuppPr","mm","mmaddfieldname",
        "mmath","mmathPict","mmathPr","mmaxdist","mmc","mmcJc","mmconnectstr",
        "mmconnectstrdata","mmcPr","mmcs","mmdatasource","mmheadersource","mmmailsubject",
        "mmodso","mmodsofilter","mmodsofldmpdata","mmodsomappedname","mmodsoname",
        "mmodsorecipdata","mmodsosort","mmodsosrc","mmodsotable","mmodsoudl",
        "mmodsoudldata","mmodsouniquetag","mmPr","mmquery","mmr","mnary","mnaryPr",
        "mnoBreak","mnum","mobjDist","moMath","moMathPara","moMathParaPr","mopEmu",
        "mphant","mphantPr","mplcHide","mpos","mr","mrad","mradPr","mrPr","msepChr",
        "mshow","mshp","msPre","msPrePr","msSub","msSubPr","msSubSup","msSubSupPr","msSup",
        "msSupPr","mstrikeBLTR","mstrikeH","mstrikeTLBR","mstrikeV","msub","msubHide",
        "msup","msupHide","mtransp","mtype","mvertJc","mvfmf","mvfml","mvtof","mvtol",
        "mzeroAsc","mzeroDesc","mzeroWid","nesttableprops","nextfile","nonesttables",
        "objalias","objclass","objdata","object","objname","objsect","objtime","oldcprops",
        "oldpprops","oldsprops","oldtprops","oleclsid","operator","panose","password",
        "passwordhash","pgp","pgptbl","picprop","pict","pn","pnseclvl","pntext","pntxta",
        "pntxtb","printim","private","propname","protend","protstart","protusertbl","pxe",
        "result","revtbl","revtim","rsidtbl","rxe","shp","shpgrp","shpinst",
        "shppict","shprslt","shptxt","sn","sp","staticval","stylesheet","subject","sv",
        "svb","tc","template","themedata","title","txe","ud","upr","userprops",
        "wgrffmtfilter","windowcaption","writereservation","writereservhash","xe","xform",
        "xmlattrname","xmlattrvalue","xmlclose","xmlname","xmlnstbl",
        "xmlopen"
    };

        private static readonly Dictionary<string, string> SpecialCharacters = new Dictionary<string, string>
    {
        { "par", "\r\n" },
        { "sect", "\n\n" },
        { "page", "\n\n" },
        { "line", "\r\n" },
        { "tab", "\t" },
        { "emdash", "\u2014" },
        { "endash", "\u2013" },
        { "emspace", "\u2003" },
        { "enspace", "\u2002" },
        { "qmspace", "\u2005" },
        { "bullet", "\u2022" },
        { "lquote", "\u2018" },
        { "rquote", "\u2019" },
        { "ldblquote", "\u201C" },
        { "rdblquote", "\u201D" },
    };

        private static readonly string[] MultipleEntriesSeperator = new string[] { "}{" };

        /// <summary>
        /// Strip RTF Tags from RTF Text
        /// </summary>
        /// <param name="inputRtf">RTF formatted text</param>
        /// <returns>Plain text from RTF</returns>
        public static string StripRichTextFormat(string inputRtf)
        {
            if (String.IsNullOrEmpty(inputRtf))
                return null;

            //Parse font table
            List<System.Text.Encoding> _fontEntries = parseRTFFontTableAndGetEncodingList(inputRtf);

			byte[] _singleByteData = new byte[1]; // for Single Byte Character Sets (SBCS) (i.e. most of them)
			byte[] _doubleByteData = new byte[2]; // for Double Byte Character Sets (DBCS) (at least 4 of them)

            var _stack = new Stack<StackEntry>(128);
            bool _ignorable = false;              // Whether this group (and all inside it) are "ignorable".
            int _ucskip = 1;                      // Number of ASCII characters to skip after a unicode character.
            int _curskip = 0;                     // Number of ASCII characters left to skip
            var _outputTextList = new List<string>();    // Output buffer.            
			Match secondByteMatch = null;         // Only used for Double Byte Character Sets (DBCS)
			string secondByteHex = null;          // Only used for Double Byte Character Sets (DBCS)

            MatchCollection _matches = RtfRegex.Matches(inputRtf);

            if (_matches.Count == 0)
                return inputRtf;

            System.Text.Encoding _currentEncoding = System.Text.Encoding.Default;//Use current system's encoding by default

            for (int i = 0; i < _matches.Count; i++)
            {
                Match match = _matches[i];
                string word = match.Groups[1].Value;
                string arg = match.Groups[2].Value;
                string hex = match.Groups[3].Value;
                string character = match.Groups[4].Value;
                string brace = match.Groups[5].Value;
                string tchar = match.Groups[6].Value;

                if (!String.IsNullOrEmpty(brace))
                {
                    _curskip = 0;
                    if (brace == "{")
                    {
                        // Push state
                        _stack.Push(new StackEntry(_ucskip, _ignorable));
                    }
                    else if (brace == "}")
                    {
                        // Pop state
                        StackEntry entry = _stack.Pop();
                        _ucskip = entry.NumberOfCharactersToSkip;
                        _ignorable = entry.Ignorable;
                    }
                }
                else if (!String.IsNullOrEmpty(character)) // \x (not a letter)
                {
                    _curskip = 0;
                    if (character == "~")
                    {
                        if (!_ignorable)
                            _outputTextList.Add("\xA0");
                    }
                    else if ("{}\\".Contains(character))
                    {
                        if (!_ignorable)
                            _outputTextList.Add(character);
                    }
                    else if (character == "*")
                    {
                        _ignorable = true;
                    }
                }
                else if (!String.IsNullOrEmpty(word)) // \foo
                {
                    _curskip = 0;
                    if (DestinationTags.Contains(word))
                    {
                        _ignorable = true;
                    }
                    else if (word == "f")
                    {
                        if (!String.IsNullOrEmpty(arg))
                        {
                            int _fontNo = Int32.Parse(arg);
                            if (_fontNo < _fontEntries.Count)
                                _currentEncoding = _fontEntries[_fontNo];
                        }
                    }
                    else if (_ignorable)
                    {
                    }
                    else if (SpecialCharacters.ContainsKey(word))
                    {
                        _outputTextList.Add(SpecialCharacters[word]);
                    }
                    else if (word == "uc")
                    {
                        _ucskip = Int32.Parse(arg);
                    }
                    else if (word == "u")
                    {
                        int c = Int32.Parse(arg);
                        if (c < 0)
                            c += 0x10000;

                        _outputTextList.Add(Char.ConvertFromUtf32(c));
                        _curskip = _ucskip;
                    }
                }
                else if (!String.IsNullOrEmpty(hex)) // \'xx
                {
                    if (_curskip > 0)
                    {
                        _curskip -= 1;
                    }
                    else if (!_ignorable)
                    {
                        int c = Int32.Parse(hex, System.Globalization.NumberStyles.HexNumber);
						if (_currentEncoding.IsSingleByte || c < 128) // "\",  "{", and "}" are always escaped!
						{
							_singleByteData[0] = (byte)c;
							_outputTextList.Add(_currentEncoding.GetString(_singleByteData));
						}
						else
						{
							_doubleByteData[0] = (byte)c;

							secondByteMatch = _matches[++i]; // increment to get next match
							secondByteHex = secondByteMatch.Groups[3].Value; // should only be hex following a DBCS lead byte
							_doubleByteData[1] = byte.Parse(secondByteHex, System.Globalization.NumberStyles.HexNumber);

							_outputTextList.Add(_currentEncoding.GetString(_doubleByteData));
						}
                    }
                }
                else if (!String.IsNullOrEmpty(tchar))
                {
                    if (_curskip > 0)
                    {
                        _curskip -= 1;
                    }
                    else if (!_ignorable)
                    {
                        _outputTextList.Add(tchar);
                    }
                }
            }

            return String.Join(String.Empty, _outputTextList.ToArray());
        }

        /// <summary>
        /// Parses the font table in RTF content and returns corresponding encoding object list for font entries
        /// </summary>
        /// <param name="inputRtf">Full RTF content</param>
        /// <returns>System.Text.Encoding object for each font entry in the font table</returns>
        private static List<System.Text.Encoding> parseRTFFontTableAndGetEncodingList(string inputRtf)
        {
            List<System.Text.Encoding> _encodingList = null;

            try
            {
                int _fontTableStartPos = inputRtf.IndexOf("fonttbl");
                if (_fontTableStartPos == -1)
                    return new List<Encoding>(0);

                _fontTableStartPos += 8;

                int _fontTableEnd = inputRtf.IndexOf("}}", _fontTableStartPos);

                string _strFontTable = inputRtf.Substring(_fontTableStartPos, _fontTableEnd - _fontTableStartPos);
                string[] _fontEntries = _strFontTable.Split(MultipleEntriesSeperator, StringSplitOptions.RemoveEmptyEntries);

                System.Text.Encoding _textEncoding = null;

                _encodingList = new List<Encoding>(_fontEntries.Length);

                for (int i = 0; i < _fontEntries.Length; i++)
                {
                    _textEncoding = Encoding.Default;//Current system's default encoding

                    string _strEntry = _fontEntries[i];
                    int _charsetIndex = _strEntry.IndexOf("fcharset");

                    if (_charsetIndex > 0)
                    {
                        _charsetIndex += 8;

                        int _charsetCode = Int32.Parse(_strEntry.Substring(_charsetIndex, _strEntry.IndexOf(" ", _charsetIndex) - _charsetIndex));

                        _textEncoding = Encoding.GetEncoding(getCodePage(_charsetCode));
                    }

                    _encodingList.Add(_textEncoding);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }

            return _encodingList ?? new List<Encoding>(0);
        }

        /// <summary>
        /// Returns the code page value for the given character set value in RTF font definition
        /// </summary>
        /// <param name="charSet">Character set value in RTF font definition</param>
        /// <returns>Encoding code page value</returns>
        private static int getCodePage(int charSet)
        {
            switch (charSet)
            {
                case 0:
                    return 1252; // ANSI
                case 1:
                    return 0; // Default
                case 2:
                    return 0;//42; // Symbol (42 is not supported by .Net Encoding class and throws exception when trying to get with this code page)
                case 77:
                    return 10000; // Mac Roman
                case 78:
                    return 10001; // Mac Shift Jis
                case 79:
                    return 10003; // Mac Hangul
                case 80:
                    return 10008; // Mac GB2312
                case 81:
                    return 10002; // Mac Big5
                case 82:
                    return 0; // Mac Johab (old)
                case 83:
                    return 10005; // Mac Hebrew
                case 84:
                    return 10004; // Mac Arabic
                case 85:
                    return 10006; // Mac Greek
                case 86:
                    return 10081; // Mac Turkish
                case 87:
                    return 10021; // Mac Thai
                case 88:
                    return 10029; // Mac East Europe
                case 89:
                    return 10007; // Mac Russian
                case 128:
                    return 932; // Shift JIS
                case 129:
                    return 949; // Hangul
                case 130:
                    return 1361; // Johab
                case 134:
                    return 936; // GB2312
                case 136:
                    return 950; // Big5
                case 161:
                    return 1253; // Greek
                case 162:
                    return 1254; // Turkish
                case 163:
                    return 1258; // Vietnamese
                case 177:
                    return 1255; // Hebrew
                case 178:
                    return 1256; // Arabic
                case 179:
                    return 0; // Arabic Traditional (old)
                case 180:
                    return 0; // Arabic user (old)
                case 181:
                    return 0; // Hebrew user (old)
                case 186:
                    return 1257; // Baltic
                case 204:
                    return 1251; // Russian
                case 222:
                    return 874; // Thai
                case 238:
                    return 1250; // Eastern European
                case 254:
                    return 437; // PC 437
                case 255:
                    return 850; // OEM
            }

            return 0;
        }
    }
}
