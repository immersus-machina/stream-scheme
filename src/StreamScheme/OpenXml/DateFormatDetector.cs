// Date format detection adapted from MiniExcel (Apache 2.0)
// which credits ExcelNumberFormat (MIT) by andersnm.
// https://github.com/MiniExcelFinancial/MiniExcel
// https://github.com/andersnm/ExcelNumberFormat

namespace StreamScheme.OpenXml;

internal interface IDateFormatDetector
{
    bool IsBuiltInDateFormat(int numberFormatId);
    bool IsDateFormatString(string? formatCode);
}

internal class DateFormatDetector : IDateFormatDetector
{
    public bool IsBuiltInDateFormat(int numberFormatId) => numberFormatId is
        14 or 15 or 16 or 17 or 18 or 19 or 20 or 21 or 22 or 45 or 46 or 47 or 58;

    public bool IsDateFormatString(string? formatCode)
    {
        if (string.IsNullOrEmpty(formatCode))
        {
            return false;
        }

        var classifier = new FormatTokenClassifier(formatCode.AsSpan());

        while (classifier.TryClassifyNextToken(out var tokenKind))
        {
            if (tokenKind == FormatTokenKind.SectionSeparator)
            {
                break;
            }

            if (tokenKind == FormatTokenKind.DatePart)
            {
                return true;
            }
        }

        return false;
    }

    private enum FormatTokenKind { Other, DatePart, SectionSeparator }

    // Classifies tokens in an Excel number format string (e.g. "yyyy-mm-dd", "#,##0.00")
    // according to the ECMA-376 SpreadsheetML number format rules.
    private ref struct FormatTokenClassifier(ReadOnlySpan<char> format)
    {
        private readonly ReadOnlySpan<char> _format = format;
        private int _position;

        public bool TryClassifyNextToken(out FormatTokenKind tokenKind)
        {
            if (_position >= _format.Length)
            {
                tokenKind = default;
                return false;
            }

            tokenKind = _format[_position] switch
            {
                '\\' or '*' or '_' => SkipEscapeSequence(),
                '"' => SkipQuotedString(),
                '[' => ClassifyBracketedExpression(),
                ';' => ClassifySectionSeparator(),
                _ => ClassifyKeywordOrLiteral(),
            };

            return true;
        }

        // \x, *x, _x — always two characters, never a date part
        private FormatTokenKind SkipEscapeSequence()
        {
            _position += Math.Min(2, _format.Length - _position);
            return FormatTokenKind.Other;
        }

        // "..." — quoted text, never a date part
        private FormatTokenKind SkipQuotedString()
        {
            _position++;
            while (_position < _format.Length && _format[_position] != '"')
            {
                _position++;
            }

            if (_position < _format.Length)
            {
                _position++;
            }

            return FormatTokenKind.Other;
        }

        // [...] — date part if elapsed time ([h], [m], [s]), otherwise other
        private FormatTokenKind ClassifyBracketedExpression()
        {
            var isDatePart = _position + 1 < _format.Length &&
                char.ToLowerInvariant(_format[_position + 1]) is 'h' or 'm' or 's';

            _position++;
            while (_position < _format.Length && _format[_position] != ']')
            {
                _position++;
            }

            if (_position < _format.Length)
            {
                _position++;
            }

            return isDatePart ? FormatTokenKind.DatePart : FormatTokenKind.Other;
        }

        private FormatTokenKind ClassifySectionSeparator()
        {
            _position++;
            return FormatTokenKind.SectionSeparator;
        }

        // Multi-char keywords (am/pm, general, e+), repeated date chars (yyyy, mm),
        // and single literal characters
        private FormatTokenKind ClassifyKeywordOrLiteral()
        {
            if (TryConsumeExact("am/pm") || TryConsumeExact("a/p"))
            {
                return FormatTokenKind.DatePart;
            }

            if (TryConsumeIgnoreCase("e+") || TryConsumeIgnoreCase("e-"))
            {
                return FormatTokenKind.Other;
            }

            if (TryConsumeIgnoreCase("general"))
            {
                return FormatTokenKind.Other;
            }

            var lower = char.ToLowerInvariant(_format[_position]);
            if (lower is 'y' or 'm' or 'd' or 'h' or 's' or 'g')
            {
                while (_position < _format.Length && char.ToLowerInvariant(_format[_position]) == lower)
                {
                    _position++;
                }

                return FormatTokenKind.DatePart;
            }

            _position++;
            return FormatTokenKind.Other;
        }

        private bool TryConsumeExact(string expected)
        {
            if (_position + expected.Length > _format.Length)
            {
                return false;
            }

            for (var i = 0; i < expected.Length; i++)
            {
                if (_format[_position + i] != expected[i])
                {
                    return false;
                }
            }

            _position += expected.Length;
            return true;
        }

        private bool TryConsumeIgnoreCase(string expected)
        {
            if (_position + expected.Length > _format.Length)
            {
                return false;
            }

            for (var i = 0; i < expected.Length; i++)
            {
                if (char.ToLowerInvariant(_format[_position + i]) != char.ToLowerInvariant(expected[i]))
                {
                    return false;
                }
            }

            _position += expected.Length;
            return true;
        }
    }
}
