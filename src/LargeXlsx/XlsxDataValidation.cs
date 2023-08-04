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

using System;
using System.Collections.Generic;
using System.Linq;

namespace LargeXlsx
{
    public class XlsxDataValidation : IEquatable<XlsxDataValidation>
    {
        public bool AllowBlank { get; }
        public string Error { get; }
        public string ErrorTitle { get; }
        public ErrorStyle? ErrorStyleValue { get; }
        public Operator? OperatorValue { get; }
        public string Prompt { get; }
        public string PromptTitle { get; }
        public bool ShowDropDown { get; }
        public bool ShowErrorMessage { get; }
        public bool ShowInputMessage { get; }
        public ValidationType? ValidationTypeValue { get; }
        public string Formula1 { get; }
        public string Formula2 { get; }

        public enum ErrorStyle
        {
            Information,
            Stop,
            Warning
        }

        public enum Operator
        {
            Between,
            Equal,
            GreaterThan,
            GreaterThanOrEqual,
            LessThan,
            LessThanOrEqual,
            NotBetween,
            NotEqual
        }

        public enum ValidationType
        {
            Custom,
            Date,
            Decimal,
            List,
            None,
            TextLength,
            Time,
            Whole
        }

        public XlsxDataValidation(
            bool allowBlank = false,
            string error = null,
            string errorTitle = null,
            ErrorStyle? errorStyle = null,
            Operator? operatorType = null,
            string prompt = null,
            string promptTitle = null,
            bool showDropDown = false,
            bool showErrorMessage = false,
            bool showInputMessage = false,
            ValidationType? validationType = null,
            string formula1 = null,
            string formula2 = null)
        {
            AllowBlank = allowBlank;
            Error = error;
            ErrorTitle = errorTitle;
            ErrorStyleValue = errorStyle;
            OperatorValue = operatorType;
            Prompt = prompt;
            PromptTitle = promptTitle;
            ShowDropDown = showDropDown;
            ShowErrorMessage = showErrorMessage;
            ShowInputMessage = showInputMessage;
            ValidationTypeValue = validationType;
            Formula1 = formula1;
            Formula2 = formula2;
        }

        public static XlsxDataValidation List(
            IEnumerable<string> choices,
            bool allowBlank = false,
            string error = null,
            string errorTitle = null,
            ErrorStyle? errorStyle = null,
            string prompt = null,
            string promptTitle = null,
            bool showDropDown = false,
            bool showErrorMessage = false,
            bool showInputMessage = false)
        {
            return new XlsxDataValidation(allowBlank, error, errorTitle, errorStyle, null, prompt, promptTitle,
                showDropDown, showErrorMessage, showInputMessage,
                ValidationType.List, '"' + string.Join(",", choices.Select(c => c.Replace("\"", "\"\""))) + '"');
        }

        #region Equality members

        public bool Equals(XlsxDataValidation other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return AllowBlank == other.AllowBlank && Error == other.Error && ErrorTitle == other.ErrorTitle
                   && ErrorStyleValue == other.ErrorStyleValue && OperatorValue == other.OperatorValue
                   && Prompt == other.Prompt && PromptTitle == other.PromptTitle && ShowDropDown == other.ShowDropDown
                   && ShowErrorMessage == other.ShowErrorMessage && ShowInputMessage == other.ShowInputMessage
                   && ValidationTypeValue == other.ValidationTypeValue
                   && Formula1 == other.Formula1 && Formula2 == other.Formula2;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((XlsxDataValidation) obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = AllowBlank.GetHashCode();
                hashCode = (hashCode * 397) ^ (Error != null ? Error.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (ErrorTitle != null ? ErrorTitle.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ ErrorStyleValue.GetHashCode();
                hashCode = (hashCode * 397) ^ OperatorValue.GetHashCode();
                hashCode = (hashCode * 397) ^ (Prompt != null ? Prompt.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (PromptTitle != null ? PromptTitle.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ ShowDropDown.GetHashCode();
                hashCode = (hashCode * 397) ^ ShowErrorMessage.GetHashCode();
                hashCode = (hashCode * 397) ^ ShowInputMessage.GetHashCode();
                hashCode = (hashCode * 397) ^ ValidationTypeValue.GetHashCode();
                hashCode = (hashCode * 397) ^ (Formula1 != null ? Formula1.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (Formula2 != null ? Formula2.GetHashCode() : 0);
                return hashCode;
            }
        }

        public static bool operator ==(XlsxDataValidation left, XlsxDataValidation right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(XlsxDataValidation left, XlsxDataValidation right)
        {
            return !Equals(left, right);
        }

        #endregion
    }
}