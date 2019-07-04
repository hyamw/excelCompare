using DiffMatchPatch;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelCompare
{
    internal class DiffAnalyzer
    {
        public const string LINE_SPLITTER = "=[$$]=";
        protected class Content
        {
            private string openEntry = null;
            private List<string> lines = new List<string>();
            private bool openDiff = false;

            public void AddLine(string line)
            {
                lines.Add(line);
            }

            public void EnqueueLine(string text, bool different)
            {
                if (openEntry == null)
                {
                    openEntry = text;
                }
                else
                {
                    openEntry += text;
                }

                openDiff |= different;
            }

            public void FinishLine(string text)
            {
                if (openEntry != null)
                {
                    if ( text != null )
                    {
                        lines.Add(openEntry + text);
                    }
                    else
                    {
                        lines.Add(openEntry);
                    }
                }
                else
                {
                    if (text != null)
                    {
                        lines.Add(text);
                    }
                    else
                    {
                        Debug.Assert(false, "Unexpected!");
                    }
                }
                openEntry = null;
                openDiff = false;
            }

            public bool FinishLine()
            {
                bool result = openDiff;
                if (openEntry != null)
                {
                    lines.Add(openEntry);
                }
                else
                {
                    Debug.Assert(false, "Unexpected!");
                }
                openEntry = null;
                openDiff = false;
                return result;
            }

            public string GetFullContent()
            {
                StringBuilder builder = new StringBuilder();
                for (int i = 0; i < lines.Count; i++)
                {
                    builder.AppendLine(lines[i]);
                }
                return builder.ToString();
            }

            public int LineCount
            {
                get
                {
                    return lines.Count;
                }
            }
        }
        private static readonly string[] SPLIT_SEPERATORS = new string[] { LINE_SPLITTER };
        private Content left = new Content();
        private Content right = new Content();
        private int totalLine = 0;

        public string GetLeftContent()
        {
            return left.GetFullContent();
        }

        public string GetRightContent()
        {
            return right.GetFullContent();
        }

        public bool Analysis(List<Diff> diffs, ExcelSheet leftSheet, ExcelSheet rightSheet)
        {
            bool isDifferent = false;
            for (int diffIndex = 0; diffIndex < diffs.Count; diffIndex++)
            {
                Diff diff = diffs[diffIndex];
                string[] lines = diff.text.Split(SPLIT_SEPERATORS, StringSplitOptions.None);
                for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++)
                {
                    string line = lines[lineIndex];

                    if (lineIndex > 0)
                    {
                        if (diff.operation == Operation.EQUAL)
                        {
                            bool leftDiff = left.FinishLine();
                            bool rightDiff = right.FinishLine();
                            totalLine++;
                            leftSheet.SetRowTarget(left.LineCount - 1, totalLine - 1, leftDiff);
                            rightSheet.SetRowTarget(right.LineCount - 1, totalLine - 1, rightDiff);
                        }
                        else if (diff.operation == Operation.DELETE)
                        {
                            left.FinishLine();
                            totalLine++;
                            leftSheet.SetRowTarget(left.LineCount - 1, totalLine - 1, true);
                        }
                        else if (diff.operation == Operation.INSERT)
                        {
                            right.FinishLine();
                            totalLine++;
                            rightSheet.SetRowTarget(right.LineCount - 1, totalLine - 1, true);
                        }
                    }

                    if (diff.operation == Operation.EQUAL)
                    {
                        left.EnqueueLine(line, false);
                        right.EnqueueLine(line, false);
                    }
                    else if (diff.operation == Operation.DELETE)
                    {
                        left.EnqueueLine(line, true);
                        isDifferent = true;
                    }
                    else if (diff.operation == Operation.INSERT)
                    {
                        right.EnqueueLine(line, true);
                        isDifferent = true;
                    }
                }
            }

            int columnCount = Math.Max(leftSheet.columnCount, rightSheet.columnCount);
            leftSheet.SetDiffViewSize(totalLine, columnCount);
            rightSheet.SetDiffViewSize(totalLine, columnCount);

            return isDifferent;
        }
    }
}
