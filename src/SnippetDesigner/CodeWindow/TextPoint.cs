// Copyright (C) Microsoft Corporation. All rights reserved.

namespace Microsoft.SnippetDesigner
{
    /// <summary>
    /// Represents a point in the TextBuffer
    /// </summary>
    public class TextPoint
    {
        public TextPoint(int line, int index)
        {
            Line = line;
            Index = index;
        }

        public TextPoint()
        {
            Line = 0;
            Index = 0;
        }

        /// <summary>
        /// The index into the line
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        ///  The line in the buffer
        /// </summary>
        public int Line { get; set; }

        // Overloading '<' operator:
        public static bool operator <(TextPoint point1, TextPoint point2) => point1.Index < point2.Index && point1.Line <= point2.Line;

        // Overloading '>' operator:
        public static bool operator >(TextPoint point1, TextPoint point2) => point1.Index > point2.Index && point1.Line >= point2.Line;

        // Override the Object.Equals(object o) method:
        public override bool Equals(object obj)
        {
            if (obj == null)
            {
                return false;
            }
            try
            {
                if (obj is not TextPoint point2)
                {
                    return false;
                }

                return Index == point2.Index && Line == point2.Line;
            }
            catch
            {
                return false;
            }
        }

        public override int GetHashCode() => base.GetHashCode();

        /// <summary>
        /// Set one text point equal to the other
        /// </summary>
        /// <param name="point"></param>
        public void SetEqualTo(TextPoint point)
        {
            Index = point.Index;
            Line = point.Line;
        }
    }
}
