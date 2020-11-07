using System;
using System.Collections.Generic;

namespace ExcelSharp.NPOI
{
    public class SheetRangeColumnSelector
    {
        private SheetRange Range;

        public SheetRangeColumnSelector(SheetRange range)
        {
            Range = range;
        }

        public SheetRange this[int? startOffset, int? endOffset]
        {
            get
            {
                if (startOffset is null) startOffset = 0;
                if (endOffset is null) endOffset = Range.ColumnLengh - 1;

                return new SheetRange(Range.Sheet, (Range.Start.Row, Range.Start.Col + startOffset.Value), (Range.End.Row, Range.Start.Col + endOffset.Value));
            }
        }

        public SheetRangeGroup Select(Func<int, bool> indexSelector)
        {
            IEnumerable<SheetRange> select()
            {
                for (int index = Range.Start.Col; index <= Range.End.Col; index++)
                {
                    if (indexSelector(index))
                        yield return new SheetRange(Range.Sheet, (Range.Start.Row, Range.Start.Col + index), (Range.End.Row, Range.Start.Col + index));
                }
            }
            return new SheetRangeGroup(select());
        }

    }
}
