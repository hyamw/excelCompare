﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelCompare
{
    internal class RowComparer
    {
        private int _rowIndex = -1;
        private int _leftRowIndex = -1;
        private int _rightRowIndex = -1;
        private bool _isDifferent = false;
        private bool[] _cellStatus = null;

        public int rowIndex
        {
            get
            {
                return _rowIndex;
            }
        }

        public int leftRowIndex
        {
            get
            {
                return _leftRowIndex;
            }
        }

        public int rightRowIndex
        {
            get
            {
                return _rightRowIndex;
            }
        }

        public bool isDifferent
        {
            get
            {
                return _isDifferent;
            }
        }

        public bool IsCellDifferent(int columnIndex)
        {
            return _cellStatus[columnIndex];
        }

        public RowComparer(int rowIndex, int leftRowIndex, int rightRowIndex, bool isDifferent)
        {
            _rowIndex = rowIndex;
            _leftRowIndex = leftRowIndex;
            _rightRowIndex = rightRowIndex;
            _isDifferent = isDifferent;
        }

        public void CompareCells(int columnCount, ExcelSheet left, ExcelSheet right)
        {
            _cellStatus = new bool[columnCount];
            ExcelRow leftRow = left.GetRowByRealIndex(_leftRowIndex);
            ExcelRow rightRow = right.GetRowByRealIndex(_rightRowIndex);

            if (leftRow == null || rightRow == null)
            {
                for (int i = 0; i < columnCount; i++)
                {
                    _cellStatus[i] = true;
                }
            }
            else
            {
                for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    string selfCell = leftRow.GetColumn(columnIndex);
                    string otherCell = rightRow.GetColumn(columnIndex);
                    if ((selfCell == null && otherCell != null) || (selfCell != null && otherCell == null))
                    {
                        _cellStatus[columnIndex] = true;
                    }
                    else
                    {
                        _cellStatus[columnIndex] = string.Compare(selfCell, otherCell) != 0;
                    }
                }
            }
        }
    }
}
