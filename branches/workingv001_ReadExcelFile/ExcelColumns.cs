using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReadExcelFile
{
    class ExcelColumns
    {
        private String[] column;
        private int endColumnNumber;

        public ExcelColumns(int numberOfColumns)
        {
            column = new string[numberOfColumns];
            endColumnNumber = numberOfColumns;
            getReferencedColumns();
        }

        private void getReferencedColumns()
        {
            int numberOfInteraction = (int)endColumnNumber / 26;
            int endColumnLastInteraction = endColumnNumber - (numberOfInteraction * 26);
            int posArray = 0;
            char lastCharLastInteraction = 'A';

            if (endColumnLastInteraction == 0)
                endColumnLastInteraction = 26;
            else if (endColumnLastInteraction > 0)
                numberOfInteraction++;

            endColumnLastInteraction += 64;
            lastCharLastInteraction = (char)endColumnLastInteraction;

            if (numberOfInteraction == 1)
            {
                for (char letterPos = 'A'; letterPos <= lastCharLastInteraction; letterPos++)
                    column[posArray++] = letterPos.ToString();
            }
            else
            {
                for (int interaction = 1; interaction <= numberOfInteraction; interaction++)
                {
                    for (char letterPos = 'A'; letterPos <= 'Z'; letterPos++)
                    {
                        if (posArray < 26)
                            column[posArray++] = letterPos.ToString();
                        else if (posArray >= 26)
                        {
                            if (numberOfInteraction > interaction)
                            {
                                for (char letterPos2 = 'A'; letterPos2 <= 'Z'; letterPos2++)
                                    column[posArray++] = letterPos.ToString() + letterPos2.ToString();

                                interaction++;
                            }
                            else if (numberOfInteraction == interaction)
                            {
                                for (char letterPos2 = 'A'; letterPos2 <= lastCharLastInteraction; letterPos2++)
                                    column[posArray++] = letterPos.ToString() + letterPos2.ToString();

                                break;
                            }
                        }

                    }// end for (char letterPos = 'A'; letterPos <= 'Z'; letterPos++)
                }// end for (int interaction = 1; interaction <= numberOfInteraction; interaction++)
            }

        }// end private void getReferencedColumns()

        public String[] Columns
        {
            get { return column; }
        }
    }
}
