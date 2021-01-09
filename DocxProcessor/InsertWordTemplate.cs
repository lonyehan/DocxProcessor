using OpenXmlPowerTools;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using System.Xml.Linq;
using System;
using System.Text.RegularExpressions;
using System.Text;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Web;
using System.Drawing;
using DocumentFormat.OpenXml;
using System.Linq;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using System.Data;

namespace DocxProcessor
{
    public class InsertWordTemplate
    {
        public void InsertTableRow(TableRow TargetRowPosition, TableRow tableRow)
        {
            //add new row to table, after last row in table
            TargetRowPosition.InsertBeforeSelf(tableRow);
        }
        public void InsertTableRow(TableRow TargetRowPosition, TableRow tableRow, int count)
        {
            for(int i = 0; i< count; i++)
            {
                InsertTableRow(TargetRowPosition, tableRow);
            }
        }

        public void InsertTableRow(Table table, TableRow tableRow, int rowIndex)
        {                           
           //add new row to table, after last row in table
           table.Descendants<TableRow>().ElementAt(rowIndex).InsertAfterSelf(tableRow);         
        }

        public void InsertTableRow(Table table,TableRow tableRow, int rowIndex, int count)
        {                        
            for (int i = 0; i <= count; i++)
            {                
                var rowToInsert = (TableRow)tableRow.Clone();

                //add new row to table, after last row in table
                InsertTableRow(table, rowToInsert, rowIndex);
            }
        }
    }
}