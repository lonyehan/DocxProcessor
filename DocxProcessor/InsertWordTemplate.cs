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
        public void InsertTableRow(Table table, TableRow tableRow)
        {                       
           //add new row to table, after last row in table
           table.Descendants<TableRow>().Last().InsertBeforeSelf(tableRow);         
        }

        public void InsertTableRow(Table table,TableRow tableRow, int count)
        {                        
            for (int i = 0; i <= count; i++)
            {
                //clone our "reference row"
                var rowToInsert = (TableRow)tableRow.Clone();                                                
                //add new row to table, after last row in table
                table.Descendants<TableRow>().Last().InsertBeforeSelf(rowToInsert);
            }
        }
    }
}