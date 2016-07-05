using System;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using System.Reflection;
using System.ComponentModel;
using System.IO;

namespace OpenXML.Excel
{
    public class ExcelReader
    {
        /// <summary>
        /// Excel文档包
        /// </summary>
        private SpreadsheetDocument document;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="fileName">Excel文件路径</param>
        /// <param name="editable">是否可以编辑</param>
        public ExcelReader(string fileName, bool editable)
        {
            //打开Excel文档包
            document = SpreadsheetDocument.Open(fileName, editable);
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="fileName">Excel文件路径</param>
        public ExcelReader(string fileName)
            : this(fileName, false)
        {

        }

        /// <summary>
        /// 将指定工作表转化为DataTable
        /// </summary>
        /// <param name="workbookPart">Excel文档</param>
        /// <param name="sheet">Excel工作表</param>
        /// <returns></returns>
        private DataTable ConvertSheetToTable(SpreadsheetDocument document, Sheet sheet)
        {
            if (document == null || sheet == null)
                return null;

            //创建数据表
            DataTable table = new DataTable(sheet.Name);

            //获取工作表
            WorkbookPart workbookPart = document.WorkbookPart;

            //获取工作表
            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

            //获取该工作表下的行
            Row[] rows = worksheetPart.Worksheet.Descendants<Row>().ToArray<Row>();

            //遍历每一行数据
            foreach (Row row in rows)
            {
                DataRow tableRow = table.NewRow();
                foreach (Cell cell in row)
                {
                    //设置列名称
                    string columnName = Regex.Match(cell.CellReference.Value, "[a-zA-Z]+").Value;
                    if (!table.Columns.Contains(columnName))
                        table.Columns.Add(columnName);

                    //设置行数据
                    tableRow[columnName] = GetRowCellValue(cell, workbookPart);
                }
                table.Rows.Add(tableRow);
            }

            //返回数据表
            return table;
        }

        /// <summary>
        /// 返回一个DataSet
        /// </summary>
        /// <returns></returns>
        public DataSet AsDataSet()
        {
            //初始化DataSet
            DataSet dataset = new DataSet();

            if (document == null)
                return null;

            //获取全部Sheet
            Sheet[] sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().ToArray<Sheet>();

            //当工作簿中无工作表时返回null
            //实际上每个工作簿最少有1个工作表
            if (sheets.Length <= 0)
                return null;

            //遍历每个工作表
            foreach (Sheet sheet in sheets)
            {
                DataTable table = ConvertSheetToTable(document, sheet);
                dataset.Tables.Add(table);
            }

            //返回数据
            return dataset;
        }

        /// <summary>
        /// 获取行中单元格数值
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="workbookPart">工作表</param>
        private string GetRowCellValue(Cell cell, WorkbookPart workbookPart)
        {
            //初始化默认数值
            string cellValue = string.Empty;

            if (cell.ChildElements.Count <= 0)
                return cellValue;

            //获取单元格InnertText
            string cellInnerText = cell.CellValue.InnerText;
            cellValue = cellInnerText;

            //获取共享数据表格
            SharedStringTable sharedTable = workbookPart.SharedStringTablePart.SharedStringTable;

            //根据数值类型对数据进行处理
            if (cell.DataType != null)
            {
                switch (cell.DataType.Value)
                {
                    case CellValues.SharedString:
                        int cellIndex = int.Parse(cellInnerText);
                        cellValue = sharedTable.ChildElements[cellIndex].InnerText;
                        break;
                    case CellValues.Boolean:
                        cellValue = cellInnerText == "1" ? "TRUE" : "FALSE";
                        break;
                    case CellValues.Date:
                        cellValue = Convert.ToDateTime(cellInnerText).ToString();
                        break;
                    case CellValues.Number:
                        cellValue = Convert.ToDecimal(cellInnerText).ToString();
                        break;
                    default:
                        cellValue = cellInnerText;
                        break;
                }
            }

            //返回数值
            return cellValue;
        }

        /// <summary>
        /// 关闭文档包引用
        /// </summary>
        public void Close()
        {
            if (document == null)
                return;

            //关闭文档
            document.Close();
            //释放对象
            document = null;
        }

        /// <summary>
        /// 返回指定表指定索引所在行
        /// </summary>
        /// <param name="sheetName">数据表名</param>
        /// <param name="index">行索引</param>
        /// <returns></returns>
        public DataRow GetRowAt(string sheetName, int index)
        {
            DataTableCollection tables = AsDataSet().Tables;
            foreach (DataTable table in tables)
            {
                if (table.TableName == sheetName)
                {
                    //索引越界则返回为空
                    if (index >= table.Rows.Count)
                        return null;

                    return table.Rows[index];
                }
            }

            //返回数据
            return null;
        }

        /// <summary>
        /// 返回指定表指定索引所在行
        /// </summary>
        /// <param name="sheetName">数据表名</param>
        /// <param name="index">列索引</param>
        /// <returns></returns>
        public DataColumn GetColumnAt(string sheetName, int index)
        {
            DataTableCollection tables = AsDataSet().Tables;
            foreach (DataTable table in tables)
            {
                if (table.TableName == sheetName)
                {
                    //索引越界则返回为空
                    if (index >= table.Columns.Count)
                        return null;

                    return table.Columns[index];
                }
            }

            //返回数据
            return null;
        }

        /// <summary>
        /// 返回指定表指定行指定列的值
        /// </summary>
        /// <param name="sheetName">数据表名</param>
        /// <param name="index">列索引</param>
        /// <returns></returns>
        public string GetValueAt(string sheetName, int row, int column)
        {
            DataTableCollection tables = AsDataSet().Tables;
            foreach (DataTable table in tables)
            {
                if (table.TableName == sheetName)
                {
                    //索引越界则返回为空
                    if (row >= table.Rows.Count || column >= table.Columns.Count)
                        return null;

                    return table.Rows[row][column].ToString();
                }
            }

            //返回数据
            return null;
        }

        /// <summary>
        /// 将指定工作表转换为实体列表
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="sheetName">工作表名称</param>
        /// <returns></returns>
        public List<T> AsList<T>(string sheetName)
        {
            //初始化对象列表
            List<T> list = new List<T>();

            //获取指定名称的数据表
            DataTable table = AsDataTable(sheetName);

            if (table == null)
                return null;

            for (int i = 1; i < table.Rows.Count; i++)
            {
                //创建泛型对象实例
                T target = CreateInstace<T>();

                for (int j = 0; j < table.Columns.Count; j++)
                {
                    if (table.Rows[i][j] != null)
                    {
                        //获取属性名称
                        string propertyName = table.Rows[0][j].ToString();
                        //获取属性值
                        object propertyValue = table.Rows[i][j];
                        //为对象赋属性值
                        SetTargetProperty(target, propertyName, propertyValue);
                    }
                }

                //添加至列表
                list.Add(target);
            }

            //返回数据
            return list;
        }

        /// <summary>
        /// 将指定工作表转化为数据表
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        /// <returns></returns>
        public DataTable AsDataTable(string sheetName)
        {
            DataTableCollection tables = AsDataSet().Tables;
            foreach (DataTable table in tables)
            {
                if (table.TableName == sheetName)
                    return table;
            }

            //返回数据
            return null;
        }

        /// <summary>
        /// 创建泛型对象实例
        /// </summary>
        /// <typeparam name="T">泛型参数</typeparam>
        /// <returns></returns>
        private T CreateInstace<T>()
        {
            Type t = typeof(T);
            ConstructorInfo ct = t.GetConstructor(System.Type.EmptyTypes);
            return (T)ct.Invoke(null);
        }

        /// <summary>
        /// 设置目标对象的属性值
        /// </summary>
        /// <param name="target">目标对象</param>
        /// <param name="propertyName">属性名称</param>
        /// <param name="propertyValue">属性值</param>
        private void SetTargetProperty(object target, string propertyName, object propertyValue)
        {
            //获取类型
            Type type = target.GetType();

            //获取属性集合
            PropertyInfo[] propertys = type.GetProperties();

            //遍历属性并对属性赋值
            foreach (PropertyInfo property in propertys)
            {
                //在这里增加对DBNull的支持
                if (property.Name == propertyName)
                    property.SetValue(target, ChangeType(propertyValue, property.PropertyType), null);
            }
        }

        /// <summary>
        /// 根据单元格地址获取单元格对应的数值
        /// </summary>
        /// <param name="sheetName">表名称</param>
        /// <param name="cellAddress">单元格地址</param>
        /// <returns>当地址不存在时返回null否则返回指定单元格地址对应数值</returns>
        public string GetCellValueByAddress(string sheetName, string cellAddress)
        {
            //获取指定工作表
            Sheet sheet = GetSheetByName(sheetName);

            if (sheet == null)
                return null;

            //遍历行和列
            Row[] rows = sheet.Descendants<Row>().ToArray<Row>();
            foreach (Row row in rows)
            {
                foreach (Cell cell in row)
                {
                       if (cell.CellReference == cellAddress){
                        return GetRowCellValue(cell, document.WorkbookPart);
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// 设置指定单元格地址对应的数值
        /// </summary>
        /// <param name="sheetName">表名称</param>
        /// <param name="cellAddress">单元格地址</param>
        /// <param name="cellValue">更新后的数值</param>
        /// <param name="cellType">单元格类型</param>
        public void SetCellValueByAddress(string sheetName, string cellAddress, string cellValue,CellValues cellType)
        {
            Worksheet worksheet = GetWorkSheetByName(sheetName);

            if (worksheet == null)
                return;

            //获取当前当前单元格
            Cell cell = GetCellByAddress(worksheet, cellAddress);

            //设置当前单元格数据
            cell.CellValue = new CellValue(cellValue.ToString());
            cell.DataType.Value = cellType;
        }

        /// <summary>
        /// 删除指定工作表
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        public void DeleteWorkSheet(string sheetName)
        {
            //获取指定表
            Sheet sheet = GetSheetByName(sheetName);

            //当指定表不存在时返回
            if (sheet == null)
                return;

            //获取当前工作表
            WorksheetPart worksheetPart = (WorksheetPart)(document.WorkbookPart.GetPartById(sheet.Id));

            //删除当前工作表
            document.WorkbookPart.DeletePart(worksheetPart);

        }

        /// <summary>
        /// 保存指定工作表
        /// </summary>
        /// <param name="sheetName">表名称</param>
        public void SaveWorkSheet(string sheetName)
        {
            //获取指定工作表
            Sheet sheet = GetSheetByName(sheetName);

            //当工作表不存在时返回
            if (sheet == null)
                return;

            //获取当前工作表
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);
            Worksheet worksheet = worksheetPart.Worksheet;

            //保存当前工作表
            worksheet.Save();
        }

        /// <summary>
        /// 保存整个工作簿
        /// </summary>
        public void Save()
        {
            document.WorkbookPart.Workbook.Save();
        }


        /// <summary>
        /// 根据工作表名返回Worksheet
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        private Worksheet GetWorkSheetByName(string sheetName)
        {
            //获取指定工作表
            Sheet sheet = GetSheetByName(sheetName);

            //当工作表不存在时返回
            if (sheet == null)
                return null;

            //获取当前工作表
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);
            Worksheet worksheet = worksheetPart.Worksheet;

            return worksheet;
        }

        /// <summary>
        /// 根据单元格地址返回一个单元格引用
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="cellAddress">单元格地址</param>
        /// <returns>如果单元格引用存在则直接返回否则先创建引用然后返回</returns>
        private Cell GetCellByAddress(Worksheet sheet, string cellAddress)
        {
            //获取SheetData
            SheetData sheetData = sheet.GetFirstChild<SheetData>();

            //解析行号和列号
            uint rowNumber = uint.Parse(Regex.Match(cellAddress, "[0-9]+").Value);
            string colName = Regex.Match(cellAddress, "[a-zA-Z]+").Value;

            //构造CellReference
            string cellReference = colName + rowNumber.ToString();

            //初始化当前单元格
            Cell theCell = null;

            //查找指定行数对应的行如不存在则需要创建新行
            var theRow = sheetData.Elements<Row>().Where(item => item.RowIndex.Value == rowNumber).FirstOrDefault();
            if (theRow == null){
                theRow = new Row();
                theRow.RowIndex = rowNumber;
                sheetData.Append(theRow);
            }

            //查找指定单元格如不存在则需要插入新单元格
            var refCell = theRow.Elements<Cell>().Where(item => item.CellReference.Value == cellReference).FirstOrDefault();
            if (refCell != null){
                theCell = refCell;
            }else{
                theCell = new Cell();
                theCell.CellReference = cellReference;
                theRow.InsertBefore(theCell, refCell);
            }

            return theCell;
        }

        /// <summary>
        /// 根据表名称返回一张工作表
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        private Sheet GetSheetByName(string sheetName)
        {
            //获取全部Sheet
            Sheet[] sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().ToArray<Sheet>();

            //获取指定Sheet
            var reval = sheets.Where(item => item.Name == sheetName);

            if(reval.Count() <= 0)
                return null;

            return reval.ElementAt<Sheet>(0);
        }

        /// <summary>
        /// 增强后的ChangeType
        /// </summary>
        /// <param name="value">数值</param>
        /// <param name="type">类型</param>
        /// <returns></returns>
        private object ChangeType(object value, Type type)
        {
            if (type == null)
                throw new ArgumentNullException("Type is null");

            if (type.IsGenericType &&
                type.GetGenericTypeDefinition().Equals(typeof(Nullable)))
            {
                if (value == null)
                    return null;

                NullableConverter converter = new NullableConverter(type);
                type = converter.UnderlyingType;
            }

            return Convert.ChangeType(value, type);
        }

    }
}
