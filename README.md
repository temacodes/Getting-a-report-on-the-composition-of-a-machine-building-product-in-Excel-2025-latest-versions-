С# Output of the composition of a machine-building product in Excel 2025

using System.Data;
using System.Linq;
using System.Xml;
using System.Collections.Generic;
using Intermech.Interfaces;
using Intermech.Expert.Scenarios;
using Intermech.Interfaces.Document;
using Intermech.Kernel.Search;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;

public class Script
{
    public ICSharpScriptContext ScriptContext { get; private set; }

    public ScriptResult Execute(IUserSession session, ImDocumentData document, Int64[] objectIDs)
    {
        //Вставьте ваш код сценария здесь
        if (Debugger.IsAttached) Debugger.Break();

        int reltypeSP = MetaDataHelper.GetRelationTypeID(Intermech.SystemGUIDs.reltypeSP);
        int objtypeAssemblyUnit = MetaDataHelper.GetObjectTypeID(Intermech.SystemGUIDs.objtypeAssemblyUnit);

        //параметры запроса (список атрибутов)
        DBRecordSetParams paramSet = new DBRecordSetParams(null, new ColumnDescriptor[]
        {
            new ColumnDescriptor(ObligatoryObjectAttributes.CAPTION, AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.Name, SortOrders.NONE, 0),
            new ColumnDescriptor(ObligatoryObjectAttributes.F_OBJECT_ID, AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.Name, SortOrders.NONE, 0),
            new ColumnDescriptor("Обозначение", AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.Name, SortOrders.NONE, 0),
            new ColumnDescriptor("Наименование", AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.Name, SortOrders.NONE, 0),
            new ColumnDescriptor(ObligatoryObjectAttributes.F_OBJECT_TYPE, AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.Name, SortOrders.NONE, 0),
            new ColumnDescriptor("Код АМТО", AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.Name, SortOrders.NONE, 0),
            new ColumnDescriptor("Масса", AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.Name, SortOrders.NONE, 0),
            new ColumnDescriptor("Материал", AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.Name, SortOrders.NONE, 0),
            new ColumnDescriptor("Извещение", AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.Name, SortOrders.NONE, 0),
            new ColumnDescriptor("Количество", AttributeSourceTypes.Relation, ColumnContents.Text, ColumnNameMapping.Name, SortOrders.NONE, 0),
            new ColumnDescriptor(ObligatoryObjectAttributes.F_LC_STEP, AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.Name, SortOrders.NONE, 0)
        }, 0, null, QueryConsts.Default);

        //  запускаем Excel
        Application excel = new Microsoft.Office.Interop.Excel.ApplicationClass();
        try
        {
            excel.Visible = true;
            Workbook workbooks = excel.Workbooks.Add(Type.Missing);
            Worksheet sheet = (Worksheet)excel.ActiveSheet;

            // задаем стиль заголовка таблицы
            Range header = sheet.Range["A1:K1"];
            header.Font.ColorIndex = 5;
            header.Font.Underline = true;
            header.NumberFormat = "@";
            header.AutoFilter(1);
            sheet.Outline.SummaryRow = 0;

            // заполняем заголовок таблицы
            header.Cells.Value = new string[]
            {
                "Заголовок",
                "Идентификатор версии объекта",
                "Обозначение",
                "Наименование",
                "Тип объекта",
                "Код АМТО",
                "Масса",
                "Материал",
                "Извещение",
                "Количество",
                "Шаг жизненного цикла",
            };

            // закрепляем первую строку
            excel.ActiveWindow.SplitRow = 1;
            //excel.ActiveWindow.FreezePanes = true;

            // выключаем прорисовку
            excel.ScreenUpdating = false;

            // заполняем первую строку
            int row = 2;
            long objID = objectIDs.FirstOrDefault();
            IDBObject obj = session.GetObject(objID);

            IDBAttribute attr3 = obj.GetAttributeByName("Обозначение", false);
            IDBAttribute attr4 = obj.GetAttributeByName("Наименование", false);
            IDBAttribute attr6 = obj.GetAttributeByName("Код УПП МВМ", false);
            IDBAttribute attr7 = obj.GetAttributeByName("Масса", false);
            IDBAttribute attr8 = obj.GetAttributeByName("Материал", false);
            IDBAttribute attr9 = obj.GetAttributeByName("Извещение", false);

            sheet.Cells[row, 01] = obj.Caption;
            sheet.Cells[row, 02] = obj.ObjectID.ToString();
            if (attr3 != null) sheet.Cells[row, 03] = attr3.Value;
            if (attr4 != null) sheet.Cells[row, 04] = attr4.Value;
            sheet.Cells[row, 05] = MetaDataHelper.GetObjectName(obj.ObjectType);
            if (attr6 != null) sheet.Cells[row, 06] = attr6.Value;
            if (attr7 != null) sheet.Cells[row, 07] = attr7.Description;
            if (attr8 != null) sheet.Cells[row, 08] = attr8.Description;
            if (attr9 != null) sheet.Cells[row, 09] = attr9.Value;
            sheet.Cells[row, 11] = MetaDataHelper.GetLCStepName(obj.LCStep);

            // заполняем состав
            row = Recursion(session, objID, reltypeSP, objtypeAssemblyUnit, paramSet, sheet, row, 0);

            // выравниваем колонки по ширине
            sheet.Rows.EntireColumn.AutoFit();
        }
        finally
        {
            //включаем прорисовку
            excel.ScreenUpdating = true;
            excel.Quit();
        }
        return new ScriptResult(false, document);
    }

    // получить состав (применяемсть) объекта
    private System.Data.DataTable GetComposition(IUserSession session, long objectID, int relationType, DBRecordSetParams paramSet, int compositionMode = 0, bool recursive = false)
    {
        // получить данные по связи
        IDBRelationCollection relations = session.GetRelationCollection(relationType);
        // поиск по локальным типам связи
        relations.LocalTypesMode = false;
        // допустимые типы объектов
        relations.ChildObjectTypes = null;
        // получить состав объекта в соответствии с условиями
        return compositionMode == 0 ? relations.ConsistFrom(paramSet, objectID, recursive) : relations.EntersIn(paramSet, objectID, recursive);
    }

    //рекурсия для разворачивания состава
    public int Recursion(IUserSession session, long objectID, int relationType, int objectType, DBRecordSetParams paramSet, Worksheet ws, int row, int level)
    {
        System.Data.DataTable dt = GetComposition(session, objectID, relationType, paramSet, 0, false);
        level++;

        int startRow = row;

        foreach (DataRow r in dt.Rows)
        {
            row++;

            long objID = Convert.ToInt64(r[1]);
            int objType = Convert.ToInt32(r[4]);

            //формируем массив с данными
            object[] dataArray = r.ItemArray;

            int cols = dataArray.Length;

            if (cols > 0)
            {
                // определяем размер строки
                Range cellFirst = (Range)ws.Cells[row, 1];
                Range cellEnd = (Range)ws.Cells[row, cols];

                // заполняем строку данными
                ws.Range[cellFirst, cellEnd].Value2 = new object[]
                {
                    r[0].ToString().PadLeft(r[0].ToString().Length + level * 4, ' '),
                    objID,
                    r[2].ToString(),
                    r[3].ToString(),
                    MetaDataHelper.GetObjectName(objType),
                    r[5].ToString(),
                    r[6].ToString(),
                    r[7].ToString(),
                    r[8].ToString(),
                    r[9].ToString(),
                    MetaDataHelper.GetLCStepName( Convert.ToInt32(r[10]))
                };
            }

            // раскрываем сосав дочерней сборки
            if (objType == objectType)
            {
                row = Recursion(session, objID, relationType, objectType, paramSet, ws, row, level);
            }
        }

        // групируем записи в excel
        if (startRow != row && level < 8)
        {
            ws.Range[(startRow + 1) + ":" + row].Rows.Group();
        }
        return row;
