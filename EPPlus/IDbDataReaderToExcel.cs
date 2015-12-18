    public static class ExportQueryToExcel_Static
    {
        public static void Export<T>(string filename, IEnumerable<T> coll)
        {
            Export(filename, coll, str => str);
        }

        public static void Export<T>(string filename, IEnumerable<T> coll, Func<string, string> headerTranslator)
        {
            var p = new ExcelPackage();



            var sheetname = "query";
            var sheet = p.Workbook.Worksheets.Add(sheetname);
            int row = 1;
            var t = typeof(T);

            var props = t.GetProperties();
            int count = props.Count();
            for (int i = 0; i < count; i++)
            {
                var header = props[i].Name;
                header = headerTranslator(header);
                sheet.Cells[1, i + 1].Value = header;
            }

            foreach (var item in coll)
            {
                row++;


                for (int i = 0; i < count; i++)
                {
                    var v = props[i].GetValue(item);
                    sheet.Cells[row, i + 1].Value = v;
                }
            }

            for (int i = 0; i < count; i++)
            {
                var ft = props[i].PropertyType;
                if (ft == typeof(DateTime))
                {
                    sheet.Cells[1, i + 1, row, i + 1].Style.Numberformat.Format = @"dd\.mm\.yyyy";
                }
            }

            File.Delete(filename);
            p.SaveAs(new FileInfo(filename));
            p.Dispose();
        }

        public static void Export(string filename, IDataReader reader)
        {
            using (var p = new ExcelPackage())
            {
                var sheet = p.Workbook.Worksheets.Add("Query");
                var row = 1;
                var count = reader.FieldCount;
                for (int i = 0; i < count; i++)
                {
                    var header = reader.GetName(i);
                    sheet.Cells[1, i + 1].Value = header;
                }

                while (reader.Read())
                {
                    row++;
                    var objects = new object[count];
                    reader.GetValues(objects);
                    for (int i = 0; i < count; i++)
                    {
                        sheet.Cells[row, i + 1].Value = objects[i];
                    }
                }

                for (int i = 0; i < count; i++)
                {
                    var ft = reader.GetFieldType(i);
                    if (ft == typeof(DateTime))
                    {
                        sheet.Cells[1, i + 1, row, i + 1].Style.Numberformat.Format = @"dd\.mm\.yyyy";
                    }
                }
                File.Delete(filename);
                p.SaveAs(new FileInfo(filename));
            }

        }
    }
