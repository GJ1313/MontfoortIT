using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MontfoortIT.Office.Excel
{
    public class SheetCollection:IEnumerable<Sheet>
    {
        private readonly Application _application;
        private List<Sheet> _sheets = new List<Sheet>();

        public int Count
        {
            get
            {
                return _sheets.Count;
            }
        }

        internal SheetCollection(Application application)
        {
            _application = application;
        }

        public Sheet Create(string name)
        {
            Sheet sheet = new Sheet(_application, name, _sheets.Count + 1);
            _sheets.Add(sheet);
            return sheet;
        }

        public IEnumerator<Sheet> GetEnumerator()
        {
            return _sheets.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public Sheet this[int index]
        {
            get { return _sheets[index]; }
        }

        internal void Read(IEnumerable<System.IO.Packaging.PackagePart> sheets, int emptyRowsToIgnore = -1)
        {
            int id=1;
            foreach (var packageSheet in sheets)
            {
                Sheet sheet = new Sheet(_application, packageSheet.Uri.ToString(),id++);
                sheet.Read(packageSheet, emptyRowsToIgnore);
                _sheets.Add(sheet);
            }
        }
    }
}
