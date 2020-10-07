using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocxToPdfMVVM.Services
{
    static public class ItemsSet
    {
        static public ObservableCollection<string> Items { get; set; }
        static public string PathFile { get; set; }
    }
}
