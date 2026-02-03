using System.Collections.Generic;

namespace ExcelAddIn1.Core.Models
{
    public class Person
    {
        public string Name { get; set; }
        public int? Age { get; set; }
        public string City { get; set; }

        public List<string> ValidationErrors { get; } = new List<string>();
    }
}

