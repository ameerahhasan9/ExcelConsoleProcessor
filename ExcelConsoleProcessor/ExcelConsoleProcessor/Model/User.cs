using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConsoleProcessor
{
    internal class User
    {
        [Key]
        public int UniqueID { get; set; }
        public string? FullName { get; set; }
        public string? Email { get; set; }
        public int IsActive { get; set; }
    }
}
