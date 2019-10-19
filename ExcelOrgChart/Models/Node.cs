using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelOrgChart.Models
{
    public class Node
    {
        public int Id { get; set; }
        public string Title { get; set; }
        public int ParentID { get; set; }
        public virtual Node Parent { get; set; }
        public virtual List<Node> Children { get; set; }
    }
}
