using System.ComponentModel;

namespace Fox.Npoi.Test.Model
{
    public class Depart
    {
        [Description("部门id")]
        public int DepId { get; set; }

        [Description("部门名称")]
        public string DepName { get; set; }
    }
}