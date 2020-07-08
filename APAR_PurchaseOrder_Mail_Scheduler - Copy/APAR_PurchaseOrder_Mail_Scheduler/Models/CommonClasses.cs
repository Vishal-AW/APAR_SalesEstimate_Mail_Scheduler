using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APAR_PurchaseOrder_Mail_Scheduler.Models
{

    public class PurchaseOrder
    {
        public int Id { get; set; }
        public string POReferenceNumber { get; set; }
        public string Author { get; set; }
        public string Created { get; set; }
        public string DepartmentName { get; set; }
        public string DepartmentID { get; set; }
        public string LocationName { get; set; }
        public string LocationID { get; set; }
        public string LocationType { get; set; }
        public string DivisionName { get; set; }
        public string DivisionID { get; set; }
        public string PONumber { get; set; }
        public string POCost { get; set; }
        public string MaterialDetails { get; set; }
        public string ApprovalStatus { get; set; }
        public string FHCode { get; set; }
        public string Modified { get; set; }

    }
    public class EmployeeMaster
    {
        public string Employee_x0020_Code { get; set; }
        public string Employee_x0020_Email { get; set; }

    }

    public class PurchaseApprovers
    {
        public string ApproverType { get; set; }
        public string ApproverName { get; set; }
        public string Location { get; set; }
        public string Division { get; set; }

    }

}
