using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Text;

namespace A2B_App.Shared.Sox
{
    public class Sod
    {
    }

    public class ConflictDefinition
    {
        public string Duty1 { get; set; }
        public string Duty2 { get; set; }
        public string RiskLevel { get; set; }
        public string Definition { get; set; }
    }

    public class SodUser
    {
        public string OwnersForManual { get; set; }
        public string Process { get; set; }
        public string DutyNum { get; set; }
        public string Function { get; set; }
        public string Users { get; set; }
    }

    public class RoleUser
    {
        public int Row { get; set; }
        public string Client { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public string Role { get; set; }

    }

    public class RolePerm 
    { 
        public string Client { get; set; }
        public int Row { get; set; }
        public string Permission { get; set; }
        public virtual List<Perm> ListPerm { get; set; }
    }

    public class Perm
    {
        public int Column { get; set; }
        public string Header { get; set; }
        public string Value { get; set; }
        
    }

    public class DescriptionToPerm
    {
        public string Client { get; set; }
        public int Row { get; set; }
        public string Description { get; set; }
        public string Permission { get; set; }
    }

    public class ConflictPerm
    {
        public string Client { get; set; }
        public int Row { get; set; }
        public string RefNum { get; set; }
        public string ProcessA { get; set; }
        public string ProcessB { get; set; }
        public string SODDescriptionA { get; set; }
        public string SODDescriptionB { get; set; }
        public string DescriptionOfConflict { get; set; }
        public string RiskComments { get; set; }
        public string System { get; set; }
        public string Manual { get; set; }
        public string Comment { get; set; }
        public string NsPermPairA { get; set; }
        public string NsPermPairB { get; set; }
    }

    public class SodSoxRoxImport
    {        
        public IFormFile FileRoleUser { get; set; }
        public IFormFile FileRolePerm { get; set; }
        public IFormFile FileConflictPerm { get; set; }
        public IFormFile FileDescToPerm { get; set; }
        public string ClientName { get; set; }
        public string RequestedBy { get; set; }
    }

    public class SoxRoxFile
    {
        public string FileName { get; set; }
        public string NewFileName { get; set; }
        public int Position { get; set; }
        public string ClientName { get; set; }
    }

    public class SodSoxRoxInput
    {
        public List<RoleUser> ListRoleUser { get; set; }
        public List<RolePerm> ListRolePerm { get; set; }
        public List<DescriptionToPerm> ListDescriptionToPerm { get; set; }
        public List<ConflictPerm> ListConflictPerm { get; set; }
        public List<SodSoxRoxRoleUser> ListRoleUserTrim { get; set; }
    }



    public class SodSoxRoxRoleUser
    {
        public string Name { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public List<string> Role { get; set; }
        public List<string> Permission { get; set; }
    }

    #region Output File Objects
    public class SodSoxRoxDescriptionPermission
    {
        public string Description { get; set; }
        public string Permission { get; set; }
        public int SummaryCount { get; set; }
        public int Count { get; set; }
    }   

    public class SODSoxRoxReportRaw2
    {
        public string Name { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public string SODRef { get; set; }
        public string RoleA { get; set; }
        public string RoleB { get; set; }
        public string PermA { get; set; }
        public string PermB { get; set; }
        public string ProcessA { get; set; }
        public string ProcessB { get; set; }
        public string FunctionTypeA { get; set; }
        public string FunctionTypeB { get; set; }
        public string DescOfConflict { get; set; }
        public string DescriptionA { get; set; }
        public string DescriptionB { get; set; }
        public string RiskComments { get; set; }
        public string SODPriority { get; set; }
        public string ConflictRolePair { get; set; }
        public string ConflictType { get; set; }
    }

    public class SODSoxRoxReportRaw3
    {
        public string SODRef { get; set; }
        public string RoleA { get; set; }
        public string RoleB { get; set; }
        public string PermA { get; set; }
        public string PermB { get; set; }
        public string ProcessA { get; set; }
        public string ProcessB { get; set; }
        public string DescOfConflict { get; set; }
        public string DescriptionA { get; set; }
        public string DescriptionB { get; set; }
        public string RiskComments { get; set; }
        public string SODPriority { get; set; }
        public string ConflictRolePair { get; set; }
        public string ConflictType { get; set; }
        public string AssignedA { get; set; }
        public string AssignedB { get; set; }
    }

    public class SODSoxRoxReportRaw4
    {
        public string AssignedA { get; set; }
        public string AssignedB { get; set; }
        public string ConflictType { get; set; }
        public string ConflictRolePair { get; set; }
        public string Count1 { get; set; }
        public string DescOfConflict { get; set; }
        public string DescriptionA { get; set; }
        public string DescriptionB { get; set; }
        public string PermA { get; set; }
        public string PermB { get; set; }
        public string ProcessA { get; set; }
        public string ProcessB { get; set; }
        public string RiskComments { get; set; }
        public string RoleA { get; set; }
        public string RoleB { get; set; }
        public string SODPriority { get; set; }
        public string SODRef { get; set; }

    }

    public class SodSoxRoxOutputFile
    {
        public string ReportFileName { get; set; }
        public string DescriptionFileName { get; set; }
    }
    #endregion

}
