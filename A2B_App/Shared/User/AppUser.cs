using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text;
using System.Text.Json.Serialization;

namespace A2B_App.Shared.User
{
    public class AppUser
    {
        public string Id { get; set; }

        [Required(ErrorMessage = "Email is required.")]
        [EmailAddress()]
        public string Email { get; set; }

        public string UserName { get; set; }

        [Required(ErrorMessage = "Password is required.")]
        [MinLength(8, ErrorMessage = "Password must be atleast 8 characters")]
        [DataType(DataType.Password)]
        public string Password { get; set; }

        public string PhoneNumber { get; set; }
        public bool PhoneNumberConfirmed { get; set; }
        public bool LockoutEnabled { get; set; }
        public bool EmailConfirmed { get; set; }
        [Required(ErrorMessage = "Confirm password is required.")]
        [DataType(DataType.Password)]
        [Compare(nameof(Password), ErrorMessage = "Password doesn't match")]
        public string ConfirmPassword { get; set; }
        public virtual ICollection<AppRole> ListAppRole { get; set; }
    }

    public class AppRole
    {
        public string Id { get; set; }
        public string RoleName { get; set; }
        public string PrevRoleName { get; set; }
    }

    public class AppUserRoleRef
    {
        public AppUser AppUser { get; set; }
        public AppRole AppRole{ get; set; }
    }

    public class PageTableFilter
    {
        public int PageSize { get; set; }
        public int PageNumber { get; set; }
    }


}
