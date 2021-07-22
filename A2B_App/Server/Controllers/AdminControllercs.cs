
using A2B_App.Server.Data;
using A2B_App.Server.Models;
using A2B_App.Server.Services;
using A2B_App.Shared.Sox;
using A2B_App.Shared.User;
using IdentityModel.Client;
using IdentityServer4.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.IdentityModel.Tokens;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IdentityModel.Tokens.Jwt;
using System.Linq;
using System.Net.Http;
using System.Security.Claims;
using System.Text;
using System.Threading.Tasks;


namespace A2B_App.Server.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class AdminController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly ILogger<AdminController> _logger;
        private readonly IWebHostEnvironment _environment;
        private readonly UserManager<ApplicationUser> _userManager;
        private readonly RoleManager<IdentityRole> _roleManager;
        private  SignInManager<ApplicationUser> _signInManager;
        private UserContext _userContext;


        public AdminController(IConfiguration config,
            ILogger<AdminController> logger,
            IWebHostEnvironment environment,
            SignInManager<ApplicationUser> signInManager,
            UserManager<ApplicationUser> userManager,
            RoleManager<IdentityRole> roleManager,
            UserContext userContext)
        {
            _config = config;
            _logger = logger;
            _environment = environment;
            _signInManager = signInManager;
            _userManager = userManager;
            _roleManager = roleManager;
            _userContext = userContext;
        }

        //[AllowAnonymous]
        [HttpGet("admin/roles")]
        public IActionResult AdminIdentityRoles()
        {
            List<AppRole> appRole = null;
            try
            {
                
                appRole = new List<AppRole>();
                var role = _roleManager.Roles.ToList();
                if(role != null)
                {
                    foreach (var item in role)
                    {
                        appRole.Add(new AppRole { Id = item.Id, RoleName = item.Name });
                    }
                }
                

                return Ok(appRole);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorAdminIdentityRoles");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "AdminIdentityRoles");
                return BadRequest(ex.ToString());
            }



        }


        [HttpPost("admin/roles")]
        public IActionResult AdminIdentityRoles([FromBody] PageTableFilter pageTableFilter)
        {
            try
            {
              
                if (pageTableFilter.PageNumber < 1)
                    pageTableFilter.PageNumber = 1;

                var role = _roleManager.Roles.Skip((pageTableFilter.PageNumber - 1) * pageTableFilter.PageSize).Take(pageTableFilter.PageSize).ToList();
                
                return Ok(role.ToArray());
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorAdminIdentityRoles");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "AdminIdentityRoles");
                return BadRequest(ex.ToString());
            }



        }


        [HttpPost("admin/roles/create")]
        public async Task<IActionResult> AdminIdentityCreateRolesAsync([FromBody] AppRole appRole)
        {

            try
            {

                var role = _roleManager.Roles.Where(x => x.NormalizedName.Equals(appRole.RoleName.ToUpper())).FirstOrDefault();
                if (role == null)
                {

                    IdentityRole newRole = new IdentityRole();
                    Guid g = Guid.NewGuid();
                    newRole.Id = g.ToString();
                    newRole.Name = appRole.RoleName;
                    newRole.NormalizedName = appRole.RoleName.ToUpper();
                    await _roleManager.CreateAsync(newRole);
                    return Ok();
                }

                return BadRequest();


            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorAdminIdentityCreateRoles");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "AdminIdentityCreateRoles");
                return BadRequest(ex.ToString());
            }



        }


        [HttpPut("admin/roles/update")]
        public async Task<IActionResult> AdminIdentityUpdateRolesAsync([FromBody] AppRole appRole)
        {

            try
            {

                var role = _roleManager.Roles.Where(x => x.NormalizedName.Equals(appRole.PrevRoleName.ToUpper())).FirstOrDefault();
                if (role != null)
                {

                    IdentityRole tempRole = new IdentityRole();
                    tempRole = role;
                    tempRole.Name = appRole.RoleName;
                    tempRole.NormalizedName = appRole.RoleName.ToUpper();

                    await _roleManager.UpdateAsync(tempRole);
                    return Ok();
                }

                return NoContent();


            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorAdminIdentityUpdateRoles");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "AdminIdentityUpdateRoles");
                return BadRequest(ex.ToString());
            }



        }


        [HttpDelete("admin/roles/delete")]
        public async Task<IActionResult> AdminIdentityDeleteRolesAsync([FromBody] AppRole appRole)
        {

            try
            {

                var role = _roleManager.Roles.Where(x => x.Id.Equals(appRole.Id)).FirstOrDefault();
                if (role != null)
                {
                    await _roleManager.DeleteAsync(role);
                    return Ok();
                }

                return NoContent();


            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorAdminIdentityDeleteRoles");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "AdminIdentityDeleteRoles");
                return BadRequest(ex.ToString());
            }



        }


        [HttpPost("admin/user")]
        public IActionResult AdminIdentityUser([FromBody] PageTableFilter pageTableFilter)
        {

            List<AppUser> listAppUser = null;
            try
            {

                if (pageTableFilter.PageNumber < 1)
                    pageTableFilter.PageNumber = 1;

                var user = _userManager.Users.Skip((pageTableFilter.PageNumber - 1) * pageTableFilter.PageSize).Take(pageTableFilter.PageSize).ToList();

                if (user != null)
                {
                    listAppUser = new List<AppUser>();
                    foreach (var item in user)
                    {

                        AppUser appUser = new AppUser();
                        appUser.Id = item.Id;
                        appUser.Email = item.Email;
                        appUser.UserName = item.UserName;
                        appUser.PhoneNumber = item.PhoneNumber;
                        appUser.PhoneNumberConfirmed = item.PhoneNumberConfirmed;
                        appUser.LockoutEnabled = item.LockoutEnabled;
                        appUser.EmailConfirmed = item.EmailConfirmed;

                        List<AppRole> listRole = new List<AppRole>();

                        var listRoleName = _userManager.GetRolesAsync(item).Result.ToList();
                        if (listRoleName != null)
                        {
                            foreach (var itemRoles in listRoleName)
                            {
                                var role = _roleManager.Roles.Where(x => x.Name.Equals(itemRoles)).ToList();
                                if (role != null)
                                {
                                    foreach (var roleItem in role)
                                    {
                                        AppRole appRole = new AppRole();
                                        appRole.Id = roleItem.Id;
                                        appRole.RoleName = roleItem.Name;
                                        listRole.Add(appRole);

                                    }
                                }
                            }
                        }
                        if (listRole.Count > 0)
                        {
                            appUser.ListAppRole = listRole;
                        }

                        listAppUser.Add(appUser);
                    }
                }


                return Ok(listAppUser);
            }
            catch (Exception ex)
            {

                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorAdminIdentityUser");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "AdminIdentityUser");
                return BadRequest(ex.ToString());
            }

        }


        //[AllowAnonymous]
        [HttpGet("admin/user/get")]
        public IActionResult AdminIdentityGetUser(string email)
        {

            AppUser appUser = new AppUser();
            try
            {
             
                var user = _userManager.Users.Where(x => x.NormalizedEmail.Equals(email.ToUpper())).FirstOrDefault();

                if (user != null)
                {

                    appUser.Id = user.Id;
                    appUser.Email = user.Email;
                    appUser.UserName = user.UserName;
                    appUser.PhoneNumber = user.PhoneNumber;
                    appUser.PhoneNumberConfirmed = user.PhoneNumberConfirmed;
                    appUser.LockoutEnabled = user.LockoutEnabled;
                    appUser.EmailConfirmed = user.EmailConfirmed;

                    List<AppRole> listRole = new List<AppRole>();

                    var listRoleName = _userManager.GetRolesAsync(user).Result.ToList();
                    if (listRoleName != null)
                    {
                        foreach (var itemRoles in listRoleName)
                        {
                            var role = _roleManager.Roles.Where(x => x.Name.Equals(itemRoles)).ToList();
                            if (role != null)
                            {
                                foreach (var roleItem in role)
                                {
                                    AppRole appRole = new AppRole();
                                    appRole.Id = roleItem.Id;
                                    appRole.RoleName = roleItem.Name;
                                    listRole.Add(appRole);

                                }
                            }
                        }
                    }
                    if (listRole.Count > 0)
                    {
                        appUser.ListAppRole = listRole;
                    }
                }


                return Ok(appUser);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "AdminIdentityGetUser");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "AdminIdentityGetUser");
                return BadRequest(ex.ToString());
            }
        }


        [HttpPost("admin/user/create")]
        public async Task<IActionResult> AdminIdentityCreateUserAsync([FromBody] AppUser appUser)
        {

            try
            {
            
                var user = _userManager.Users.Where(x => x.NormalizedEmail.Equals(appUser.Email.ToUpper())).FirstOrDefault();
                if (user == null)
                {
                    user = new ApplicationUser();
                    user.Email = appUser.Email;
                    user.UserName = appUser.Email;
                    user.EmailConfirmed = true;

                    var result = await _userManager.CreateAsync(user, appUser.Password);

                    if (result.Succeeded)
                    {
                        return Ok();
                    }

                    return BadRequest(result.Errors.ToArray());
                }

                return NoContent();

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorAdminIdentityCreateUser");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "AdminIdentityCreateUser");
                return BadRequest(ex.ToString());
            }



        }


        [HttpPut("admin/user/update")]
        public async Task<IActionResult> AdminIdentityUpdateUserAsync([FromBody] AppUser appUser)
        {

            try
            {
                
                var user = _userManager.Users.Where(x => x.NormalizedEmail.Equals(appUser.Email.ToUpper())).FirstOrDefault();
                if (user != null)
                {
                    string resetToken = await _userManager.GeneratePasswordResetTokenAsync(user);
                    var updateResult = await _userManager.ResetPasswordAsync(user, resetToken, appUser.Password);

                    //var resultPassChange = await _userManager.ChangePasswordAsync(user, null, appUser.Password);
                    if (updateResult.Succeeded)
                    {
                        var result = await _userManager.UpdateAsync(user);

                        if (result.Succeeded)
                        {
                            return Ok();
                        }

                        return BadRequest(result.Errors.ToArray());

                    }
                }

                return NoContent();

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorAdminIdentityUpdateUser");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "AdminIdentityUpdateUser");
                return BadRequest(ex.ToString());
            }



        }


        [HttpDelete("admin/user/delete")]
        public async Task<IActionResult> AdminIdentityDeleteUserAsync([FromBody] AppUser appUser)
        {

            try
            {
                var user = _userManager.Users.Where(x => x.NormalizedEmail.Equals(appUser.Email.ToUpper())).FirstOrDefault();
                if (user != null)
                {

                    var result = await _userManager.DeleteAsync(user);

                    if (result.Succeeded)
                    {
                        return Ok();
                    }
                    return BadRequest(result.Errors.ToArray());
                }

                return NoContent();

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorAdminIdentityDeleteUser");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "AdminIdentityGetUser");
                return BadRequest(ex.ToString());
            }



        }


        [HttpPut("admin/user/updaterole")]
        public async Task<IActionResult> AdminIdentityUpdateUserRoleAsync([FromBody] AppUser appUser)
        {

            try
            {
                var user = _userManager.Users.Where(x => x.NormalizedEmail.Equals(appUser.Email.ToUpper())).FirstOrDefault();
                if (user != null)
                {
                    foreach (var item in appUser.ListAppRole)
                    {
                        var checkInRole = _userManager.IsInRoleAsync(user, item.RoleName).Result;
                        if (!checkInRole)
                        {
                            var roleResult = await _userManager.AddToRoleAsync(user, item.RoleName);
                        }
                    }

                    appUser.Id = user.Id;
                    appUser.UserName = user.UserName;
                    appUser.PhoneNumber = user.PhoneNumber;
                    appUser.PhoneNumberConfirmed = user.PhoneNumberConfirmed;
                    appUser.LockoutEnabled = user.LockoutEnabled;
                    appUser.EmailConfirmed = user.EmailConfirmed;

                    List<AppRole> listRole = new List<AppRole>();

                    var listRoleName = _userManager.GetRolesAsync(user).Result.ToList();
                    if (listRoleName != null)
                    {
                        foreach (var itemRoles in listRoleName)
                        {
                            var role = _roleManager.Roles.Where(x => x.Name.Equals(itemRoles)).ToList();
                            if (role != null)
                            {
                                foreach (var roleItem in role)
                                {
                                    AppRole appRole = new AppRole();
                                    appRole.Id = roleItem.Id;
                                    appRole.RoleName = roleItem.Name;
                                    listRole.Add(appRole);

                                }
                            }
                        }
                    }
                    if (listRole.Count > 0)
                    {
                        appUser.ListAppRole = listRole;
                    }

                    return Ok(appUser);


                }

                return NoContent();

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorAdminIdentityUpdateUser");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "AdminIdentityUpdateUser");
                return BadRequest(ex.ToString());
            }



        }


        [HttpPut("admin/user/deleterole")]
        public async Task<IActionResult> AdminIdentityDeleteUserRoleAsync([FromBody] AppUser appUser)
        {

            try
            {
                var user = _userManager.Users.Where(x => x.NormalizedEmail.Equals(appUser.Email.ToUpper())).FirstOrDefault();
                if (user != null)
                {
                    foreach (var item in appUser.ListAppRole)
                    {
                        var checkInRole = _userManager.IsInRoleAsync(user, item.RoleName).Result;
                        if (checkInRole)
                        {
                            var roleResult = await _userManager.RemoveFromRoleAsync(user, item.RoleName);

                        }
                    }

                    return Ok(user);

                }

                return NoContent();

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorAdminIdentityUpdateUser");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "AdminIdentityUpdateUser");
                return BadRequest(ex.ToString());
            }



        }


        [HttpPost("admin/user/adduserroleref")]
        public async Task<IActionResult> AdminIdentityAddUserRoleRefAsync([FromBody] AppUserRoleRef appUserRoleRef)
        {

            try
            {
                var role = _roleManager.Roles.Where(x => x.Id.Equals(appUserRoleRef.AppRole.Id)).FirstOrDefault();
                var user = _userManager.Users.Where(x => x.Id.Equals(appUserRoleRef.AppUser.Id)).FirstOrDefault();
                if (user != null && role != null)
                {

                    var isExists = await _userManager.IsInRoleAsync(user, role.Name);

                    if (!isExists)
                    {
                        var result = await _userManager.AddToRoleAsync(user, role.Name);

                        if (result.Succeeded)
                        {
                            return Ok();
                        }
                        return BadRequest(result.Errors.ToArray());
                    }
                }

                return NoContent();

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorAdminIdentityAddUserRoleRefAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "AdminIdentityAddUserRoleRefAsync");
                return BadRequest(ex.ToString());
            }



        }


        [HttpPut("user/deluserroleref")]
        public async Task<IActionResult> AdminIdentityDelUserRoleRefAsync([FromBody] AppUserRoleRef appUserRoleRef)
        {

            try
            {
                var role = _roleManager.Roles.Where(x => x.Id.Equals(appUserRoleRef.AppRole.Id)).FirstOrDefault();
                var user = _userManager.Users.Where(x => x.Id.Equals(appUserRoleRef.AppUser.Id)).FirstOrDefault();
                if (user != null && role != null)
                {

                    var isExists = await _userManager.IsInRoleAsync(user, role.Name);

                    if (isExists)
                    {
                        var result = await _userManager.RemoveFromRoleAsync(user, role.Name);

                        if (result.Succeeded)
                        {
                            return Ok();
                        }
                        return BadRequest(result.Errors.ToArray());
                    }
                }

                return NoContent();

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorAdminIdentityDelUserRoleRefAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "AdminIdentityDelUserRoleRefAsync");
                return BadRequest(ex.ToString());
            }



        }

        //[AllowAnonymous]
        //[HttpPost("createUserDetail")]
        //public async Task<IActionResult> CreateUserDetail(AppUserDetail userDetail)
        //{
        //    try
        //    {
        //        using (var context = _userContext.Database.BeginTransaction())
        //        {
        //            var checkExistUser = _userContext.AppUserDetail.Where(id => id.UserId.Equals(userDetail.UserId)).FirstOrDefault();
        //            if (checkExistUser == null)
        //            {
        //                _userContext.Add(userDetail);
        //                await _userContext.SaveChangesAsync();
        //                context.Commit();
        //            }
        //        }
                
        //        return Ok();

        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest(ex.ToString());
        //    }

        //}


        [AllowAnonymous]
        [HttpPost("token")]
        public IActionResult CreateJWTTOken()
        {
            try
            {
                var header = Request.Headers["Authorization"];
                if(header.ToString().StartsWith("Basic"))
                {
                    var credValue = header.ToString().Substring("Basic ".Length).Trim();
                    var credEnc = Encoding.UTF8.GetString(Convert.FromBase64String(credValue));
                    //credEnc =  "email:password"
                    var credArray = credEnc.Split(":");
                    if(credArray[0] == "Admin" && credArray[1] == "123qwe")
                    {
                        var key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(_config["TokenSettings:IssuerSigningKey"]));
                        var claimsParam = new[] { new Claim(ClaimTypes.Name, "username") };
                        var signInCredParam = new SigningCredentials(key, SecurityAlgorithms.HmacSha256);

                        var token = new JwtSecurityToken(
                                issuer: _config["TokenSettings:Issuer"],
                                audience: _config["TokenSettings:Audience"],
                                expires: DateTime.UtcNow.AddMinutes(double.Parse(_config["TokenSettings:ExpireInMinutes"])),
                                notBefore: DateTime.UtcNow,
                                claims: claimsParam,
                                signingCredentials: signInCredParam
                            );
                        var tokenStr = new JwtSecurityTokenHandler().WriteToken(token);
                        return Ok(tokenStr);
                    }
                }

                return BadRequest();

            }
            catch (Exception ex)
            {
                return BadRequest(ex.ToString());
            }
            
        }

        [AllowAnonymous]
        [HttpPost("gettoken")]
        public async Task<IActionResult> Token()
        {
            try
            {
                var header = Request.Headers["Authorization"];
                if (header.ToString().StartsWith("Basic"))
                {
                    var credValue = header.ToString().Substring("Basic ".Length).Trim();
                    var credEnc = Encoding.UTF8.GetString(Convert.FromBase64String(credValue));
                    //credEnc =  "email:password"
                    var credArray = credEnc.Split(":");

                    //var userName = await _userManager.FindByNameAsync(credArray[0]);
                    var user = _userManager.Users.Where(x => x.NormalizedEmail.Equals(credArray[0].ToUpper())).FirstOrDefault();
                    if (user != null)
                    {

                        if (!_userManager.IsEmailConfirmedAsync(user).Result) //if email address is not verified
                            return Forbid("Error 403");

                        //var signIn = _signInManager.SignInAsync(user, true);
                        //if (signIn.IsCompletedSuccessfully)
                        //{
                        //    //var token = await _userManager.CreateSecurityTokenAsync(user);
                        //    //var token = await _userManager.GetAuthenticationTokenAsync(user, TokenOptions.DefaultProvider, "access_token");
                        //    //string accessToken = await HttpContext.GetTokenAsync("access_token");
                        //    //return Ok(accessToken);
                        //    var userClaims = await _userManager.GetClaimsAsync(user);

                        //    return Ok(accessToken);
                        //}
                        var client = new HttpClient();
                        var response = await client.RequestTokenAsync(new IdentityModel.Client.TokenRequest
                        {
                            Address = "http://localhost",
                            GrantType = "custom",

                            ClientId = "client",
                            ClientSecret = "secret",

                            Parameters =
                            {
                                { "custom_parameter", "custom value"},
                                { "scope", "api1" }
                            }
                        });
                        if (response.IsError) throw new Exception(response.Error);

                        var token = response.AccessToken;
                        var custom = response.Json.TryGetString("custom_parameter");


                    }


                }

                
            }
            catch (Exception ex)
            {
                return BadRequest(ex.ToString());
            }


            return NoContent();

        }

        [HttpPost("createclient")]
        public IActionResult CreateClient()
        {
            var client = new IdentityServer4.Models.Client
            {
                ClientId = "service.client",
                ClientSecrets = { new IdentityServer4.Models.Secret("secret".Sha256()) },

                AllowedGrantTypes = GrantTypes.ClientCredentials,
                AllowedScopes = { "api1", "api2.read_only" }
            };

            return Ok(client);
        }

    }



}



