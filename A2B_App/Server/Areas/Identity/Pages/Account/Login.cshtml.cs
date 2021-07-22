using A2B_App.Server.Models;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using PodioAPI;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Web;

namespace A2B_App.Server.Areas.Identity.Pages.Account
{
    [AllowAnonymous]
    public class LoginModel : PageModel
    {
        private readonly UserManager<ApplicationUser> _userManager;
        private readonly SignInManager<ApplicationUser> _signInManager;
        private readonly RoleManager<IdentityRole> _roleManager;
        private readonly ILogger<LoginModel> _logger;
        private readonly IConfiguration _config;

        public LoginModel(SignInManager<ApplicationUser> signInManager, 
            ILogger<LoginModel> logger,
            UserManager<ApplicationUser> userManager,
            RoleManager<IdentityRole> roleManager,
            IConfiguration config)
        {
            _userManager = userManager;
            _signInManager = signInManager;
            _logger = logger;
            _config = config;
            _roleManager = roleManager;
        }

        [BindProperty]
        public InputModel Input { get; set; }

        public IList<AuthenticationScheme> ExternalLogins { get; set; }

        public string ReturnUrl { get; set; }

        [TempData]
        public string ErrorMessage { get; set; }

        public class InputModel
        {
            [Required]
            [EmailAddress]
            public string Email { get; set; }

            [Required]
            [DataType(DataType.Password)]
            public string Password { get; set; }

            [Display(Name = "Remember me?")]
            public bool RememberMe { get; set; }
        }

        public async Task<IActionResult> OnGetAsync(string returnUrl = null, string code = null)
        {
            if (!string.IsNullOrEmpty(ErrorMessage))
            {
                ModelState.AddModelError(string.Empty, ErrorMessage);
            }

            returnUrl = returnUrl ?? Url.Content("~/");

            // Clear the existing external cookie to ensure a clean login process
            await HttpContext.SignOutAsync(IdentityConstants.ExternalScheme);

            ExternalLogins = (await _signInManager.GetExternalAuthenticationSchemesAsync()).ToList();

            ReturnUrl = returnUrl;

            if (!string.IsNullOrEmpty(code))
            {
                var podio = new Podio(_config.GetSection("PodioApi").GetSection("ClientId").Value, _config.GetSection("PodioApi").GetSection("ClientSecret").Value);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithAuthorizationCode(code, this.Host);
                if (podio.IsAuthenticated())
                {
                    _logger.LogInformation("Podio Authenticated!");
                    var puser = await podio.UserService.GetUser();

                    _logger.LogInformation("Podio user status:" + puser.Status);
                    if (puser.Status == "active")
                    {
                        string email = puser.Mail;
                        var user = new ApplicationUser { UserName = email, Email = email };

                        #region Check if valid email domain "a2bhq.com a2q2.com ateamhq.com
                        //var validEmailDomain = _config.GetSection("AcceptedEmailHost").Value;
                        if(isValidEmailDomain(email))
                        {
                            var user2 = _userManager.Users.Where(x => x.NormalizedEmail.Equals(email.ToUpper())).FirstOrDefault();
                            if (user2 == null)
                            {
                                string password = "Az10#" + puser.Mail + puser.UserId;
                                var create_result = await _userManager.CreateAsync(user, password);
                                if (create_result.Succeeded)
                                {
                                    _logger.LogInformation("User created a new account with password.");

                                    var user_code = await _userManager.GenerateEmailConfirmationTokenAsync(user);
                                    var result_conf = await _userManager.ConfirmEmailAsync(user, user_code);

                                    if (result_conf.Succeeded)
                                    {

                                        _logger.LogInformation("User Created!");
                                        var login_result = await _signInManager.PasswordSignInAsync(user.Email, password, true, lockoutOnFailure: false);
                                        if (login_result.Succeeded)
                                            return LocalRedirect(returnUrl);
                                        else
                                            ModelState.AddModelError(string.Empty, "Podio Account is not active!");
                                    }
                                    else
                                    {
                                        ModelState.AddModelError(string.Empty, "Something went wrong, please contact administration.");
                                    }
                                }
                                else
                                {
                                    ModelState.AddModelError(string.Empty, "Something went wrong, please contact administration.");
                                }
                            }
                            else
                            {
                                await _signInManager.SignInAsync(user2, true);
                                _logger.LogInformation("User logged in.");
                                return LocalRedirect(returnUrl);
                            }
                        }
                        else
                            ModelState.AddModelError(string.Empty, "Podio account is not allowed!");
                        #endregion



                    }
                    else
                        ModelState.AddModelError(string.Empty, "Podio account is not active!");
                }
                else
                {
                    ModelState.AddModelError(string.Empty, "Invalid login attempt.");
                }
            }
            return Page();
        }

        public async Task<IActionResult> OnPostAsync(string returnUrl = null)
        {
            returnUrl = returnUrl ?? Url.Content("~/");

            if (ModelState.IsValid)
            {
                // This doesn't count login failures towards account lockout
                // To enable password failures to trigger account lockout, set lockoutOnFailure: true
                var result = await _signInManager.PasswordSignInAsync(Input.Email, Input.Password, Input.RememberMe, lockoutOnFailure: false);
                if (result.Succeeded)
                {
                    _logger.LogInformation("User logged in.");
                    return LocalRedirect(returnUrl);
                }
                if (result.RequiresTwoFactor)
                {
                    return RedirectToPage("./LoginWith2fa", new { ReturnUrl = returnUrl, RememberMe = Input.RememberMe });
                }
                if (result.IsLockedOut)
                {
                    _logger.LogWarning("User account locked out.");
                    return RedirectToPage("./Lockout");
                }
                else
                {
                    ModelState.AddModelError(string.Empty, "Invalid login attempt.");
                    return Page();
                }
            }

            // If we got this far, something failed, redisplay form
            return Page();
        }

        public string Host
        {
            get
            {

                bool is_development = _config.GetSection("ServerHost").GetSection("Active").Value == "development";
                String return_url = _config.GetSection("ServerHost").GetSection("Prod").Value;
                if (is_development)
                    return_url = _config.GetSection("ServerHost").GetSection("Dev").Value;
                return return_url;
            }
        }

        public string PodioUrl
        {
            get
            {
                string client_id = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                return "https://podio.com/oauth/authorize?response_type=code&client_id=" + client_id + "&redirect_uri=" + HttpUtility.UrlEncode(this.Host + "/Identity/Account/Login");
            }

        }

        private bool isValidEmailDomain(string email)
        {
            bool result = false;

            var validEmailDomain = _config.GetSection("AcceptedEmailHost").AsEnumerable().Where(x => x.Value != null);
            if(validEmailDomain != null)
            {
                Uri uri = new Uri($"mailto:{email}");
                string emailHost = uri.Host;

                var match = validEmailDomain.Where(x => x.Value.ToLower().Equals(emailHost.ToLower())).FirstOrDefault();
                if (match.Value != null)
                    result = true;                
            }

            return result;
        }

    }
}
