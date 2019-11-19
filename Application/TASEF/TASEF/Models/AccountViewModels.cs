using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace TASEF.Models
{
    public class ExternalLoginConfirmationViewModel
    {
        [Required]
        [Display(Name = "Email")]
        public string Email { get; set; }
    }

    public class ExternalLoginListViewModel
    {
        public string ReturnUrl { get; set; }
    }

    public class SendCodeViewModel
    {
        public string SelectedProvider { get; set; }
        public ICollection<System.Web.Mvc.SelectListItem> Providers { get; set; }
        public string ReturnUrl { get; set; }
        public bool RememberMe { get; set; }
    }

    public class VerifyCodeViewModel
    {
        [Required]
        public string Provider { get; set; }

        [Required]
        [Display(Name = "Code")]
        public string Code { get; set; }
        public string ReturnUrl { get; set; }

        [Display(Name = "Remember this browser?")]
        public bool RememberBrowser { get; set; }

        public bool RememberMe { get; set; }
    }

    public class ForgotViewModel
    {
        [Required]
        [Display(Name = "email")]
        public string Email { get; set; }
    }

    public class LoginViewModel
    {
        [Required]
        [Display(Name = "email")]
        [EmailAddress]
        public string Email { get; set; }

        [Required]
        [DataType(DataType.Password)]
        [Display(Name = "mot de passe")]
        public string Password { get; set; }

        [Display(Name = "Sauvegarder mon mot de passe")]
        public bool RememberMe { get; set; }
    }

    public class RegisterViewModel
    {

        [Required]
        [EmailAddress]
        [Display(Name = "email")]
        public string Email { get; set; }

        [Required]
        [StringLength(100, ErrorMessage = "Le {0} doit être au moins de longuer {2} .", MinimumLength = 6)]
        [DataType(DataType.Password)]
        [Display(Name = "mot de passe")]
        public string Password { get; set; }

        [DataType(DataType.Password)]
        [Display(Name = "Confirmer mot de passe")]
        [Compare("Password", ErrorMessage = "Le nouveau mot de passe et le mot de passe de confirmation ne sont pas identiques.")]
        public string ConfirmPassword { get; set; }

        [Required]
        [Display(Name = "Nom")]
        public string FirstName { get; set; }

    }

    public class ResetPasswordViewModel
    {
        [Required]
        [EmailAddress]
        [Display(Name = "email")]
        public string Email { get; set; }

        [Required]
        [StringLength(100, ErrorMessage = "Le {0} doit être au moins de longeur {2}.", MinimumLength = 6)]
        [DataType(DataType.Password)]
        [Display(Name = "mot de passe")]
        public string Password { get; set; }

        [DataType(DataType.Password)]
        [Display(Name = "Confirmer mot de passe")]
        [Compare("Password", ErrorMessage = "Le nouveau mot de passe et le mot de passe de confirmation ne sont pas identiques.")]
        public string ConfirmPassword { get; set; }

        public string Code { get; set; }
    }

    public class ForgotPasswordViewModel
    {
        [Required]
        [EmailAddress]
        [Display(Name = "email")]
        public string Email { get; set; }
    }
}
