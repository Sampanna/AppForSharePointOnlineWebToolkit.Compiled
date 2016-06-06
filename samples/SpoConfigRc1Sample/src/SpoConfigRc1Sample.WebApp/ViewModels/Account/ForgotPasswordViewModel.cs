using System.ComponentModel.DataAnnotations;

namespace SpoConfigRc1Sample.WebApp.ViewModels.Account
{
    public class ForgotPasswordViewModel
    {
        [Required]
        [EmailAddress]
        public string Email { get; set; }
    }
}
