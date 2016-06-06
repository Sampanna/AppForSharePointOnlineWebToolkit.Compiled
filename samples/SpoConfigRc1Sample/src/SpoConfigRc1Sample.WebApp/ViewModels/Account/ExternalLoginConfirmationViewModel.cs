using System.ComponentModel.DataAnnotations;

namespace SpoConfigRc1Sample.WebApp.ViewModels.Account
{
    public class ExternalLoginConfirmationViewModel
    {
        [Required]
        [EmailAddress]
        public string Email { get; set; }
    }
}
