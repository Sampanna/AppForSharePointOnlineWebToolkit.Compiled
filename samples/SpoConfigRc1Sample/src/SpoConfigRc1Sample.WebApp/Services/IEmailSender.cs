using System.Threading.Tasks;

namespace SpoConfigRc1Sample.WebApp.Services
{
    public interface IEmailSender
    {
        Task SendEmailAsync(string email, string subject, string message);
    }
}
