using System.Threading.Tasks;

namespace SpoConfigRc1Sample.WebApp.Services
{
    public interface ISmsSender
    {
        Task SendSmsAsync(string number, string message);
    }
}
