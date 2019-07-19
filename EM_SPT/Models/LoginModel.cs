using System.ComponentModel.DataAnnotations;

namespace EM_SPT.Models
{
    public class LoginModel
    {
        [Required(ErrorMessage = "�� ������ �����")]
        public string Login { get; set; }

        [Required(ErrorMessage = "�� ������ ������")]
        [DataType(DataType.Password)]
        public string Password { get; set; }
    }
}