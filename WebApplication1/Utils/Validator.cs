using System.Text.RegularExpressions;
using System.Web;

namespace WebApplication1.Models
{
    public static class Validator
    {
        public static bool IsValidFormat(HttpPostedFileBase file)
        {
            return Regex.IsMatch(file.FileName, @"(\.pdf|\.docx|\.txt)$");
        }
    }
}