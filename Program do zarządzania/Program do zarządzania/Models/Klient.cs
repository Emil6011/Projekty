using Microsoft.EntityFrameworkCore;
using System.ComponentModel.DataAnnotations;

namespace Zadanie_1_MVC_Dobre.Models
{
    public class Klient
	{
		

		[Required(ErrorMessage = "Pole jest wymagane")]
        [Key] public int id { get; set; }
        [MaxLength(50)]
		[Required(ErrorMessage = "Pole jest wymagane")]
		public string name { get; set; }
        [MaxLength(50)]
		[Required(ErrorMessage = "Pole jest wymagane")]
		public string surname { get; set; }
        [MaxLength(11)]
		[Required(ErrorMessage = "Pole jest wymagane")]
		[RegularExpression("^[0-9]*$", ErrorMessage = "Wpisz tylko cyfry.")]
		public string pesel { get; set; }

      
        public int rok_urodzenia { get; set; }
	
        public int plec { get; set; }



    }
}
