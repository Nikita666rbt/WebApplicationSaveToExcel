using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace WebApplication4.Models
{
    public class Organization
    {
        [Required(ErrorMessage = "Требуется поле Наименование организации")]
        public string Name { get; set; }
        [Required(ErrorMessage = "Требуется поле Email")]
        public string Email { get; set; }
        [Required(ErrorMessage = "Требуется поле Телефон")]
        public string Phone { get; set; }
        [Required(ErrorMessage = "Требуется поле Юр.адрес")]
        public string LegalAddress { get; set; }
    }
}