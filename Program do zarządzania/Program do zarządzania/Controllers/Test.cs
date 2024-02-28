using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Newtonsoft.Json.Linq;
using PostgreSQL.Data;
using System.Security.Policy;
using System.Text;
using Zadanie_1_MVC_Dobre.Models;

namespace Zadanie_1_MVC.Controllers
{
    public class Test : Controller
    {
        private readonly AppDbContext _appDbContext;

        public Test(AppDbContext appDbContext)
        {
            _appDbContext = appDbContext;
        }


        public IActionResult Index()
        {

			

			List<Klient> listaKlientow = new List<Klient>(_appDbContext.Klienci.OrderBy(c => c.id).ToList());



            return View(listaKlientow);
        }



        public IActionResult Create()
        {

            return View();
        }

        public int rok (string substring)
        {
			String rok = substring;

			int temp = Convert.ToInt32(rok);

			if (temp <= Convert.ToInt32(System.DateTime.Now.ToString("yy")))
			{
				rok = "20";
			}
			else
			{
				rok = "19";
			}

			return ( Convert.ToInt32(rok + substring));
		}


		public int plec(string substring)
		{
			int temp = Convert.ToInt32(substring);


			if (temp % 2 == 0)
			{
				return 0;


			}
			else
			{
				return 1;

			}


			
		}


			[HttpPost]
		public IActionResult Dodaj(Klient klient)
		{
            if (ModelState.IsValid) {

				klient.rok_urodzenia = rok(klient.pesel.Substring(0, 2));
				
				klient.plec = plec(klient.pesel.Substring(9, 1));

				_appDbContext.Klienci.Add(klient);
            _appDbContext.SaveChanges();
			return RedirectToAction("Index");
			}
            else
            {
				return View("Create", klient);
			}
			
		}


		public IActionResult Edit(Klient klient)
        {
         
            return View(klient);
        }


       


		[HttpPost]
		public IActionResult Edytuj(Klient klient)
		{

            Klient stary  = _appDbContext.Klienci.Find(klient.id);

            stary.name = klient.name;
			stary.surname = klient.surname;
            stary.pesel = klient.pesel;
			stary.plec = plec(klient.pesel.Substring(9, 1));
			stary.rok_urodzenia =  rok(klient.pesel.Substring(0, 2));


			_appDbContext.Klienci.Update(stary);
            _appDbContext.SaveChanges();




			return RedirectToAction("Index");
		}


		[HttpPost]
		public IActionResult Usun(int id)
		{
           
            Klient stary = _appDbContext.Klienci.Find(id);
			

			_appDbContext.Klienci.Remove(stary);


            _appDbContext.SaveChanges();

			return RedirectToAction("Index");
		}

		[HttpPost]
		public IActionResult CSV()
		{





			return View();
		}


		[HttpPost]
		public IActionResult Importuj(IFormFile file)
		{
		

	
			List<Klient> listaKlientow = new List<Klient>(_appDbContext.Klienci.OrderBy(c => c.id).ToList());
			if (file == null || file.Length <= 0 )
			{
				
				return View("Index" , listaKlientow);
			}
			else
			{

			
			try
			{

					if (file.ContentType.Equals("text/csv"))
					{
						StreamReader reader = new StreamReader(file.OpenReadStream());
						var line = reader.ReadLine();

						while (!reader.EndOfStream)
						{
							line = reader.ReadLine();
							var values = line.Split(',');


							Klient klient = new Klient
							{
								name = values[1],
								surname = values[2],
								pesel = values[3],
								rok_urodzenia = Convert.ToInt32(values[4]),
								plec = Convert.ToInt32(values[5])


							};


							_appDbContext.Klienci.Add(klient);


						}


						reader.Close();
					}
					else
					{

						using (var workbook = new XLWorkbook(file.OpenReadStream()))
						{
							var worksheet = workbook.Worksheet(1); 
							var rows = worksheet.RowsUsed();

							foreach (var row in rows.Skip(1)) 
							{
								Klient klient = new Klient
								{

									name = row.Cell(2).Value.ToString(),
									surname = row.Cell(3).Value.ToString(),
									pesel = row.Cell(4).Value.ToString(),
									rok_urodzenia = (int)row.Cell(5).Value,
									plec = (int)row.Cell(6).Value
									
						
					

						};
			
							_appDbContext.Klienci.Add(klient);

							}
						}


					}



				}
			catch
			{
				
		}
			}

			
		_appDbContext.SaveChanges();

		listaKlientow = new List<Klient>(_appDbContext.Klienci.OrderBy(c => c.id).ToList());
			return View("Index", listaKlientow);
		}


		[HttpPost]
		public IActionResult Eksportuj()
		{
			string selectedValue = Request.Form["grupa"];
		
			
                
			List<Klient> listaKlientow = new List<Klient>(_appDbContext.Klienci.OrderBy(c => c.id).ToList());

			try
			{
				if (selectedValue.Equals("opcja1"))
				{
					var memoryStream = new MemoryStream();




					StreamWriter streamWriter = new StreamWriter(memoryStream, Encoding.UTF8);

					streamWriter.WriteLine("id,Imie,Nazwisko,PESEL,rok_urodzenia,Płeć");

					foreach (var item in listaKlientow)
					{
						streamWriter.WriteLine(
							$"{item.id},{item.name},{item.surname},{item.pesel},{item.rok_urodzenia},{item.plec}"
						);
					}

					streamWriter.Flush();
					memoryStream.Seek(0, SeekOrigin.Begin);

					var bytes = memoryStream.ToArray();

					streamWriter.Close();
					memoryStream.Close();

					var fileName = "Lista.csv";
					var mimeType = "text/csv";


					return File(bytes, mimeType, fileName);
				}
				else
				{
					using (var arkusz = new XLWorkbook())
					{

						var strona = arkusz.Worksheets.Add("Lista klientów");
						var wiersz = 1;
					
						strona.Cell(wiersz, 1).Value = "Id";
						strona.Cell(wiersz, 2).Value = "Imie";
						strona.Cell(wiersz, 3).Value = "Nazwisko";
						strona.Cell(wiersz, 4).Value = "PESEL";
						strona.Cell(wiersz, 5).Value = "Rok urodzenia";
						strona.Cell(wiersz, 6).Value = "Płeć";


						foreach( var  item in listaKlientow)
						{
							wiersz++;
							strona.Cell(wiersz, 1).Value = item.id;
							strona.Cell(wiersz, 2).Value = item.name;
							strona.Cell(wiersz, 3).Value = item.surname;
							strona.Cell(wiersz, 4).Value = item.pesel;
							strona.Cell(wiersz, 5).Value = item.rok_urodzenia;
							strona.Cell(wiersz, 6).Value = item.plec;

						}



						using (var stream = new MemoryStream())
						{
							arkusz.SaveAs(stream);
							var content = stream.ToArray();

							return File(
								content,
								"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
								"ListaKlientow.xlsx");
						}
					}

				}
			}
			catch (Exception ex)
			{
		
				return RedirectToAction("Index"); 
			}
			
			
		}






		[HttpPost]
		public IActionResult EksportujWidok()
		{
			return View("Eksportuj");
		}


	}
}
