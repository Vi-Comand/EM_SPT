using EM_SPT.Models;
//using Attest.Views; // пространство имен моделей RegisterModel и LoginModel
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Cryptography.KeyDerivation;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace EM_SPT.Controllers
{
    public class AccountController : Controller
    {
        private DataContext db;

        public AccountController(DataContext context)
        {
            db = context;
        }

        [HttpGet]
        public IActionResult Login()
        {
            return View();

        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Login(LoginModel model)
        {
            if (ModelState.IsValid)
            {


                string password = model.Password;

                // generate a 128-bit salt using a secure PRNG
                string a = "Соль";

                byte[] salt = Encoding.Default.GetBytes(a);

                // derive a 256-bit subkey (use HMACSHA1 with 10,000 iterations)
                string hashed = Convert.ToBase64String(KeyDerivation.Pbkdf2(
                    password: password,
                    salt: salt,
                    prf: KeyDerivationPrf.HMACSHA1,
                    iterationCount: 10000,
                    numBytesRequested: 256 / 8));
                string remoteIpAddress = Request.HttpContext.Connection.RemoteIpAddress.ToString();
                user user = new user();
                try
                {
                    user = await db.User.FirstOrDefaultAsync(u => u.login == model.Login && u.pass == hashed);
                }
                catch (Exception ex)
                {
                    ex.Message.ToString();
                }
                if (user != null)
                {


                    await Authenticate(model.Login); // аутентификация
                    if (user.role == 1)
                    { return RedirectToAction("adm_klass", "Home"); }
                    if (user.role == 2)
                    { return RedirectToAction("adm_oo", "Home"); }
                    if (user.role == 3)
                    { return RedirectToAction("adm_mo", "Home"); }
                    if (user.role == 4)
                    { return RedirectToAction("adm_full", "Home"); }
                    else { return RedirectToAction("anketa_a", "Home"); }

                }


                ModelState.AddModelError("", "Некорректные логин и(или) пароль");
            }

            return View(model);
        }


        private async Task Authenticate(string userName)
        {

            //CompositeModel compositeModel=new CompositeModel(db);
            string remoteIpAddress = Request.HttpContext.Connection.RemoteIpAddress.ToString();

            ViewData["Message"] = remoteIpAddress;
            /* if (remoteIpAddress == "193.242.149.177" || remoteIpAddress == "193.242.149.14" || remoteIpAddress == "::1")
             {



     */

            // создаем один claim
            var claims = new List<Claim>
            {
                new Claim(ClaimsIdentity.DefaultNameClaimType, userName)
            };
            // создаем объект ClaimsIdentity
            ClaimsIdentity id = new ClaimsIdentity(claims, "ApplicationCookie", ClaimsIdentity.DefaultNameClaimType,
                ClaimsIdentity.DefaultRoleClaimType);
            // установка аутентификационных куки
            await HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, new ClaimsPrincipal(id));
            //  }





        }



        public async Task<IActionResult> Logout()
        {

            string remoteIpAddress = Request.HttpContext.Connection.RemoteIpAddress.ToString();


            await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
            return RedirectToAction("Index", "Home");
        }
    }
}