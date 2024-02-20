using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using Sikuli4Net.sikuli_UTIL;
using Sikuli4Net.sikuli_REST;
using System.Threading;
using OpenQA.Selenium.Winium;
using System;
using System.Diagnostics;
using System.Linq;
using System.Collections.ObjectModel;
using AutoIt;
using Microsoft.SqlServer;
using System.Data.SqlClient;
using OpenQA.Selenium.Support.UI;
using iText.Layout;
using iText.Kernel.Pdf;
using iText.Layout.Element;
using System.Collections.Generic;


namespace SeleniumAndSikuli
{
    [TestClass]
    public class DS50App
    {
        //private IWebDriver webDriver = null;
        private APILauncher launcher = new APILauncher(true);
        //private object session;
        DesktopOptions option = new DesktopOptions();
        Screen screen = new Screen();

        [TestMethod]
        public void MTCFT079DS50App()
        {
            launcher.Start();

            option.ApplicationPath = @"C:\Xpos\Xpos.exe";
            string winDriverPath = @"C:\Sikulix";
            string folderPath = @"C:\Sikulix\Images\";
            WiniumDriver winDriver = new WiniumDriver(winDriverPath, option);

            //Insumos de Prueba
            string userCajero = "RUGAMA8502065";
            string passCajero = "Comercio.1";

            //Objetos de Xpos (Winium Desktop Driver - Inspect.exe)
            WiniumBy btnTAE = WiniumBy.AutomationId("btnTAE");
            WiniumBy txtPrinter = WiniumBy.AutomationId("PrinterOutput");
            WiniumBy btnClearPrinter = WiniumBy.AutomationId("Clear");
            WiniumBy wndPrinter = WiniumBy.Name("Microsoft PosPrinter Simulator");
            WiniumBy listUser = WiniumBy.AutomationId("usersListView");
            WiniumBy contentUseR = WiniumBy.Name("Oxxo.Xpos.Commons.Security.DTO.Info.OperatorInfo");
            WiniumBy usertxt = WiniumBy.ClassName("TextBlock");
            WiniumBy XPOS = WiniumBy.Name("Xpos");
            WiniumBy txtBoxJournal = WiniumBy.XPath("//*[@AutomationId='txtCode' and @ClassName='TextBox' and @IsEnabled=true()]");

            //Objetos de Xpos (Pattern Images for SikuliX)
            Pattern btnIniciarDS50 = new Pattern(folderPath + "btnIniciarDS50.png");
            Pattern btnDetenerDS50 = new Pattern(folderPath + "btnDetenerDS50.png");
            Pattern btnRealizarComunicacionDS50 = new Pattern(folderPath + "btnRealizarComunicacionDS50.png");
            Pattern btnLimpiarLogDS50 = new Pattern(folderPath + "btnLimpiarLogDS50.png");
            Pattern btnConfiguracionDS50 = new Pattern(folderPath + "btnConfiguracionDS50.png");
            Pattern btnSalirDS50 = new Pattern(folderPath + "btnSalirDS50.png");
            Pattern txtJournal = new Pattern(folderPath + "txtJournal.png");
            Pattern btnSubtotalizar = new Pattern(folderPath + "btnSubtotalizar.png");
            Pattern btnRegresarPromo = new Pattern(folderPath + "btnRegresarPromo.png");
            Pattern msjNoDeseoAcumularPuntos = new Pattern(folderPath + "msjNoDeseoAcumularPuntos.png");
            Pattern btnPagarEfectivo = new Pattern(folderPath + "btnPagarEfectivo.png");
            Pattern screenJournal = new Pattern(folderPath + "screenJournal.png");
            Pattern cajerosTitulo = new Pattern(folderPath + "cajerosTitulo.png", 0.5);
            Pattern cajerosBotones = new Pattern(folderPath + "cajerosBotones.png", 0.9);
            Pattern btnAceptar = new Pattern(folderPath + "btnAceptar.png");
            Pattern txtPassCajero = new Pattern(folderPath + "txtPassCajero.png", 0.8);
            Pattern btnPosponer = new Pattern(folderPath + "btnPosponer.png");

            //Desactiva la tecla bloq mayus
            AutoItX.Send("{CAPSLOCK 0}");

            //Obtiene el texto del portapapeles
            var clipg = AutoItX.ClipGet();
            Console.WriteLine(clipg);

            //Activa la ventana de Printer Simulator
            AutoItX.WinActivate("Microsoft PosPrinter Simulator");

            //Borra el texto de la Printer Simulator
            winDriver.FindElement(btnClearPrinter).Click();

            //Regresa el focus a la ventana Xpos
            AutoItX.WinActivate("Xpos");

            //// Ruta donde se guardará el archivo PDF
            //string filePath = @"C:\EvidenciaSikulix\01.pdf";

            //// Crear el documento PDF
            //using (PdfWriter writer = new PdfWriter(filePath))
            //{
            //    using (PdfDocument pdf = new PdfDocument(writer))
            //    {
            //        Document document = new Document(pdf);

            //        // Agregar contenido al documento
            //        document.Add(new Paragraph("Evidencia de Test Case"));
            //        document.Add(new Paragraph("Fecha: " + DateTime.Now.ToString()));

            //        // Guardar y cerrar el documento
            //        document.Close();
            //    }
            //}

            //Console.WriteLine("Informe PDF creado con éxito en: " + filePath);
        
            //Ejemplo de uso de explicit wait
            try
            { 
                WebDriverWait wait = new WebDriverWait(winDriver, TimeSpan.FromSeconds(30));
                wait.Until(x => x.FindElement(txtBoxJournal));
                Console.WriteLine("text habilitado");
            } catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            
            //Inicio de Ejecución
            if (screen.Exists(txtJournal,30) == false)
            {
                Process[] process = Process.GetProcessesByName("Xpos");
                process.AsParallel().ForAll(p => { p.Kill(); p.WaitForExit(); });
                Process.Start(@"C:\Xpos\Xpos.exe");

                //Realiza Login de Cajero para entrar al Journal
                do
                {
                    Thread.Sleep(5000);
                }
                while (screen.Exists(cajerosTitulo, 1) == false && screen.Exists(cajerosBotones, 1) == false);
                screen.Find(cajerosBotones, true);
                screen.Find(cajerosTitulo, true);
                
                ReadOnlyCollection<WiniumElement> listofelements = winDriver.FindElements(contentUseR);
                Console.WriteLine("elementos: " + listofelements.Count());
                for (int i = 0; i < listofelements.Count; i++)
                {
                    String nameCajero = listofelements[i].FindElement(usertxt).GetAttribute("Name").Trim();
                    Console.WriteLine("Cajero:" + nameCajero);
                    if (nameCajero == userCajero)
                    {
                        listofelements[i].FindElement(usertxt).Click();
                        screen.Click(btnAceptar, true);
                        Thread.Sleep(10000);
                        screen.Type(txtPassCajero, passCajero + Key.ENTER);
                        if (screen.Exists(btnPosponer, 20))
                        {
                            screen.Click(btnPosponer, true);
                        }
                        break;
                    }
                }
            }
            else
            {
                Console.WriteLine("Xpos en ejecución, se continua testCase");
                if (screen.Exists(btnPosponer, 10))
                {
                    screen.Click(btnPosponer, true);
                }
            }

            //TEST SCRIPT//

            //Inicia la aplicación DS50
            Process.Start(@"C:\POS\Ds50Tools.exe");

            //Valida se despliegue la ventana
            if (screen.Exists(btnIniciarDS50, 30))
            {
                Console.WriteLine("Se despliega módulo DS50");
            }
            else
            {
                Assert.Fail("Error, no se desplegó el módulo DS50 o tardó demasiado");
            }
            
            //valida existan todos los botones y da clic en botón Salir
            List<Pattern> listbtn = new List<Pattern>() {btnIniciarDS50, btnDetenerDS50, 
                btnRealizarComunicacionDS50, btnLimpiarLogDS50, 
                btnConfiguracionDS50, btnSalirDS50};

            foreach (var item in listbtn)
            {
                if (screen.Exists(item, 30))
                {
                    screen.Find(item,true);
                    Console.WriteLine("correcto, encontró: " + item.ImagePath);
                    if (item == btnSalirDS50)
                    {
                        screen.Click(item);
                    }
                }
                else
                {
                    Assert.Fail("Error, no se encontró el botón: " + item.ImagePath);
                }
            }
            
            //TEST SCRIPT//

            //FIN DE EJECUCIÓN
            winDriver.Quit();
            launcher.Stop();
        }
    }

}
