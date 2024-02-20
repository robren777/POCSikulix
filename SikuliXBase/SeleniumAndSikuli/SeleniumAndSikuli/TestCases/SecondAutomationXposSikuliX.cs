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


namespace SeleniumAndSikuli
{
    [TestClass]
    public class ServicioEnLineaSKY
    {
        //private IWebDriver webDriver = null;
        private APILauncher launcher = new APILauncher(true);
        //private object session;
        DesktopOptions option = new DesktopOptions();
        Screen screen = new Screen();
        SqlConnection SqlConn = new SqlConnection();

        [TestMethod]
        public void MTCFT079ServicioSKY()
        {
            launcher.Start();

            option.ApplicationPath = @"C:\Xpos\Xpos.exe";
            string winDriverPath = @"C:\Sikulix";
            string folderPath = @"C:\Sikulix\Images\";
            WiniumDriver winDriver = new WiniumDriver(winDriverPath, option);

            //Insumos de Prueba
            string FolioServicio = "123456789102";
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
            Pattern btnServicios = new Pattern(folderPath + "btnServicios.png", 0.9);
            Pattern txtBuscarServicios = new Pattern(folderPath + "txtBuscarServicios.png", 0.9);
            Pattern btnSkyEnLinea = new Pattern(folderPath + "btnSkyEnLinea.png", 0.9);
            Pattern btnSI = new Pattern(folderPath + "btnSI.png", 0.9);
            Pattern txtFolio1 = new Pattern(folderPath + "txtFolio1.png", 0.9);
            Pattern txtFolio2 = new Pattern(folderPath + "txtFolio2.png", 0.9);
            Pattern btnConsulta = new Pattern(folderPath + "btnConsulta.png", 0.9);
            Pattern btnProcesar = new Pattern(folderPath + "btnProcesar.png", 0.9);
            Pattern btnSiAviso = new Pattern(folderPath + "btnSiAviso.png", 0.9);
            Pattern btnNo = new Pattern(folderPath + "btnNo.png", 0.9);
            Pattern txtJournal = new Pattern(folderPath + "txtJournal.png");
            Pattern btnSubtotalizar = new Pattern(folderPath + "btnSubtotalizar.png");
            Pattern btnRegresarPromo = new Pattern(folderPath + "btnRegresarPromo.png");
            Pattern msjNoDeseoAcumularPuntos = new Pattern(folderPath + "msjNoDeseoAcumularPuntos.png");
            Pattern btnPagarEfectivo = new Pattern(folderPath + "btnPagarEfectivo.png");
            Pattern btnAceptarYContinuarTae = new Pattern(folderPath + "btnAceptarYContinuarTae.png");
            Pattern screenJournal = new Pattern(folderPath + "screenJournal.png");
            Pattern cajerosTitulo = new Pattern(folderPath + "cajerosTitulo.png", 0.5);
            Pattern cajerosBotones = new Pattern(folderPath + "cajerosBotones.png", 0.9);
            Pattern btnAceptar = new Pattern(folderPath + "btnAceptar.png");
            Pattern txtPassCajero = new Pattern(folderPath + "txtPassCajero.png", 0.8);
            Pattern btnPosponer = new Pattern(folderPath + "btnPosponer.png");

            //Desactiva la tecla bloq mayus
            AutoItX.Send("{CAPSLOCK 0}");
           
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

            //Se valida exista el botón Servicios
            if (screen.Exists(btnServicios, 30))
            {
                screen.Click(btnServicios, KeyModifier.NONE, true);
            }
            else
            {
                Assert.Fail("Error, no se encontró el botón del menú Servicios");
            }

            //Busca el servicio SKY
            if (screen.Exists(txtBuscarServicios, 30))
            {
                screen.Type(txtBuscarServicios, "Sky En Linea" + Key.ENTER);
                Thread.Sleep(5000);
                screen.Click(btnSkyEnLinea, true);
                Thread.Sleep(5000);
                screen.Click(btnSI, true);
            }
            else
            {
                Assert.Fail("Error, no se encontró el CAMPO para buscar servicios");
            }

            //Teclea el folio del servicio
            if (screen.Exists(txtFolio1, 60))
            {
                AutoItX.Send("{BS}");//tecla backspace
                screen.Type(txtFolio1, FolioServicio);
                AutoItX.Send("{ENTER}"); //tecla ENTER
                Thread.Sleep(5000);
                screen.Type(txtFolio2, FolioServicio + Key.ENTER);
            }
            else
            {
                Assert.Fail("Error, no se encontró el CAMPO para capturar folio");
            }

            //Da click en botón Consulta,Procesar
            if (screen.Exists(btnConsulta, 30))
            {
                screen.Click(btnConsulta, true);
                Thread.Sleep(10000);
                screen.Click(btnProcesar, true);
                Thread.Sleep(8000);
                screen.Click(btnSiAviso, true);
            }
            else
            {
                Assert.Fail("Error, no se encontró el botón Consulta, posible falla de ambiente");
            }


            Thread.Sleep(10000);
            
            //Subtotaliza Venta
            screen.Wait(btnSubtotalizar, 30);
            screen.Click(btnSubtotalizar, KeyModifier.NONE, true);

            //Cierra Mensajes de Sugerencia de Venta
            screen.Wait(btnRegresarPromo, 30);
            screen.Click(btnRegresarPromo, KeyModifier.NONE, true);

            //Mensaje de recarga
            screen.Wait(btnNo, 30);
            screen.Click(btnNo, KeyModifier.NONE, true);

            //ACUMULACION DE PUNTOS
            screen.Wait(msjNoDeseoAcumularPuntos, 30);
            screen.Click(msjNoDeseoAcumularPuntos, KeyModifier.NONE, true);

            //REDONDEO
            screen.Wait(btnNo, 30);
            screen.Click(btnNo, KeyModifier.NONE, true);

            //Cierra la venta en efectivo
            screen.Wait(btnPagarEfectivo, 30);
            screen.Click(btnPagarEfectivo, KeyModifier.NONE, true);

           
            //Extrae el texto del PosPrinter Simulator
            string ticket = winDriver.FindElement(txtPrinter).Text;

            while (ticket.Length==0)
            {
                ticket = winDriver.FindElement(txtPrinter).Text;
            }

            Console.WriteLine("Ticket: " + ticket);
           
            //Encuentra info en el ticket
            if (ticket.Contains("SKY"))
            {
                Console.WriteLine("correcto, aparece SKY en el ticket");
            }
            else
            {
                Assert.Fail("Error, no aparece el texto en el ticket");
            }
           
            //TEST SCRIPT//

            //FIN DE EJECUCIÓN
            winDriver.Quit();
            launcher.Stop();
        }
    }

}
