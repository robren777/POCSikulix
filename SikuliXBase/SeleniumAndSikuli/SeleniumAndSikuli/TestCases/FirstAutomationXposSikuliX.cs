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
    public class RecargasTAE
    {
        //private IWebDriver webDriver = null;
        private APILauncher launcher = new APILauncher(true);
        //private object session;
        DesktopOptions option = new DesktopOptions();
        Sikuli4Net.sikuli_REST.Screen screen = new Sikuli4Net.sikuli_REST.Screen();
        SqlConnection SqlConn = new SqlConnection();

        [TestMethod]
        public void MTCFT079RecargaTAE()
        {
            launcher.Start();

            option.ApplicationPath = @"C:\Xpos\Xpos.exe";
            string winDriverPath = @"C:\Sikulix";
            string folderPath = @"C:\Sikulix\Images\";
            WiniumDriver winDriver = new WiniumDriver(winDriverPath, option);

            //Insumos de Prueba
            string numCel = "8100000001";
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
            Pattern btnMenuTae = new Pattern(folderPath + "btnMenuTae.png", 0.9);
            Pattern btnTelcel = new Pattern(folderPath + "btnTelcel.png", 0.9);
            Pattern btnSubMenuTelcel = new Pattern(folderPath + "btnSubMenuTelcel.png", 0.9);
            Pattern itemJournalTae = new Pattern(folderPath + "itemJournalTae.png");
            Pattern txtJournal = new Pattern(folderPath + "txtJournal.png");
            Pattern btnSubtotalizar = new Pattern(folderPath + "btnSubtotalizar.png");
            Pattern btnRegresarPromo = new Pattern(folderPath + "btnRegresarPromo.png");
            Pattern msjNoDeseoAcumularPuntos = new Pattern(folderPath + "msjNoDeseoAcumularPuntos.png");
            Pattern btnPagarEfectivo = new Pattern(folderPath + "btnPagarEfectivo.png");
            Pattern txtNumCelularTae = new Pattern(folderPath + "txtNumCelularTae.png");
            Pattern txtNumCelularConfirmarTae = new Pattern(folderPath + "txtNumCelularConfirmarTae.png");
            Pattern btnAceptarYContinuarTae = new Pattern(folderPath + "btnAceptarYContinuarTae.png");
            Pattern screenJournal = new Pattern(folderPath + "screenJournal.png");
            Pattern cajerosTitulo = new Pattern(folderPath + "cajerosTitulo.png", 0.5);
            Pattern cajerosBotones = new Pattern(folderPath + "cajerosBotones.png", 0.9);
            Pattern btnAceptar = new Pattern(folderPath + "btnAceptar.png");
            Pattern txtPassCajero = new Pattern(folderPath + "txtPassCajero.png", 0.8);
            Pattern btnPosponer = new Pattern(folderPath + "btnPosponer.png");

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

            //Se valida exista el botón Recarga tiempo aire
            if (screen.Exists(btnMenuTae, 30))
            {
                screen.Click(btnMenuTae, KeyModifier.NONE, true);
                screen.Wait(btnTelcel, 30);
                screen.Click(btnTelcel, KeyModifier.NONE, true);
                screen.Wait(btnSubMenuTelcel, 30);  
                screen.Click(btnSubMenuTelcel, KeyModifier.NONE, true);
            }
            else
            {
                Assert.Fail("Error, no se encontró el botón del menú TAE");
            }
            
            //Valida agregue el Item al Journal
            if (screen.Exists(itemJournalTae, 70))
            {
                screen.Find(itemJournalTae, true);
            }
            else
            {
                Assert.Fail("Error, no agregó el Item al Journal");
            }

            //Extrae un SKU de BD
            SqlConn.ConnectionString = "Data Source=localhost;Initial Catalog=XposStore;User Id=sa;Password=12345678;";
            SqlConn.Open();
            string skuBD = "";
            string Sqlstr = " SELECT TOP 1 SKU FROM ITEM " +
                            " WHERE ITEM_STATUS_CODE_ID = 33 " +
                            " AND ITEM_TYPE_CODE_ID = 30 " +
                            " AND NAME LIKE 'COCA-COLA%'";
            SqlCommand command = new SqlCommand(Sqlstr, SqlConn);
            SqlDataReader reader = command.ExecuteReader();

            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    skuBD = reader.GetString(0);
                    Console.WriteLine(skuBD);
                }
            }
            else
            {
                Console.WriteLine("No se encontraron resultados en la base de datos.");
            }
            reader.Close();
            SqlConn.Close();

            //Captura el SKU en el Journal
            screen.Type(txtJournal, skuBD);
            AutoItX.Send("{ENTER}");

            Thread.Sleep(10000);
            
            //Subtotaliza Venta
            screen.Wait(btnSubtotalizar, 30);
            screen.Click(btnSubtotalizar, KeyModifier.NONE, true);

            //Cierra Mensajes de Sugerencia de Venta
            screen.Wait(btnRegresarPromo, 30);
            screen.Click(btnRegresarPromo, KeyModifier.NONE, true);

            screen.Wait(msjNoDeseoAcumularPuntos, 30);
            screen.Click(msjNoDeseoAcumularPuntos, KeyModifier.NONE, true);

            //Cierra la venta en efectivo
            screen.Wait(btnPagarEfectivo, 30);
            screen.Click(btnPagarEfectivo, KeyModifier.NONE, true);

            //Captura y Confirma Número de Celular, da clic en Aceptar
            screen.Wait(txtNumCelularTae, 30);
            screen.Type(txtNumCelularTae, numCel + Key.ENTER);
            if (screen.Exists(txtNumCelularConfirmarTae, 7))
            {
                screen.Type(txtNumCelularConfirmarTae, numCel);
            }
            screen.Wait(btnAceptarYContinuarTae, 30);
            Thread.Sleep(5000);
            screen.Click(btnAceptarYContinuarTae, KeyModifier.NONE, true);

            //Extrae el texto del PosPrinter Simulator
            string ticket = winDriver.FindElement(txtPrinter).Text;

            while (ticket.Length==0)
            {
                ticket = winDriver.FindElement(txtPrinter).Text;
            }

            Console.WriteLine("Ticket: " + ticket);
           
            //Encuentra info en el ticket
            if (ticket.Contains("TAE TELCEL"))
            {
                Console.WriteLine("correcto, aparece TAE TELCEL en el ticket");
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
