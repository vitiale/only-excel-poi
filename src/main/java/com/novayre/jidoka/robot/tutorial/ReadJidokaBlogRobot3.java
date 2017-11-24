package com.novayre.jidoka.robot.tutorial;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URI;
import java.net.URL;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import javax.swing.text.html.HTML;

import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.util.CellReference;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

import com.google.common.io.Files;
import com.novayre.jidoka.browser.api.EBrowsers;
import com.novayre.jidoka.browser.api.IWebBrowserSupport;
import com.novayre.jidoka.client.api.IJidokaRobot;
import com.novayre.jidoka.client.api.IJidokaServer;
import com.novayre.jidoka.client.api.IRobot;
import com.novayre.jidoka.client.api.JidokaFactory;
import com.novayre.jidoka.client.api.annotations.Robot;
import com.novayre.jidoka.client.api.exceptions.JidokaException;
import com.novayre.jidoka.client.api.exceptions.JidokaUnsatisfiedConditionException;
import com.novayre.jidoka.data.provider.api.EExcelType;
import com.novayre.jidoka.data.provider.api.IExcel;
import com.novayre.jidoka.windows.api.EShowWindowState;
import com.novayre.jidoka.windows.api.IWindows;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Date;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Reading Jidoka blog robot.
 *
 * @author Jidoka
 *
 */
@Robot
public class ReadJidokaBlogRobot3 implements IRobot {

    /**
     * Default pause between actions imitating human behavior.
     */
    private static final int PAUSE = 500;

    /**
     * URL to navigate to.
     */
    private static final String HOME_URL = "http://blog.jidoka.io";

    /**
     * Firefox browser parameter
     */
    private static final String PARAM_BROWSER_FIREFOX = "firefoxBrowser";

    /**
     * Internet Explorer browser parameter
     */
    private static final String PARAM_BROWSER_IE = "ieBrowser";

    /**
     * Chrome browser parameter
     */
    private static final String PARAM_BROWSER_CHROME = "chromeBrowser";

    /**
     *
     * /**
     * Server.
     */
    private IJidokaServer<?> server;

    /**
     * Windows module.
     */
    private IWindows windows;

    /**
     * Browser module.
     */
    private IWebBrowserSupport browser;

    /**
     * Data-Provider module.
     */
    private IExcel excel;
    private IExcel excel1;
    //para el uso ed apache poi
    private FileOutputStream out;
    private FileInputStream in;
    private File file_poi;
    private File file_poi_lectura;
    private String ruta_local = "C:\\Users\\X220\\Documents\\NetBeansProjects\\ApachePoi\\PaisesIdiomasMonedas.xlsx";
    private String ruta_server_jidoka;
    private static final String NOTEPAD_REGEXP = ".*Bloc de notas";

    /**
     * Current item index. The first one is the number '0'.
     */
    private int currentItemIndex = 0;

    /**
     * Number of items.
     */
    private int numberOfItems = 0;
    
     private String name_input_file;//esta variable va a optener el nombre del archivo dado como parametro que introduce el usuario

    /**
     * File list to upload to the server.
     */
    private List<String> files = new ArrayList<>();

    /**
     * List with the post links.
     */
    private List<String> postLinks = new ArrayList<>();

    private String[] datos_proceasar = new String[]{"Victor Alejandro", "Programador", "Doble O Consulting", "Nuevo Ingreso"};
    private Map<String, Object[]> map = new TreeMap<String, Object[]>();
    private XSSFWorkbook libro;
    private Workbook libro_ambos_formatos;
    private Sheet hoja_ambos_formatos;
    private Row row_ambos_formatos;
    private XSSFSheet hoja;
    private XSSFRow row;
    private String dir = "";

    /**
     * Initial action 'Init'.
     *
     * @throws Exception
     */
    public void init() throws Exception {

        // Initialization of the robot components
        server = (IJidokaServer<?>) JidokaFactory.getServer();
        windows = IJidokaRobot.getInstance(this);

        // Default pause after typing or using the mouse
        windows.typingPause(PAUSE);
        windows.mousePause(PAUSE);
        
        name_input_file=server.getParameters().get("excel"); //va a tomar el nombre del archivo en consola

        ruta_server_jidoka=server.getCurrentDir()+"\\"+name_input_file;
        server.info("direccion completa del fichero "+ruta_server_jidoka);
    }

    /**
     * Action 'Create Excel'.
     *
     * @throws Exception
     */
    public void createExcel() throws Exception {

        /*
		 * Get an instance of IExcel, one of the interfaces included in the
		 * Data-Provider module.
         */
        excel1 = IExcel.getExcelInstance(this);

        // Build the complete path of the file.
        String excel_file = Paths.get(server.getCurrentDir(),
                "salida.xlsx").toString();

        // Create the Excel file from the scratch.
        excel1.create(excel_file, EExcelType.XLSX);

        /*
		 * Add the excel to the list of files to upload to the server when the
		 * robot execution ends.
         */
        files.add(excel_file);

        // Set the value for the first cell.
        excel1.setCellValueByRef(new CellReference(0, 0),
                "Título del post");
        excel1.setCellValueByRef(new CellReference(0, 1), "Subtitulo del post");
    }

    public void mi_create_excel() throws JidokaException, Exception {
        //creando mi excel*************
        String directorio_actual = server.getCurrentDir();
        server.info("El directorio actual es :" + directorio_actual);
        //definiendo nombre unicopara mi excel
        String name = String.valueOf(new Date().getTime()) + "xlsx";
        //cualquiera de las maneras que esta comentada se pueden usar
        File file = Paths.get(directorio_actual, name).toFile();
        //File file = new File(directorio_actual+ "\\" + name);
        server.info("el fichero esta en " + file.getAbsolutePath());
        server.info("dentro de crear excel");
        //IExcel mi_exc=new IExcel
        String excelPath = file.getAbsolutePath();

        excel1 = IExcel.getExcelInstance(this);

        String excel_file = Paths.get(server.getCurrentDir(),
                "out.xlsx").toString();

        excel1.create(excel_file, EExcelType.XLSX);

        if (excel1.createSheet("Hoja1")) {
            server.info("la hoja hoja ha sido creada stisfactoriamente");
            System.out.println("");
        } else {
            server.info("no se creo la hoja en el libro");
        }

        files.add(excel_file);

    }

    public void create_excel_apache_poi() throws JidokaException, FileNotFoundException, IOException {
        String directorio_actual = server.getCurrentDir();
        //File file = Paths.get(directorio_actual, name).toFile();
        String camino = directorio_actual + "\\" + "excel_apache_poi.xlsx";
        server.info("ruta del fichero " + camino);
        file_poi = new File(camino);
        out = new FileOutputStream(file_poi);
        server.setNumberOfItems(datos_proceasar.length);
        libro = new XSSFWorkbook();
        //creamos hoja blanco
        hoja = libro.createSheet("Informacion");
        dir = camino;
        //creamos un objeto row(fila)
        //out.close();
        //files.add(camino);

    }

    public void escribir_excel() {
        //notificando al servidor de lacantidad de items
        server.setNumberOfItems(datos_proceasar.length);
        server.info("editando encabezado");
        server.info("si esta abierto " + excel1.isOpen());
        server.info("si es diferente de null " + excel1.getWorkbook().getSheet("Hoja1") != null);
        if (excel1.isOpen() && excel1.getWorkbook().getSheet("Hoja1") != null) {
            server.info("dentro del if que me permite escribir en el excel");
            Row row = excel1.getSheet().createRow(0);
            row.createCell(0, CellType.STRING).setCellValue("Título");
            row.createCell(1, CellType.STRING).setCellValue("Autor");
            row.createCell(2, CellType.STRING).setCellValue("Precio");
            row.createCell(3, CellType.STRING).setCellValue("Existencia");

        }
    }

    public void escribir_excel_apache_poi() {

        int row_id = currentItemIndex;
        Object[] act = (Object[]) map.get(String.valueOf(row_id + 1));
        //server.info("contenido de act "+act.length);
        row = hoja.createRow(row_id);
        int cell_id = 0;
        for (Object elemento : act) {
            Cell cell = row.createCell(cell_id++);
            cell.setCellValue((String) elemento);
        }
        server.info("contenido de las filas " + act[0] + "       " + act[1] + "       " + act[2]);
    }

    public void abrir_excel_ruta_local() throws FileNotFoundException, IOException {
        file_poi_lectura = new File(ruta_local);
        in = new FileInputStream(file_poi_lectura);
        libro = new XSSFWorkbook(in);
        hoja = libro.getSheetAt(0);
    }

    public void abrir_excel_ruta_local_ambos_formatos() throws FileNotFoundException, IOException, InvalidFormatException {
        file_poi_lectura = new File(ruta_server_jidoka);
        in = new FileInputStream(file_poi_lectura);
//            libro=new XSSFWorkbook(in);
//            hoja=libro.getSheetAt(0);

        libro_ambos_formatos = WorkbookFactory.create(in);
        hoja_ambos_formatos = libro_ambos_formatos.getSheetAt(0);
        numberOfItems = hoja_ambos_formatos.getLastRowNum();
        server.info("tenemos "+numberOfItems+"  filas para leer del archivo");

        server.info("Se ha abierto correctamente el documento");
    }

    public void leer_excel_ruta_local() throws JidokaException, IOException {
        int row_id = 0;
        row = hoja.getRow(row_id++);
        Iterator<Row> iterar_filas = hoja.iterator();
        server.info("INFORMACION POR COLUMNA");
        while (iterar_filas.hasNext()) {
            row = (XSSFRow) iterar_filas.next();
            server.info(row.getCell(0) + "    " + row.getCell(1) + "    " + row.getCell(2) + "    " + row.getCell(3) + "    " + row.getCell(4) + "    " + row.getCell(5) + "    " + row.getCell(6) + "    " + row.getCell(7) + "    " + row.getCell(8) + "    " + row.getCell(9) + "    " + row.getCell(10) + "    " + row.getCell(11) + "    " + row.getCell(12) + "    " + row.getCell(13) + "    " + row.getCell(14));
            Iterator<Cell> iterar_cell = row.cellIterator();
            Cell cell;
            while (iterar_cell.hasNext()) {
                cell = iterar_cell.next();
                //escribiendo resultado de las celdas en el log
                //server.info("iformacion de cada celda "+cell);
            }
        }
        in.close();
    }

    public void leer_excel_ruta_local_ambos_formatos() throws JidokaException, IOException {
        row_ambos_formatos = hoja_ambos_formatos.getRow(currentItemIndex);
        server.info("primera celda de cada fila "+row_ambos_formatos.getCell(0));
        Iterator<Cell> iterar_cell = row_ambos_formatos.cellIterator();
        Cell cell;
        while (iterar_cell.hasNext()) {
            cell = iterar_cell.next();
            //escribiendo resultado de las celdas en el log
            //server.info("iformacion de cada celda "+cell);
        }

        in.close();
    }
    
    public void open_notepad() throws JidokaException, JidokaUnsatisfiedConditionException{
        //Win + r abriendo notepad
        windows.keyboard().windows("r").pause().type("notepad").enter().pause();
        
        //uso de espera inteligente para abrir en este caso notepad
        windows.waitCondition(10, 100, "Wait for notepad to be opened", null, true, (i,c)->windows.getWindow(NOTEPAD_REGEXP)!=null);
        
        //maximazar la ventana
        windows.showWindow(windows.getWindow(NOTEPAD_REGEXP).gethWnd(),	EShowWindowState.SW_MAXIMIZE);
        
        windows.pause();
        
    }
    
    public void write_notepad() throws JidokaException{
        windows.typeText(String.valueOf(row_ambos_formatos.getCell(0))+"  ");
    }
    
    public void safe_notepad() throws Exception{
        windows.typeText(String.valueOf(row_ambos_formatos.getCell(0)));
        windows.typeText(windows.keyboardSequence().pressControl().type("g").releaseControl());
        windows.pause();
        String name_notpad = "ejemplo.txt";
        String path = Paths.get(server.getCurrentDir(), name_notpad).toFile().getAbsolutePath();
        windows.typeText(path);
        windows.typeText(windows.keyboardSequence().typeReturn());
        windows.pause();
        windows.typeText(windows.keyboardSequence().typeAltF(4));
        //windows.typeText(windows.keyboardSequence().pressAlt().type("o"));
        //windows.typeText(windows.keyboardSequence().typeAltF(4));
        windows.pause();
    }

    public void extraer_xlsx_url() {
        //necesito una url para extraer la info en este formato
    }

    /**
     * Action 'Open Browser'.
     *
     * @throws Exception
     */
    public void openBrowser() throws Exception {

        // Get the Browser module instance.
        browser = IWebBrowserSupport.getInstance(this, windows);

        // Set the browser to use
        setBrowser();

        // Initialize the browser.
        browser.initBrowser();
        server.info("inicio el browser");
    }

    /**
     * Sets the browser to use
     *
     * @throws JidokaException
     */
    private void setBrowser() throws JidokaException {

        browser.setBrowserType(EBrowsers.CHROME);
        server.info("selecciono el browser");

//		if (Boolean.parseBoolean(server.getParameters().get(PARAM_BROWSER_FIREFOX))) {
//			browser.setBrowserType(EBrowsers.FIREFOX);
//			return;
//		}
//		if (Boolean.parseBoolean(server.getParameters().get(PARAM_BROWSER_IE))) {
//			browser.setBrowserType(EBrowsers.INTERNET_EXPLORER);
//			return;
//		}
////		if (Boolean.parseBoolean(server.getParameters().get(PARAM_BROWSER_CHROME))) {
////			browser.setBrowserType(EBrowsers.CHROME);
////			return;
////		}
//
//		throw new JidokaException("No browser selected");
    }

    /**
     * Action 'Navigate to Jidoka Blog'.
     *
     * @throws Exception
     */
    public void navigate() throws Exception {

        // Navigate to the home page of the blog.
        browser.navigate(HOME_URL);

        // Wait for the Jidoka Blog page to be loaded 
        checkJidokaBlogLoaded();

        // XPath expression to the button that loads more posts.
        By loadMoreButton = By.cssSelector(".pager_load_more .button_icon .button_label");

        // Count all pages posts.
        while (true) {

            // Post links in the current page.
            List<WebElement> currentPagePostLinks
                    = browser.getElements(By.cssSelector(".post-more"));

            // Number of posts in the current page.
            numberOfItems += currentPagePostLinks.size();

            // Save all post links.
            currentPagePostLinks.forEach(
                    (link) -> postLinks.add(link.getAttribute("href")));

            // Check if there are more posts to load.
            boolean morePosts = windows.waitFor(this).wait(
                    3, "Waiting for more posts to be loaded", false, false,
                    () -> browser.existsElement(loadMoreButton));

            if (!morePosts) {
                // When no more posts, end loop.
                break;
            }

            // Get button to click later to get more posts
            WebElement element = browser.getElement(loadMoreButton);

            browser.moveTo(element);

            windows.clickOnCenter();
            windows.getKeyboard().up(4).pause();

            // Load the next page.
            browser.clickOnElement(loadMoreButton);

            // Pause to allow the browser for updating the DOM.
            windows.pause(5000);
        }

        // Go back to the home page.
        browser.navigate(HOME_URL);

        /*
		 * At this point, all items have been retrieved, so the number of items
		 * can be set.
         */
        server.setNumberOfItems(numberOfItems);
    }

    private void checkJidokaBlogLoaded() throws JidokaUnsatisfiedConditionException {

        boolean newsletterSuscriptionBoxLoaded = windows.getWaitFor(this)
                .wait(15, "Waiting for the blog to be loaded",
                        false, () -> browser.existsElement(By.id("sgcboxClose")));

        WebElement closeSgcBox = browser.getElement(By.id("sgcboxClose"));
        if (newsletterSuscriptionBoxLoaded && closeSgcBox.isDisplayed()) {

            browser.moveTo(closeSgcBox);
            closeSgcBox.click();

            // Wait to close overlay div 
            windows.getWaitFor(this).wait(10, "Waiting for the overlaying div to be closed", true,
                    () -> browser.getElement(By.id("sgcolorbox")).getCssValue("display").equals("none"));

        }

        server.info("Jidoka Blog successfully loaded");
    }

    /**
     * Action 'Extract post'.
     */
    public void extractPost() throws Exception {

        // Get the title.
        String title = browser.getText(By.cssSelector("h1.entry-title"));

        // Notify the server the start of the item process.
        server.setCurrentItem(currentItemIndex, title);

        // Write the title in the Excel, at first column (index 0).
        excel.setCellValueByRef(new CellReference(currentItemIndex, 0), title);

        // Get the image for the current element.
        WebElement image_element = browser.getElement(
                By.cssSelector("img.wp-post-image"));

        // Get the image 'src' attribute escaping special characters
        String image = URI.create(image_element.getAttribute(HTML.Attribute.SRC.toString())).toASCIIString();

        /*
		 * Create a temporary file using the same extension as the specified by
		 * the image URL.
         */
        File file = File.createTempFile("jidoka", String.format(
                "post%s", image.substring(image.lastIndexOf('.'))));

        // Make a copy from the URL to the temporary file.
        try (InputStream is = new URL(image).openStream();
                OutputStream os = new FileOutputStream(file)) {

            IOUtils.copy(is, os);

        } catch (IOException e) {

            /*
			 * In case of an IOException, notify the server to write it to the
			 * execution log.
             */
            server.error(String.format("Error getting image %s", image), e);

            // Notify the server the result is a warning.
            server.setCurrentItemResultToWarn(e.getLocalizedMessage());

            // Continue with the rest of the process.
            return;
        }

        /*
		 * Add the local file containing the image to the list of files to send
		 * to the server when the robot execution ends.
         */
        files.add(file.getAbsolutePath());

        // Notify the server the result is OK.
        server.setCurrentItemResultToOK();
    }

    /**
     * Action 'More?'.
     *
     * @return the wire name
     * @throws Exception
     */
    public String hasMoreItems() throws Exception {

        // Increase index
        currentItemIndex++;

        /*
		 * If the index is greater than the total number of elements, the
		 * process is finished. Otherwise, continue with the next element.
         */
        if (currentItemIndex <= numberOfItems) {
            return "yes";
        }

        // Load the next post.
        //browser.navigate(postLinks.get(currentItemIndex - 1));
        return "no";
    }

    /**
     * Action 'End'.
     *
     * @throws Exception
     */
    public void end() throws Exception {
        // Continue the process. At this step, the robot ends its execution.
    }

    /**
     * @see com.novayre.jidoka.client.api.IRobot#cleanUp()
     */
    @Override
    public String[] cleanUp() throws Exception {

        // Cleaning up browser and excel.
//		browserCleanUp();
//		excelCleanUp();
//		excelCleanUp1();
        excelCleanUp2();

        // Return the file list to upload to the server.
        return files.toArray(new String[files.size()]);
    }

    /**
     * Close the browser.
     */
    private void browserCleanUp() {

        // If the browser was initialized, close it.
        if (browser == null) {
            return;
        }

        browser.close();
        browser = null;
    }

    /**
     * Close Excel.
     */
    private void excelCleanUp() {

        if (excel == null) {
            return;
        }

        try {

            // Save the Excel file
            File target = excel.getFile();
            File tmp = excel.saveToTmp();
            Files.copy(tmp, target);

            // If the Excel was initialized, close it.
            excel.close();

            excel = null;

        } catch (IOException e) {

            // Notify a warning to the server
            server.warn(e.getMessage(), e);
        }
    }

    private void excelCleanUp1() {

        if (excel1 == null) {
            return;
        }

        try {

            // Save the Excel file
            File target = excel1.getFile();
            File tmp = excel1.saveToTmp();
            Files.copy(tmp, target);

            // If the Excel was initialized, close it.
            excel1.close();

            excel1 = null;

        } catch (IOException e) {

            // Notify a warning to the server
            server.warn(e.getMessage(), e);
        }
    }

    private void excelCleanUp2() {
        try {
            libro_ambos_formatos.write(out);
            out.close();
            files.add(dir);
            windows.killAllProcesses("EXCEL.EXE", 500);
            //cerramos tambien el notepad
            windows.killAllProcesses("notepad.exe", 500);
        } catch (IOException ex) {
            // Notify a warning to the server
            server.warn(ex.getMessage(), ex);
        }
    }

}
