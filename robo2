import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.ss.util.*;
import java.io.*;
import com.eviware.soapui.support.GroovyUtils;
import groovy.util.XmlParser;
import groovy.util.XmlSlurper;
import groovy.json.JsonSlurper;



  // Clase de ExcelReader que contiene una funcion "readData", para leer datos del archivo Excel "--".

  class ExcelReader{

  	def readData(){

  		// Ruta del excel

  		def path = "C:\\Datos\\Datos_Entrada_Ecommerce.xlsx";
  		InputStream inputStream = new FileInputStream(path);
  		Workbook workbook = WorkbookFactory.create(inputStream);
  		Sheet sheet = workbook.getSheetAt(0);

  		// Carga los datos de tipo String del Excel

  		Iterator rowIterator = sheet.rowIterator();
  		rowIterator.next()
  		Row row;

  		def rowsData = []

  		while(rowIterator.hasNext()){

  			row = rowIterator.next()
  			def rowIndex = row.getRowNum()
  			def colIndex;
  			def rowData = []

  			for(Cell cell : row){
  				colIndex = cell.getColumnIndex()
  				rowData[colIndex] = cell.getRichStringCellValue().getString();
  				
  			}

  			rowsData << rowData
  		}

  		rowsData
  	}
  }

  		def groovyUtils = new com.eviware.soapui.support.GroovyUtils(context)
  		def myTestCase = context.testCase
        
  		ExcelReader  excelReader =  new ExcelReader();
  		
  		List rows = excelReader.readData();
  		def d = []
    		Date date = new Date()
     	def newDate = date.format("YYYY-MM-dd-HH-mm-ss.SSS")               	 
     	def subdir = new File('C:\\Datos\\LOG')
     	subdir.mkdir();
  		Iterator i = rows.iterator();
 		// create a new file
  		FileOutputStream out = new FileOutputStream("C:\\Datos\\LOG\\DT-SER-01 - "+newDate+".xlsx");
  		// create a new workbook
  		Workbook wb = new XSSFWorkbook(); 
		// Pinta los datos en el archivo Excel
		/*rownum++
		Sheet s = wb.createSheet(name = "Ejecución_"+rownum);  */
		Sheet s = wb.createSheet(name = "Ejecución"+newDate);

		XSSFCellStyle my_style = wb.createCellStyle();
		XSSFFont my_font = wb.createFont();
		my_font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
		my_style.setFont(my_font);
		int rownum = 0
          int rownum1 = 1		
		int rownum2 = 1
		int rownum3 = 2
		Row r = s.createRow(rownum1)                                          
		Cell c0 = r.createCell(0);
		c0.setCellValue("# Ejecución");
		c0.setCellStyle(my_style);
		
		Cell c1 = r.createCell(1);
		c1.setCellValue(" Request: ");
		c1.setCellStyle(my_style);

		//Row r1 = s.createRow(rownum1)
		Cell c2 = r.createCell(2);
		c2.setCellValue(" Response: ");
		c2.setCellStyle(my_style);

		//Row r9 = s.createRow(rownum1)
		Cell c4 = r.createCell(3);
		c4.setCellValue(" Resultado:");
		c4.setCellStyle(my_style);

		Cell c5 = r.createCell(4);
		c5.setCellValue(" Tiempo respuesta m/s:");
		c5.setCellStyle(my_style);
		
               while(i.hasNext()){
               	 d = i.next();
               	 myTestCase.setPropertyValue("tipoDocumento",d[0])
               	 myTestCase.setPropertyValue("numeroDocumento",d[1])



               	 // runTestStepByName se utiliza para ejecutar el caso de prueba "REST Request".

               	 testRunner.runTestStepByName("DT-SER-01")
               	 def TimeResponse = testRunner.testCase.testSteps ["DT-SER-01"]. testRequest.response.timeTaken; 
                     // log.info response;
               	 def xmlString =
				 'Cabecera:'+"\n"+
               	 'Name '+'  Value '+"\n"+
               	 'tipoDocumento: '+d[0]+' '+"\n"+
               	 'numeroDocumento: '+d[1]+"\n"
				
               	 try{

                     //Muestra el Response de la ejecución. 
               	 def res = context.expand('${DT-SER-01#Response}')
                     //def parsedJson = new groovy.json.JsonSlurper().parseText(res)
                     //def pr = parsedJson.status.statusCode;                                         

                     //log.info pr
               	 // Valida si el caso es exitoso o fallido.               	
               	 if(res == '[]' || res == 'Bad Request' || res == 'Not Found' || res == 'Request failed.'
               	 || res == '' || res == 'No existe la cabecera headerRQ'){

               	 	result = 'Fallido'
               	 }else{

               	 	result = 'Exitoso'
               	 }
                     // Pinta los datos en el archivo Excel
                     /*rownum++
                     Sheet s = wb.createSheet(name = "Ejecución_"+rownum);  

                     XSSFCellStyle my_style = wb.createCellStyle();
                     XSSFFont my_font = wb.createFont();
                     my_font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
                     my_style.setFont(my_font);*/         
				rownum1++
				rownum++	
				Row r11 = s.createRow(rownum1)
      			Cell c22 = r11.createCell(0);
      			c22.setCellValue(rownum);
				
      			//Row r11 = s.createRow(rownum1)
      			Cell c23 = r11.createCell(1);
      			c23.setCellValue(xmlString);
                 			
      			//rownum1++
      			//Row r2 = s.createRow(rownum1)
      			Cell c3 = r11.createCell(2);
      			c3.setCellValue(re);
      			//s.autoSizeColumn(1);
								
      			//Row r5 = s.createRow(rownum1)
      			Cell c6 = r11.createCell(3);
      			c6.setCellValue(result);

      			Cell c7 = r11.createCell(4);
      			c7.setCellValue(TimeResponse);

                     log.info "Request: "+xmlString+ "\n";
               	 log.info "Response: "+res+ "\n";
               	 log.info "Resultado:"+result+"\n";

               	 }
               	 catch(Exception expObj)
               	 {
               	 	//  Exception Handler
               	def res1 = context.expand('${DT-SER-01#Response}')
               	def TimeResponse1 = testRunner.testCase.testSteps ["DT-SER-01"]. testRequest.response.timeTaken;
                    	
				Row r11 = s.createRow(rownum1)
      			Cell c22 = r11.createCell(0);
      			c22.setCellValue(rownum);
				
      			//Row r11 = s.createRow(rownum1)
      			Cell c23 = r11.createCell(1);
      			c23.setCellValue(xmlString);
                 			
      			//rownum1++
      			//Row r2 = s.createRow(rownum1)
      			Cell c3 = r11.createCell(2);
      			c3.setCellValue(res1);
      			//s.autoSizeColumn(1);
								
      			//Row r5 = s.createRow(rownum1)
      			Cell c6 = r11.createCell(3);
      			c6.setCellValue("Error:" + expObj);

      			Cell c7 = r11.createCell(4);
      			c7.setCellValue(TimeResponse1);
               	 	
               	 	log.info "Request: "+xmlString+ "\n";
               	 	log.info "Resultado:"+expObj+"Error 01: Se presentó fallo en el servicio "+"\n";
               	 }               	 
               }
               wb.write(out);
               out.close();
