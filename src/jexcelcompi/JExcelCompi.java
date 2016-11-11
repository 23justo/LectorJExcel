 /*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package jexcelcompi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;


/**
 *
 * @author justo
 */
public class JExcelCompi {

    /**
     * @param args the command line arguments
     */
    public  static void main(String[] args) {     
        String nombre;
        nombre = "hoja2.xls";
         celdaUnitaria(nombre,4,2);
        // TODO code application logic here
    }
    
    //extrae valor numerico de celdas del tipo hoja1[fila,columna]
     static void celdaUnitaria(String nombre,int fila,int columna)
    {
        try {
            FileInputStream fis = new FileInputStream(new File(nombre));
            HSSFWorkbook wb = new HSSFWorkbook(fis);
            HSSFSheet sheet = wb.getSheetAt(0); 
            System.out.println(wb.getSheetAt(0).getRow(fila).getCell(columna));
            
        } catch (IOException ex) {
            Logger.getLogger(JExcelCompi.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }
    
    
    public void ejemplo()throws FileNotFoundException, IOException 
    {
        FileInputStream fis = new FileInputStream(new File("hoja2.xls"));
        HSSFWorkbook wb = new HSSFWorkbook(fis);
        HSSFSheet sheet = wb.getSheetAt(0);
        FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
        int x=0;
        for(Row row : sheet){
            for(Cell cell : row){
                switch(formulaEvaluator.evaluateInCell(cell).getCellType())
                {
                    case Cell.CELL_TYPE_NUMERIC:
                    {
                        System.out.println(cell.getNumericCellValue()+"\t\t");
                        
                    }
                    //case Cell.CELL_TYPE_STRING:
                        //System.out.println(cell.getStringCellValue()+"\t\tcadena");    
                      x++;  
                }
                
                
            }
            System.out.println();
        }
        
    }
    
}
