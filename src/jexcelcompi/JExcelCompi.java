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
        String nombre="hoja.xls,hoja2.xls,hoja3.xls";
        System.out.println("resultado "+menor(nombre, 4, 2));
        
        
        
         
        // TODO code application logic here
    }
    
    //extrae valor numerico de celdas del tipo hoja1[fila,columna]
     static int celdaUnitaria(String nombre,int fila,int columna)
    {
        int valor=0;
        try {
            FileInputStream fis = new FileInputStream(new File(nombre));
            HSSFWorkbook wb = new HSSFWorkbook(fis);
            HSSFSheet sheet = wb.getSheetAt(0); 
            System.out.println("dato "+wb.getSheetAt(0).getRow(fila).getCell(columna));
            valor = (int)(double)(wb.getSheetAt(0).getRow(fila).getCell(columna).getNumericCellValue());
        } catch (IOException ex) {
            Logger.getLogger(JExcelCompi.class.getName()).log(Level.SEVERE, null, ex);
        }
        return valor;
    }
     
     /*saca el valor de la celda especifica que se encuentra en la hoja numHoja*/
     static void HojaNceldaUnitaria(String nombre,int numHoja,int fila,int columna)
    {
        try {
            FileInputStream fis = new FileInputStream(new File(nombre));
            HSSFWorkbook wb = new HSSFWorkbook(fis);
            HSSFSheet sheet = wb.getSheetAt(numHoja); 
            System.out.println(wb.getSheetAt(numHoja).getRow(fila).getCell(columna));
            
        } catch (IOException ex) {
            Logger.getLogger(JExcelCompi.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }
     
     static double suma(String nombre,int fila,int columna)
     {
         String nombres[] = nombre.split(",");
        int totalSuma=0;
        
        for(int i =0 ; i < nombres.length ;i++)
            {
                totalSuma = totalSuma + celdaUnitaria(nombres[i], fila, columna);
            }       
         
         return totalSuma;
     }
     
     static double resta(String nombre,int fila,int columna)
     {
         String nombres[] = nombre.split(",");
        int totalResta=0;
        totalResta = celdaUnitaria(nombres[0], fila, columna);
        for(int i =1 ; i < nombres.length ;i++)
            {
                totalResta = totalResta - celdaUnitaria(nombres[i], fila, columna);
            }       
         
         return totalResta;
     }
     
    static double promedio(String nombre,int fila,int columna)
     {
         String nombres[] = nombre.split(",");
        int total=0;
        int i =0;
        for(i=0 ; i < nombres.length ;i++)
            {
                total = total + celdaUnitaria(nombres[i], fila, columna);
            }       
         total = total/i;
         return total;
     }
    
    static double mayor(String nombre,int fila,int columna)
     {
         String nombres[] = nombre.split(",");
        int mayor=0;
        int i =0;
        for(i=0 ; i < nombres.length ;i++)
            {
                if(celdaUnitaria(nombres[i], fila, columna)>mayor)
                    mayor = celdaUnitaria(nombres[i], fila, columna);
                
            }       
         
         return mayor;
     }
    
    static double menor(String nombre,int fila,int columna)
     {
         String nombres[] = nombre.split(",");
        int menor=celdaUnitaria(nombres[0], fila, columna);
        int i =0;
        for(i=0 ; i < nombres.length ;i++)
            {
                if(celdaUnitaria(nombres[i], fila, columna)<menor)
                    menor = celdaUnitaria(nombres[i], fila, columna);
                
            }       
         
         return menor;
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
