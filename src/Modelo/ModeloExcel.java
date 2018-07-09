package Modelo;
import java.io.*;
import java.util.*;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import  org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.*;

public class ModeloExcel {

    

    public String Importar(File archivo, JTable tablaE){
        Workbook workbook; // lo meti dentro del metodo para 
        Sheet hoja;//agrega este aca
        //evitar conflicto eso implica que en exportar tambien se creo un tipo de dato workbook
        String respuesta = "No se pudo realizar la importacion";
        DefaultTableModel modeloE = new DefaultTableModel();
        tablaE.setModel(modeloE);
        tablaE.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        
        try {
            //workbook = WorkbookFactory.create(new FileInputStream(archivo));
            FileInputStream file = new FileInputStream(archivo);//simplemente leemos tu archivo y cargamos en memoria
            
            if (archivo.getName().endsWith("xls")) {
                workbook = new HSSFWorkbook(file);
                hoja = workbook.getSheetAt(0);
            }else{
                workbook = new XSSFWorkbook(file);
                hoja = workbook.getSheetAt(0);
            }
            Iterator filaIterator = hoja.rowIterator();
            int indiceFila = -1;

            while (filaIterator.hasNext()) {
                indiceFila++;
                Row fila = (Row) filaIterator.next();
                Iterator columnaIterator = fila.cellIterator(); //Recorre la lista ya creada
                Object[] listaColumna = new Object[1000];//COMO ARREGLO ESTO
                int indiceColumna = -1;

                while (columnaIterator.hasNext()) {
                    indiceColumna++;
                    Cell celda = (Cell) columnaIterator.next();

                    if (indiceFila == 0) {
                        modeloE.addColumn(celda.getStringCellValue());

                    } else {
                        if (celda != null) {
                            switch (celda.getCellType()) {

                                case Cell.CELL_TYPE_NUMERIC:
                                    listaColumna[indiceColumna] = (int) Math.round(celda.getNumericCellValue());
                                    break;

                                case Cell.CELL_TYPE_STRING:
                                    listaColumna[indiceColumna] = celda.getStringCellValue();
                                    break;

                                case Cell.CELL_TYPE_BOOLEAN:
                                    listaColumna[indiceColumna] = celda.getBooleanCellValue();
                                    break;
                                    
                                default:
                                    listaColumna[indiceColumna] = celda.getDateCellValue();
                                    break;
                            }
                            System.out.println("col"+indiceColumna+" valor: true - "+celda+".");
                        }
                    }
                }
                if (indiceFila != 0)modeloE.addRow(listaColumna);
            }
            respuesta = "Importacion exitosa";
        } catch (IOException | EncryptedDocumentException e) {
            System.err.println(e.getMessage());
        }
        return respuesta;
    }
    
    public String Exportar(File archivo, JTable tablaE) {
        Workbook workbook;//este cree aca para que no explote ni probe si funcionaba kp
        String respuesta = "No se realizo la importacion";
        int numFila = tablaE.getRowCount(), numColumna = tablaE.getColumnCount();
        
        if (archivo.getName().endsWith("xls")) {
            workbook = new HSSFWorkbook();
            
        }else{
            workbook = new XSSFWorkbook();
        }
        Sheet hoja = workbook.createSheet();
        try {
            //Para las filas
            for (int i = -1; i < numFila; i++) {
                Row fila = hoja.createRow(i+1);
                
                //Para las columnas
                for (int j = 0; j < numColumna; j++) {
                    Cell celda = fila.createCell(j);
                    
                    if (i == -1) {
                        celda.setCellValue(String.valueOf(tablaE.getColumnName(j)));
                    }else{
                        celda.setCellValue(String.valueOf(tablaE.getValueAt(i, j)));
                    }
                    workbook.write(new FileOutputStream(archivo));
                }
            }
        respuesta = "Exportacion exitosa";  
        } catch (IOException e) {
            System.err.println(e.getMessage());
        }
        return respuesta;
    }   
}
