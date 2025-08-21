/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package mx.org.inegi.leerExcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 *
 * @author ANTONIO.CORIA
 */
public class LeerCelda {
    String Valor="";
    
    public String LeerCelda(Sheet sheet,int NF,int NC ){
       
        Row row = sheet.getRow(NF);
       Cell cell = row.getCell(NC);     
        if (cell == null) return "";

        CellType tipo = cell.getCellType();

        switch (tipo) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                double valor = cell.getNumericCellValue();
                // si es entero, quitar decimales
                if (valor == Math.floor(valor)) {
                    return String.valueOf((long) valor);
                } else {
                    return String.valueOf(valor);
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                // puedes evaluar la fórmula si quieres, aquí se devuelve como String
                return cell.getCellFormula();
            case BLANK:
            default:
                return "";
        }
    }
    
}
