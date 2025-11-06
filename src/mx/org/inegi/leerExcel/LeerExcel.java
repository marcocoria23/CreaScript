package mx.org.inegi.leerExcel;

import java.awt.BorderLayout;
import java.awt.Container;
import java.awt.FileDialog;
import java.io.BufferedWriter;
import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.Map;
import javax.swing.BorderFactory;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JProgressBar;
import javax.swing.border.Border;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LeerExcel {

    StringBuilder createSQL, createReferencias;
    String sql;

    public void GeneraScript(String Ruta) throws IOException {

        String filePath = Ruta;
        LeerCelda Cel = new LeerCelda();
        createSQL = new StringBuilder();
        createReferencias = new StringBuilder();
        boolean BGuarArchivo=true; 
        StringBuilder createSQL = new StringBuilder();
        StringBuilder createReferencias = new StringBuilder();

        JFrame f = new JFrame("Progreso Genera Script .xls");
        f.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
        Container content = f.getContentPane();
        JProgressBar progressBar = new JProgressBar();
        progressBar.removeAll();
        progressBar.setValue(0);
        progressBar.setStringPainted(true);
        Border border = BorderFactory.createTitledBorder("Cargando...");
        progressBar.setBorder(border);
        content.add(progressBar, BorderLayout.CENTER);
        f.setSize(300, 100);
        f.setResizable(false);
        f.setLocationRelativeTo(null);
        f.setVisible(true);

        progressBar.setValue(0);
        progressBar.setValue(5);
        try ( FileInputStream fis = new FileInputStream(filePath);  Workbook workbook = new XSSFWorkbook(fis)) { 
            int numberOfSheets = workbook.getNumberOfSheets();
            progressBar.setValue(15);
            for (int i = 0; i < numberOfSheets; i++) {
                String sheetName = workbook.getSheetName(i);

                // Solo procesa las hojas que empiezan con TC_
                if (sheetName.startsWith("TC_")) {
                    Sheet sheet = workbook.getSheet(sheetName);
                    if (sheet != null) {
                        Row headerRow = sheet.getRow(1); // primera fila = encabezados

                        if (headerRow != null) {
                            List<String> encabezados = new ArrayList<>();

                            for (Cell cell : headerRow) {
                                encabezados.add(cell.getStringCellValue());
                            }

                            // === GENERAR CREATE TABLE ===
                            List<String> pkFields = new ArrayList<>();

                            createSQL.append("CREATE TABLE ").append(Cel.LeerCelda(sheet, 0, 0)).append(" (\n");
                            

                            for (int j = 0; j < encabezados.size(); j++) {
                                String colName = encabezados.get(j);
                                if (colName.startsWith("CAT") || colName.startsWith("ID")) {
                                    if (j == 0) {
                                       
                                            createSQL.append("    ").append(colName).append(" NUMBER    ");
                                             
                                          
                                    } else {
                                        createSQL.append("    ").append(colName).append(" NUMBER");
                                    }
                                } else {
                                    createSQL.append("    ").append(colName).append(" VARCHAR2(600)");
                                }
                                if (j < encabezados.size() - 1) {
                                    createSQL.append(",");
                                }
                            }
                            /* CREAR PRIMARY KEY */
                            
                            
                            createSQL.append(",  \n PRIMARY KEY ( ").append(Cel.LeerCelda(sheet, 0, 2)).append(") ");
                                createSQL.append("\n");
                                
                          /*  if (encabezados.size() > 2) {
                                if (!pkFields.isEmpty()) {
                                    createSQL.append("    PRIMARY KEY (");
                                    createSQL.append(String.join(", ", pkFields));
                                    createSQL.append(")\n");
                                }
                            }*/
                            createSQL.append(");\n");

                            // System.out.println(createSQL.toString());
                            // === GENERAR INSERTS ===
                            for (int r = 2; r <= sheet.getLastRowNum(); r++) {
                                Row dataRow = sheet.getRow(r);
                                if (dataRow == null) {
                                    continue;
                                }

                                createSQL.append("INSERT INTO ").append(Cel.LeerCelda(sheet, 0, 0)).append(" (");

                                // Columnas
                                for (int j = 0; j < encabezados.size(); j++) {
                                    createSQL.append(encabezados.get(j));
                                    if (j < encabezados.size() - 1) {
                                        createSQL.append(", ");
                                    }
                                }

                                createSQL.append(") VALUES (");

                                // Valores
                                for (int j = 0; j < encabezados.size(); j++) {
                                    Cell cell = dataRow.getCell(j);
                                    if (cell == null) {
                                        createSQL.append("NULL");
                                    } else {
                                        switch (cell.getCellType()) {
                                            case STRING:
                                                createSQL.append("'").append(cell.getStringCellValue().replace("'", "''")).append("'");
                                                break;
                                            case NUMERIC:
                                                double valor = cell.getNumericCellValue();
                                                if (valor == Math.floor(valor)) {
                                                    createSQL.append((long) valor); // entero
                                                } else {
                                                    createSQL.append(valor); // decimal
                                                }
                                                break;
                                            case BOOLEAN:
                                                createSQL.append(cell.getBooleanCellValue() ? 1 : 0);
                                                break;
                                            default:
                                                createSQL.append("NULL");
                                        }
                                    }

                                    if (j < encabezados.size() - 1) {
                                        createSQL.append(", ");
                                    }
                                }

                                createSQL.append(");\n");
                                // System.out.println(createSQL.toString());
                            }
                        }
                    }
                }
            }
            progressBar.setValue(50);
            //Procesa Script estructura TR   
Map<String, StringBuilder> referenciasPorTabla = new LinkedHashMap<>();

for (int i = 0; i < numberOfSheets; i++) {
    String sheetName = workbook.getSheetName(i);
    if (sheetName.startsWith("TR_")) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet != null) {
            Row headerRow = sheet.getRow(4);
            if (headerRow != null) {

                int colCampo = -1;
                int colTipo = -1;
                int colKey = -1;
                int colObligatorio = -1;
                int colCatalogo = -1;
                int colReferencia = -1;
                int ColNombreCatalogoRef = -1;
                int NombreColumna = 0;

                // 🔹 Buscar las columnas del encabezado
                for (int c = 0; c < headerRow.getLastCellNum(); c++) {
                    String header = Cel.LeerCelda(sheet, 4, c).toUpperCase();
                    if (header.startsWith("CAMPO")) colCampo = c;
                    if (header.startsWith("TIPO")) colTipo = c;
                    if (header.startsWith("KEY")) colKey = c;
                    if (header.startsWith("NULLA")) colObligatorio = c;
                    if (header.equals("CATALOGO")) colCatalogo = c;
                    if (header.startsWith("REFERENCIA")) colReferencia = c;
                    if (header.equals("CATALOGO O TABLA REFERENCIADO")) ColNombreCatalogoRef = c;
                    if (!header.isEmpty()) NombreColumna++;
                }

                if (NombreColumna >= 7) {
                    if (colCampo == -1 || colTipo == -1 || colKey == -1 || colObligatorio == -1) {
                        continue; // Saltar si faltan columnas
                    }

                    // 🔹 Comienza el CREATE TABLE
                    String nombreTabla = Cel.LeerCelda(sheet, 3, 0);
                    createSQL.append("CREATE TABLE ").append(nombreTabla).append(" (\n");

                    List<String> pkFields = new ArrayList<>();
                    StringBuilder referenciasTabla = new StringBuilder();

                    // 🔹 Recorrer las filas de datos
                    for (int r = 5; r <= sheet.getLastRowNum(); r++) {
                        Row dataRow = sheet.getRow(r);
                        if (dataRow == null) continue;

                        String campo = Cel.LeerCelda(sheet, r, colCampo);
                        String tipo = Cel.LeerCelda(sheet, r, colTipo);
                        String key = Cel.LeerCelda(sheet, r, colKey);
                        String obligatorio = Cel.LeerCelda(sheet, r, colObligatorio);

                        // Campos
                        if (obligatorio != null && obligatorio.equalsIgnoreCase("NO")) {
                            if (!campo.isEmpty() && !tipo.isEmpty()) {
                                createSQL.append("    ").append(campo).append(" ").append(tipo).append(" NOT NULL,\n");
                            }
                        } else if (obligatorio != null) {
                            if (!campo.isEmpty() && !tipo.isEmpty()) {
                                createSQL.append("    ").append(campo).append(" ").append(tipo).append(",\n");
                            }
                        }

                        // PK
                        if (key != null && key.equalsIgnoreCase("PK")) {
                            pkFields.add(campo);
                        }

                        // 🔹 Guardar referencias de la tabla
                        if ((colCatalogo != -1) && (colReferencia != -1) && (ColNombreCatalogoRef != -1)) {
                            String Ref = Cel.LeerCelda(sheet, r, colCatalogo);
                            String CampoRef = Cel.LeerCelda(sheet, r, colReferencia);
                            String ColCatRef = Cel.LeerCelda(sheet, r, ColNombreCatalogoRef);

                            if (Ref != null && Ref.equalsIgnoreCase("SI")) {
                                referenciasTabla.append("ALTER TABLE ").append(nombreTabla)
                                        .append(" ADD CONSTRAINT FK")
                                        .append(nombreTabla.replace("TR_", "").replace("_", ""))
                                        .append("_").append(campo)
                                        .append(" FOREIGN KEY (").append(campo).append(") ")
                                        .append("REFERENCES ").append(ColCatRef)
                                        .append("(").append(CampoRef).append(");\n");
                            }
                        }

                        if (key != null && key.equalsIgnoreCase("TR-FK")) {
                            String CampoRef = Cel.LeerCelda(sheet, r, colReferencia);
                            String ColCatRef = Cel.LeerCelda(sheet, r, ColNombreCatalogoRef);
                            referenciasTabla.append("ALTER TABLE ").append(nombreTabla)
                                    .append(" ADD CONSTRAINT FK")
                                    .append(nombreTabla.replace("TR_", "").replace("_", ""))
                                    .append("_").append(campo)
                                    .append(" FOREIGN KEY (").append(CampoRef).append(") ")
                                    .append("REFERENCES ").append(ColCatRef)
                                    .append("(").append(CampoRef).append(");\n");
                        }
                    }

                    // 🔹 Agregar PRIMARY KEY si existe
                    if (!pkFields.isEmpty()) {
                        createSQL.append("    PRIMARY KEY (")
                                 .append(String.join(", ", pkFields))
                                 .append(")\n");
                    } else {
                        int lastIndex = createSQL.lastIndexOf(",");
                        if (lastIndex != -1) createSQL.deleteCharAt(lastIndex);
                    }

                    createSQL.append(");\n\n");

                    // 🔹 Guardar las referencias en el mapa (si tiene)
                    if (referenciasTabla.length() > 0) {
                        referenciasPorTabla.put(nombreTabla, referenciasTabla);
                    }

                } else {
                    JOptionPane.showMessageDialog(null, "Verifique encabezado en la pestaña " + sheetName);
                    BGuarArchivo = false;
                    break;
                }
            }
        }
    }
}

// 🔹 Sección final de FOREIGN KEYS agrupadas
createSQL.append("-- =============================================\n");
createSQL.append("-- FOREIGN KEYS (REFERENCIAS)\n");
createSQL.append("-- =============================================\n\n");

for (Map.Entry<String, StringBuilder> entry : referenciasPorTabla.entrySet()) {
    createSQL.append("-- REFERENCIAS PARA ").append(entry.getKey()).append("\n");
    createSQL.append(entry.getValue()).append("\n");
}
        progressBar.setValue(100);
        f.setVisible(false);    
            
        } catch (IOException e) {
            e.printStackTrace();
        }
        
        if (BGuarArchivo==true)
        {
        SaveSQLToFile(createSQL.toString());    
        }
        
    }

    public static void SaveSQLToFile(String createSQL) throws IOException {
        FileDialog d = new FileDialog(new JFrame(), "Guardar archivo SQL", FileDialog.SAVE);
        d.setFile("ScriptBD.sql"); // nombre sugerido
        d.setVisible(true);

        String dir = d.getDirectory();
        String fileName = d.getFile();

        if (dir != null && fileName != null) {
            // aseguramos extensión .sql
            File sqlFile = new File(dir + fileName);
            if (!sqlFile.getName().toLowerCase().endsWith(".sql")) {
                sqlFile = new File(sqlFile.getAbsolutePath() + ".sql");
            }

            try ( FileWriter fw = new FileWriter(sqlFile);  BufferedWriter bw = new BufferedWriter(fw)) {

                bw.write(createSQL);
                bw.flush();
            }

            JOptionPane.showMessageDialog(null,
                    "Archivo SQL guardado en:\n" + sqlFile.getAbsolutePath(),
                    "Éxito",
                    JOptionPane.INFORMATION_MESSAGE);
        } else {
            JOptionPane.showMessageDialog(null,
                    "No se seleccionó archivo.",
                    "Aviso",
                    JOptionPane.WARNING_MESSAGE);
        }
    }

}
