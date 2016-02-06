package xlsgenerator;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Types;
import java.util.TreeMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.formula.Formula;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.formula.ptg.RefPtg;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author James Ka
 */
public class XlsGenerator {
    
    static final String version = "0.1.0.1";

    static TreeMap<String,String> properties = new TreeMap<>();
    
    static String inputFileName;
    static String inputName;
    static String templateFileName;
    static String outputFileName;
    
    static FormulaEvaluator evaluator;
    static XSSFEvaluationWorkbook evaluatorWb;
    
    class CellInfo {
        String name;
        int row;
        int col;
        int source;

        public CellInfo(String name, int row, int col) {
            this.name = name;
            this.row = row;
            this.col = col;
        }
        
        public int getKind()
        {
            return 0;
        }
        
    }
    class DataSource {
        String name;
        
        TreeMap<String, CellInfo> cellules = new TreeMap<>();

        public DataSource(String name) {
            this.name = name;
        }
    }
    
    static void copyRow(XSSFSheet ws, int from, int to)
    {
        XSSFRow srcRow = ws.getRow(from);
        
        XSSFRow dstRow = ws.getRow(to);
        if (dstRow==null)
            dstRow = ws.createRow(to);
        dstRow.setHeight(srcRow.getHeight());
        
        XSSFWorkbook wb = ws.getWorkbook();
        for (int N=srcRow.getFirstCellNum(); N<=srcRow.getLastCellNum(); ++N)
        {
            XSSFCell srcCell = srcRow.getCell(N);
            if (srcCell==null) continue;
            
            // Copy style from old cell and apply to new cell
            XSSFCell newCell = dstRow.createCell(N);
            XSSFCellStyle newCellStyle = wb.createCellStyle();
            newCellStyle.cloneStyleFrom(srcCell.getCellStyle());
            
            newCell.setCellStyle(newCellStyle);

            // If there is a cell comment, copy
            if (srcCell.getCellComment() != null) {
                newCell.setCellComment(srcCell.getCellComment());
            }

            // If there is a cell hyperlink, copy
            if (srcCell.getHyperlink() != null) {
                newCell.setHyperlink(srcCell.getHyperlink());
            }

            // Set the cell data type
            newCell.setCellType(srcCell.getCellType());

            // Set the cell data value
            switch (srcCell.getCellType()) {
                case Cell.CELL_TYPE_BLANK:
                    newCell.setCellValue(srcCell.getStringCellValue());
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    newCell.setCellValue(srcCell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_ERROR:
                    newCell.setCellErrorValue(srcCell.getErrorCellValue());
                    break;
                case Cell.CELL_TYPE_FORMULA:
                {
                    Ptg [] ptgs = FormulaParser.parse(srcCell.getCellFormula(), evaluatorWb, FormulaType.CELL, wb.getSheetIndex(ws));
                    for (Ptg ptg : ptgs) {
                        if (ptg instanceof RefPtg && ((RefPtg)ptg).isRowRelative())
                            try {
                                ((RefPtg)ptg).setRow(((RefPtg)ptg).getRow()+(to-from));
                                //System.out.printf("part(%s)=%s\n", ptg.getClass().getName(), ptg.toFormulaString());
                            }
                            catch (Exception ex) {}
                    }
                    Formula formula = Formula.create(ptgs);
                    if (formula!=null) {
                        String txtFormula = FormulaRenderer.toFormulaString(evaluatorWb, ptgs);
                        //System.out.println(txtFormula);
                        newCell.setCellFormula(txtFormula);
                    }
                    else
                        newCell.setCellFormula(srcCell.getCellFormula());
                }
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    newCell.setCellValue(srcCell.getNumericCellValue());
                    break;
                case Cell.CELL_TYPE_STRING:
                    newCell.setCellValue(srcCell.getRichStringCellValue());
                    break;
            }                
        }
    }
    
    /**
     * @param args the command line arguments
     * @throws java.io.FileNotFoundException
     */
    public static void main(String[] args) 
            throws FileNotFoundException, IOException 
    {
        new XlsGenerator().build(args);
    }
    
    /*
    
    ## BDXX_START ##
     ## CDXX_NAME  
     ## RDXX_NAME
    
                                                             ## BDXX_END ##
     */
    
    /**
     * @param args the command line arguments
     */
    public void build(String[] args) throws FileNotFoundException, IOException {
        
        String propName = null;
        for (String arg : args) {
            if (propName!=null) {
                properties.put(propName, arg);
                propName = null;
                continue;
            }
            if (arg.startsWith("-")) {
                propName = arg.substring(1);
            }
        }
        if (properties.containsKey("in"))
            inputFileName = properties.get("in");
        if (properties.containsKey("template"))
            templateFileName = properties.get("template");
        if (properties.containsKey("out"))
            outputFileName = properties.get("out");
        if (properties.containsKey("inname"))
            inputName = properties.get("inname");
        
        if (inputFileName==null || templateFileName==null || outputFileName==null) {
            System.out.printf("java -jar XlsGenerator.jar -in <input> -inname <name> -template <xlsx> -out <xlsx>\n");
            return;
        }
        // ouverture du template
        File excel =  new File (templateFileName);
        try (FileInputStream fis = new FileInputStream(excel);
             XSSFWorkbook wb = new XSSFWorkbook(fis))
        {
            evaluator = wb.getCreationHelper().createFormulaEvaluator();
            //FormulaParser.parse(propName, null, formulaType, sheetIndex);
            evaluatorWb = XSSFEvaluationWorkbook.create(wb);
            
            XSSFSheet ws = wb.getSheetAt(0);
            
            System.out.printf("sheet name: %s\n", ws.getSheetName());
            
            int firstRow = -1;
            TreeMap<String,DataSource> datasources = new TreeMap<>();
            
            System.out.println("-- Names --");
            // -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
            // détection des sources de données
            // -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
            for (int N=0; N<wb.getNumberOfNames(); ++N)
            {
                XSSFName name = wb.getNameAt(N);
                CellReference reference = new CellReference(name.getRefersToFormula());
                String cellName = name.getNameName();
                if (cellName.contains("_")) {
                    String []items = cellName.split("_");
                    if ((items[0].length() > 1 && 
                         (items[0].startsWith("D") || items[0].startsWith("B"))) ||
                        (items[0].length() > 2 && 
                         (items[0].startsWith("CD") ||
                          items[0].startsWith("RD")))) {
                        System.out.print("(*)");
                        DataSource ds;
                        if (null==(ds=datasources.get(items[0])))
                        {
                            ds = new DataSource(items[0]);
                            datasources.put(ds.name, ds);
                        }
                        ds.cellules.put(items[1], new CellInfo(items[1], reference.getRow(), reference.getCol()));
                    }
                }
                System.out.printf("\t%s:%s - row:%d, column:%d, sheet:%s\n", 
                        name.getNameName(), 
                        name.getRefersToFormula(),
                        reference.getRow(),
                        reference.getCol(),
                        reference.getSheetName()
                );
                
                firstRow = reference.getRow();
            }
            System.out.println("-- Names --");
        
            // -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
            // association des champs de la source de données
            // -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
        
            TextReader rslt = new TextReader();
            
            rslt.fillData(properties.get("in"));
            
            ResultSetMetaData rsltMeta = rslt.getMetaData();
            {
                DataSource ds = datasources.get("RDPAYS");
                for (int N=1; N<=rsltMeta.getColumnCount(); ++N)
                {
                    String colName = rsltMeta.getColumnName(N);
                    if (ds!=null) {
                        CellInfo ci;
                        if (null!=(ci=ds.cellules.get(colName)))
                            ci.source = N;  // columnIndex
                    }
                    System.out.println(colName);
                }
            }
            // -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
            // remplissage des sources de données
            // -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
            int initialRow = firstRow;
            int rowIndex = 0;
            while (rslt.next()) {
                if (rowIndex>0) {
                    ws.shiftRows(firstRow, firstRow+2, 1, true, false);
                    
                    copyRow(ws, initialRow, firstRow);
                }
                
                for (DataSource ds : datasources.values()) {
                    for (CellInfo ci : ds.cellules.values()) {
                        if (ci.source<1) continue;
                        XSSFRow row = ws.getRow(firstRow);
                        if (row==null)
                            row = ws.createRow(firstRow);
                        XSSFCell cell = row.getCell(ci.col);
                        if (cell==null)
                            cell = row.createCell(ci.col);
                        //rslt.getObject(ci.source);
                        switch (rsltMeta.getColumnType(ci.source)) {
                            case Types.INTEGER:
                                cell.setCellValue(rslt.getInt(ci.source)); break;
                            case Types.DECIMAL:
                                cell.setCellValue(rslt.getDouble(ci.source)); break;
                            default:
                                cell.setCellValue(rslt.getString(ci.source));
                        }
                    }
                }
                // add 
                ++firstRow;
                ++rowIndex;
            }
            // -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
            // recalcul des cellules
            // -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
            
            evaluator.evaluateAll();
            
            // -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
            // Sauvegarde du document final
            // -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
            try (FileOutputStream output = new FileOutputStream(properties.get("out")))
            {
                wb.write(output);
            }
        } catch (SQLException ex) {
            Logger.getLogger(XlsGenerator.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }
    
}
