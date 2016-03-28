package xlsgenerator;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Types;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Map;
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
    
    public static final String version = "0.2.0.1";
    
    /**
     * Public
     */

    public String inputFileName;
    public String inputName;
    public String templateFileName;
    public String outputFileName;
    
    /**
     * Static 
     */

    static FormulaEvaluator evaluator;
    static XSSFEvaluationWorkbook evaluatorWb;
    static TreeMap<String,String> properties = new TreeMap<>();
    
    class DataSourceSql {
        String dsName;
        ResultSet rslt;

        public DataSourceSql(String dsName, ResultSet rslt) {
            this.dsName = dsName;
            this.rslt = rslt;
        }
    }
    
    /**
     * Private
     */
    
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
        int firstFixedRow;
        int firstRow;
        int rowCount;
        
        TreeMap<String, CellInfo> cellules = new TreeMap<>();
        
        //public DataSource() {}

        public DataSource(String name) {
            this.name = name;
            firstRow = Integer.MAX_VALUE;
            firstFixedRow = Integer.MAX_VALUE;
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
                    Ptg [] ptgs = FormulaParser.parse(srcCell.getCellFormula(), 
                                                      evaluatorWb, 
                                                      FormulaType.CELL, 
                                                      wb.getSheetIndex(ws));
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
    
    /*
    
    ## BDXX_START ##
     ## CDXX_NAME  
     ## RDXX_NAME
    
                                                             ## BDXX_END ##
     */
    
    /**
     * @param dataSourcesSql
     * @throws java.io.FileNotFoundException
     */
    public void build(Map<String,DataSourceSql> dataSourcesSql) 
    //public void build(String dsName, ResultSet rslt) 
            throws FileNotFoundException, IOException 
    {
        // ouverture du template
        try (FileInputStream fis = new FileInputStream(templateFileName);
             //XSSFWorkbook wbTemplate = new XSSFWorkbook(fis);
             XSSFWorkbook wb = new XSSFWorkbook(fis))
        {
            evaluator = wb.getCreationHelper().createFormulaEvaluator();
            //FormulaParser.parse(propName, null, formulaType, sheetIndex);
            evaluatorWb = XSSFEvaluationWorkbook.create(wb);
            
            XSSFSheet ws = wb.getSheetAt(0);
            
            System.out.printf("sheet name: %s\n", ws.getSheetName());
            
            //int firstRow = -1;
            int lastRow;
            TreeMap<String,DataSource> datasources = new TreeMap<>();
            //TreeMap<String,DataSource> datasources = new TreeMap<>();
            //TreeMap<String,DataSource> datasources = new TreeMap<>();
            lastRow = ws.getLastRowNum();
            
            analyseNames(wb, datasources);

            // -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
            // association des champs de la source de données
            // -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
            System.out.println("-- DataSources --");
            for (Map.Entry<String,DataSourceSql> e : dataSourcesSql.entrySet()) {
            //for (DataSourceSql dss : dataSourcesSql.values()) {
                DataSourceSql dss = e.getValue();
                ResultSet rslt = dss.rslt;
                String dsName = dss.dsName;

                if (rslt==null) {
                    TextReader textReader = new TextReader();
                    String fileName;
                    String temp = e.getKey();

                    if (temp.contains("="))
                        fileName = temp.substring(temp.indexOf('=')+1);
                    else
                        fileName = inputFileName;
                    textReader.fillData(fileName);

                    dss.rslt = rslt = textReader;
                }
            
                ResultSetMetaData rsltMeta = rslt.getMetaData();
                {
                    //                              vvvvvv
                    DataSource ds = datasources.get(dsName);
                    for (int N=1; N<=rsltMeta.getColumnCount(); ++N)
                    {
                        String colName = rsltMeta.getColumnName(N).toUpperCase();
                        if (ds!=null) {
                            CellInfo ci;
                            if (null!=(ci=ds.cellules.get(colName)))
                                ci.source = N;  // columnIndex
                        }
                        System.out.println(colName);
                    }
                }
            }
            System.out.println("-- DataSources --");
            // -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
            // remplissage des sources de données
            // -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
            // tri des datasource par numéro de lignes décroissant
            ArrayList<DataSource> datasourcesByFirstRow = new ArrayList<>();
            //Collections.copy(datasourcesByFirstRow, datasources.values());
            datasources.values().forEach((DataSource t) -> {
                datasourcesByFirstRow.add(t);
            });
            Collections.sort(datasourcesByFirstRow, (DataSource o1, DataSource o2) -> -(o1.firstRow-o2.firstRow));
            
            for (DataSource ds : datasourcesByFirstRow) {
                DataSourceSql dss = dataSourcesSql.get(ds.name);
                if (dss!=null)
                {
                    int firstRow = ds.firstRow;
                    int initialRow = firstRow;
                    int rowIndex = 0;

                    ResultSet rslt = dss.rslt;
                    ResultSetMetaData rsltMeta = rslt.getMetaData();

                    while (rslt.next()) {

                        // Ajout d'une nouvelle ligne par duplication
                        if (rowIndex>0) {
                            ws.shiftRows(firstRow, lastRow + (firstRow-initialRow), 1, true, false);

                            copyRow(ws, initialRow, firstRow);
                        }
                        for (CellInfo ci : ds.cellules.values()) {
                            if (ci.source<1) continue;
                            XSSFRow row = ws.getRow(firstRow);
                            if (row==null)
                                row = ws.createRow(firstRow);
                            XSSFCell cell = row.getCell(ci.col);
                            if (cell==null)
                                cell = row.createCell(ci.col);
                            // rslt.getObject(ci.source);
                            switch (rsltMeta.getColumnType(ci.source)) {
                                case Types.INTEGER:
                                    cell.setCellValue(rslt.getInt(ci.source)); break;
                                case Types.DECIMAL:
                                    cell.setCellValue(rslt.getDouble(ci.source)); break;
                                case Types.DATE:
                                    cell.setCellValue(rslt.getDate(ci.source)); break;
                                default:
                                    cell.setCellValue(rslt.getString(ci.source));
                            }
                        }
                        // add 
                        ++firstRow;
                        ++rowIndex;
                    }
                }
            }
            // -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
            // recalcul des cellules
            // -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
            
            evaluator.evaluateAll();
            
            // -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
            // Sauvegarde du document final
            // -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
            try (FileOutputStream output = new FileOutputStream(outputFileName))
            {
                wb.write(output);
            }
        } catch (SQLException ex) {
            System.out.println(ex.getMessage());
            Logger.getLogger(XlsGenerator.class.getName()).log(Level.SEVERE, null, ex);
        } catch (Exception ex) {
            System.out.println(ex.getMessage());
            throw ex;
        }
    }

    private void analyseNames(final XSSFWorkbook wb, TreeMap<String, DataSource> datasources) {
        System.out.println("-- Names --");
        // -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
        // détection des sources de données
        // -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
        for (int N=0; N<wb.getNumberOfNames(); ++N)
        {
            XSSFName name = wb.getNameAt(N);
            String cellName = name.getNameName();
            if (!cellName.contains("_")) continue;
            
            CellReference reference = new CellReference(name.getRefersToFormula());
            int firstRow = reference.getRow();
            String []items = cellName.split("_");
            if (items[0].length() > 3 &&
                (items[0].startsWith("CXD") || items[0].startsWith("RXD")))
            {
                System.out.print("(X)");
                DataSource ds;
                String dsName;
                if (null==(ds=datasources.get(dsName=items[0].substring(3))))
                {
                    ds = new DataSource(dsName);
                    datasources.put(ds.name, ds);
                }
                ds.cellules.put(items[1], new CellInfo(items[1], reference.getRow(), reference.getCol()));
                if (ds.firstRow>firstRow)
                    ds.firstRow=firstRow;
            }
            else if (items[0].length() > 2 &&
                (items[0].startsWith("CD") || items[0].startsWith("RD")))
            {
                System.out.print("(*)");
                DataSource ds;
                String dsName;
                if (null==(ds=datasources.get(dsName=items[0].substring(2))))
                {
                    ds = new DataSource(dsName);
                    datasources.put(ds.name, ds);
                }
                ds.cellules.put(items[1], new CellInfo(items[1], reference.getRow(), reference.getCol()));
                if (ds.firstRow>firstRow)
                    ds.firstRow=firstRow;
            }
            else if ((items[0].length() > 1 &&
                (items[0].startsWith("D") || items[0].startsWith("B"))))
            {
                DataSource ds;
                String dsName;
                if (null==(ds=datasources.get(dsName=items[0].substring(1))))
                {
                    ds = new DataSource(dsName);
                    datasources.put(ds.name, ds);
                }
                ds.cellules.put(items[1], new CellInfo(items[1], reference.getRow(), reference.getCol()));
                if (ds.firstFixedRow>firstRow)
                    ds.firstFixedRow=firstRow;
            }
            else continue;
            
            System.out.printf("\t%s:%s - row:%d, column:%d, sheet:%s\n",
                name.getNameName(),
                name.getRefersToFormula(),
                reference.getRow(),
                reference.getCol(),
                reference.getSheetName()
            );
            
        }
        System.out.println("-- End Names --");
    }
    
    /**
     * @param args the command line arguments
     * @throws java.io.FileNotFoundException
     */
    public static void main(String[] args) 
            throws FileNotFoundException, IOException 
    {
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
        
        new XlsGenerator().run(args);
        
        // affiche le classeur généré
        Desktop.getDesktop().open(new File(properties.get("out")));
    }
    public void run(String[] args) 
            throws FileNotFoundException, IOException 
    {
        if (properties.containsKey("template"))
            templateFileName = properties.get("template");
        if (properties.containsKey("out"))
            outputFileName = properties.get("out");
        if (properties.containsKey("in"))
            inputFileName = properties.get("in");
        if (properties.containsKey("inname"))
            inputName = properties.get("inname");
        
        if (inputFileName==null || 
            templateFileName==null || 
            outputFileName==null) {
            
            System.out.printf("java -jar XlsGenerator.jar -in <input> -inname <name> -template <xlsx> -out <xlsx>\n");
            System.out.printf("-template nom du fichier de template\n");
            System.out.printf("-out      nom du document de sortie\n");
            System.out.printf("-inname   nom de la source\n");
            System.out.printf("-in       nom du fichier de données\n");
            return;
        }
        
        Map<String,DataSourceSql> sources = new TreeMap<>();
        
        if (inputName!=null)
            sources.put(inputName, new DataSourceSql(inputName, null));
        for (int i=1; properties.get("inname"+i)!=null; ++i) {
            String inname=properties.get("inname"+i);
            String insrc=properties.get("in"+i);
            if (inname!=null && insrc!=null)
                sources.put(inname+"="+insrc, new DataSourceSql(inname, null));
        }
        
        System.out.println("-- start build --");
        build(sources);
        System.out.println("-- exit build --");
        //generator.build("RDTEST", rslt);
    }
}
