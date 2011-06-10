/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package uk.co.brenn.metadata;

import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.exec.ExecutionContext;
import at.jku.xlwrap.spreadsheet.Cell;
import at.jku.xlwrap.spreadsheet.Sheet;
import at.jku.xlwrap.spreadsheet.TypeAnnotation;
import at.jku.xlwrap.spreadsheet.Workbook;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Name;

/**
 *
 * @author Christian
 */
public class MetaDataCreator {

    static private String MAIN_ROOT = "c:Dropbox/FishLink XLWrap data/";
    static private String MASTER_FILE = "data/MetaMaster.xlsx";
    static private String LIST_SHEET = "Lists";

    private String metaMaster;
    private String metaRoot;
    private String dataRoot;
    
    static private int CATEGORY_ROW = 1;
    static private int FIELD_ROW = 2;
    private Sheet Sheet;

    public MetaDataCreator(){
    }

    public MetaDataCreator(String metaMasterPath, String metaDir, String dataDir){
        metaMaster = metaMasterPath;
        metaRoot = metaDir;
        dataRoot = dataDir;
    }

    private void addMetaDataSheet(CYAB_Workbook dataWorkbook, String dataFile, String doi){
        CYAB_Sheet sheet = dataWorkbook.getSheet("MetaData");
        sheet.setValue("A", 1, "File");
        sheet.setValue("A", 2, dataFile);
        sheet.setValue("B", 1, "Doi");
        sheet.setValue("B", 2, doi);
    }

    private void writeMeta(CYAB_Workbook metaWorkbook, String dataName) throws FileNotFoundException, IOException{
        String fileFront = dataName.substring(0, dataName.lastIndexOf("."));
        metaWorkbook.write(metaRoot + fileFront + "MetaData.xls");
    }

    private void createNamedRanges (Workbook masterWorkbook, CYAB_Workbook metaWorkbook)
            throws XLWrapException, XLWrapEOFException, JavaToExcelException{
        Sheet masterSheet =  masterWorkbook.getSheet(LIST_SHEET);
        CYAB_Sheet metaSheet = metaWorkbook.getSheet(LIST_SHEET);
        int zeroColumn = 0;
        String rangeName =  getTextZeroBased(masterSheet, zeroColumn, 0);
        int categories = -1;
        while (!rangeName.isEmpty()) {
            String columnName = POI_Utils.indexToAlpha(zeroColumn);
            metaSheet.setValueZeroBased(zeroColumn, 0, rangeName);
            int zeroRow = 1;
            String fieldName = getTextZeroBased(masterSheet, zeroColumn, zeroRow);
            while (!fieldName.isEmpty()) {
                metaSheet.setValueZeroBased(zeroColumn, zeroRow, fieldName);
                zeroRow++;
                fieldName = getTextZeroBased(masterSheet, zeroColumn, zeroRow);
            } while (!fieldName.isEmpty());
            Name range = metaWorkbook.createName();
            range.setNameName(rangeName);
            String rangeDef = "'" + LIST_SHEET + "'!$" + columnName + "$2:$" + columnName + "$" + zeroRow;
            //=OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)
            //String rangeDef = "OFFSET('" + LIST_SHEET + "'!$" + columnName + "$1,0,0,COUNTA('" + LIST_SHEET + "'!$" +
            //        columnName + ":$" + columnName + ")-1,1)";
            //ystem.out.println (rangeName + " = " + rangeDef);
            range.setRefersToFormula(rangeDef);
            zeroColumn++;
            rangeName = getTextZeroBased(masterSheet, zeroColumn, 0);
            if (rangeName.isEmpty() && categories == -1){
                //ystem.out.println("doing categories");
                //found the space between Categories and the other nameRanges
                categories = zeroColumn -1;
                zeroColumn++;
                rangeName = getTextZeroBased(masterSheet, zeroColumn, 0);
            }
        }
        Name range = metaWorkbook.createName();
        range.setNameName("Category");
        String columnName = POI_Utils.indexToAlpha(categories);
        range.setRefersToFormula("'" + LIST_SHEET + "'!$A1:$" + columnName + "$1");
    }

    private String getTextZeroBased(Sheet sheet, int column, int row) throws XLWrapException, XLWrapEOFException{
        if (column >= sheet.getColumns()){
            return "";
        }
        if (row >= sheet.getRows()){
            return "";
        }
        Cell cell = sheet.getCell(column, row);
        return cell.getText();
    }

    private int prepareColumnA(Sheet masterSheet, CYAB_Sheet metaSheet, Sheet dataSheet) throws XLWrapException, XLWrapEOFException{
        int zeroRow = 0;
        String value;
        do {
            value = getTextZeroBased(masterSheet, 0, zeroRow);
            metaSheet.setValue("A",zeroRow + 1, value);
            zeroRow++;
            //ystem.out.println (value);
        } while (!value.isEmpty());
        metaSheet.autoSizeColumn(0);
        int maxRow = dataSheet.getRows();
        //ystem.out.println(maxRow);
        if (maxRow > 100){
            maxRow = 100;
        }
        metaSheet.setValue("A",zeroRow + 1, "Column in Data File");
        metaSheet.setValue("A",zeroRow + 2, "Header");
        for (int row = 2; row <= maxRow; row++){
           metaSheet.setValue("A",zeroRow + row + 1, row);
        }
        return zeroRow - 1;
    }

    private void prepareDropDowns(Sheet masterSheet, int lastMetaRow, CYAB_Sheet metaSheet, String column)
            throws JavaToExcelException, XLWrapException, XLWrapEOFException{
        for (int zeroRow = 0; zeroRow < lastMetaRow; zeroRow ++) {
            String list = getTextZeroBased(masterSheet, 2, zeroRow);
            //ystem.out.println(column + ":" + list);
            String popupTitle = getTextZeroBased(masterSheet, 3, zeroRow);
            String popupMessage =  getTextZeroBased(masterSheet, 4, zeroRow);
            String errorStyleString =  getTextZeroBased(masterSheet, 5, zeroRow);
            int errorStyle = DataValidation.ErrorStyle.STOP;
            if (errorStyleString.startsWith("W")){
                errorStyle = DataValidation.ErrorStyle.WARNING;
            }
            if (errorStyleString.startsWith("I")){
                errorStyle = DataValidation.ErrorStyle.INFO;
            }
            String errorTitle =  getTextZeroBased(masterSheet, 6, zeroRow);
            String errorMessage =  getTextZeroBased(masterSheet, 7, zeroRow);
            metaSheet.addValidation(column, zeroRow + 1, list, popupTitle, popupMessage,
                    errorStyle, errorTitle, errorMessage);
       // metaSheet.addListValidation(column, FIELD_ROW,  "INDIRECT(SUBSTITUTE($" + column + "$1,\" \",\"_\"))",
       //         "Type of feild", "Please select the feild that data in this column belongs to");
        }
    }

    private void copyData(int letterRow, CYAB_Sheet metaSheet, Sheet dataSheet, String metaColumn)
            throws XLWrapException{
        int zeroColumn = POI_Utils.alphaToIndex(metaColumn) - 1; //-1 as Column A of Data goes in B of Meta
        String dataColumn = POI_Utils.indexToAlpha(zeroColumn);
        System.out.println("copying from "+ dataColumn + " to " + metaColumn);
        metaSheet.setValue(metaColumn, letterRow, dataColumn);
        int maxRow = dataSheet.getRows();
        if (maxRow > 100){
            maxRow = 100;
        }
        for (int zeroRow = 0; zeroRow < maxRow; zeroRow++){
            //ystem.out.println(zeroRow);
            try {
                Cell cell = dataSheet.getCell(zeroColumn, zeroRow);
                TypeAnnotation typeAnnotation = cell.getType();
                switch (typeAnnotation){
                    case BOOLEAN:
                        boolean booleanValue = cell.getBoolean();
                        //ystem.out.println(booleanValue);
                        metaSheet.setValue(metaColumn, letterRow + zeroRow + 1, booleanValue);
                        break;
                    case NUMBER:
                        double doubleValue = cell.getNumber();
                        //ystem.out.println("number" + doubleValue);
                        metaSheet.setValue(metaColumn, letterRow + zeroRow + 1, doubleValue);
                        break;
                    case TEXT:
                        String textValue = cell.getText();
                        //ystem.out.println(textValue);
                        metaSheet.setValue(metaColumn, letterRow + zeroRow + 1, textValue);
                        break;
                    case DATE:
                        Date dateValue = cell.getDate();
                        String format = cell.getDateFormat();
                       //ystem.out.println(dateValue);
                        metaSheet.setValue(metaColumn, letterRow + zeroRow + 1, dateValue, format);
                        break;
                    case NULL:
                        //ystem.out.println("null");
                        //Moving to copy
                        break;
                    default:
                        throw new XLWrapException("Unexpected Cell Type");
                }
            } catch (XLWrapEOFException ex) {
                //ystem.out.println("Exception");
            }
        }
    }

    private void prepareSheet(Sheet masterSheet, CYAB_Sheet metaSheet, Sheet dataSheet) 
            throws JavaToExcelException, XLWrapException, XLWrapEOFException{
        int lastMetaRow = prepareColumnA(masterSheet, metaSheet, dataSheet);
        int lastColumn = dataSheet.getColumns();
        for ( int zeroDataColumn = 0;  zeroDataColumn < lastColumn; zeroDataColumn++){
            String metaColumn = POI_Utils.indexToAlpha(zeroDataColumn + 1); //Plus one as metaColumn on over from DetaColumn
            prepareDropDowns(masterSheet, lastMetaRow, metaSheet, metaColumn);
            copyData(lastMetaRow + 2, metaSheet, dataSheet, metaColumn);
        }
    }

    private void prepareSheets(Workbook masterWorkbook, CYAB_Workbook metaWorkbook, Workbook dataWorkbook) 
            throws JavaToExcelException, XLWrapException, XLWrapEOFException{
        Sheet masterSheet = masterWorkbook.getSheet("Sheet1");
        String[] dataSheets = dataWorkbook.getSheetNames();
        for (int i = 0; i  < dataSheets.length; i++){
            Sheet dataSheet = dataWorkbook.getSheet(dataSheets[i]);
            CYAB_Sheet metaSheet = metaWorkbook.getSheet(dataSheets[i]);
            prepareSheet(masterSheet, metaSheet, dataSheet);
        }
    }

    public void prepareMetaData(ExecutionContext context, String dataFile, String doi) throws FileNotFoundException,
            IOException, InvalidFormatException, JavaToExcelException, XLWrapException, XLWrapEOFException{
        System.out.println("Preparing meta data collector for " + dataFile);
        Workbook dataWorkbook = context.getWorkbook("file:" + dataRoot +dataFile);
        Workbook masterWorkbook = context.getWorkbook("file:" + metaMaster);
        CYAB_Workbook metaWorkbook = new CYAB_Workbook();
        addMetaDataSheet(metaWorkbook, dataFile, doi);
        createNamedRanges(masterWorkbook, metaWorkbook);
        //ListWriter.writeLists(masterWorkbook, metaWorkbook);
        prepareSheets(masterWorkbook, metaWorkbook,dataWorkbook);
        writeMeta(metaWorkbook, dataFile);
    }

    public void prepareMetaData(String dataPath, String targetPath) throws FileNotFoundException,
            IOException, InvalidFormatException, JavaToExcelException, XLWrapException, XLWrapEOFException{
        ExecutionContext context =  new ExecutionContext();
        Workbook dataWorkbook;
        try {
            dataWorkbook = context.getWorkbook(dataPath);
        } catch (Exception e){
            //assume the "file:" bit is missing
            dataWorkbook = context.getWorkbook("file:" + dataPath);
        }
        Workbook masterWorkbook = context.getWorkbook("file:" + MASTER_FILE);
        CYAB_Workbook metaWorkbook = new CYAB_Workbook();
        //addMetaDataSheet(metaWorkbook, dataFile, doi);
        createNamedRanges(masterWorkbook, metaWorkbook);
        //ListWriter.writeLists(masterWorkbook, metaWorkbook);
        prepareSheets(masterWorkbook, metaWorkbook,dataWorkbook);
        metaWorkbook.write(targetPath);
    }

    public static void main(String[] args) throws IOException, FileNotFoundException, InvalidFormatException, 
            JavaToExcelException, XLWrapException, XLWrapEOFException{
        MetaDataCreator creator = new MetaDataCreator(MASTER_FILE, MAIN_ROOT + "Meta Data/",MAIN_ROOT + "Raw Data/");
        ExecutionContext context = new ExecutionContext();
        creator.prepareMetaData (context, "Records.xls", "rec12564");
        creator.prepareMetaData (context, "Species.xls", "spec564");
        creator.prepareMetaData (context, "Stokoe.xls", "stokoe32433232");
        creator.prepareMetaData (context, "Tarns.xls", "tarns33exdw2");
    }
}
