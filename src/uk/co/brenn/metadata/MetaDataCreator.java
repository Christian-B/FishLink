package uk.co.brenn.metadata;

import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.exec.ExecutionContext;
import at.jku.xlwrap.spreadsheet.Cell;
import at.jku.xlwrap.spreadsheet.Sheet;
import at.jku.xlwrap.spreadsheet.TypeAnnotation;
import at.jku.xlwrap.spreadsheet.Workbook;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Date;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Name;

/**
 *
 * @author Christian
 */
public class MetaDataCreator {

    static private String MAIN_ROOT = "c:/Dropbox/FishLink XLWrap data/";
    static private String MASTER_FILE = "data/MetaMaster.xlsx";
    static private String LIST_SHEET = "Lists";

    private String metaRoot;
    
    static private int CATEGORY_ROW = 1;
    static private int FIELD_ROW = 2;
    private Sheet Sheet;

    private static ExecutionContext context1;

    private static Workbook masterWorkbook;

    public MetaDataCreator(String metaDir){
        metaRoot = metaDir;
    }

    public MetaDataCreator(){
    }
    
    private static Workbook getMasterWorkbook() throws XLWrapException{
        if (masterWorkbook == null){
            masterWorkbook = getExecutionContext().getWorkbook("file:" + MASTER_FILE);
        }
        return masterWorkbook;
    }

    public static ExecutionContext getExecutionContext(){
        if (context1 == null){
            context1 = new ExecutionContext();
        }
        return context1;
    }

    public static Sheet getMasterDropdownSheet() throws XLWrapException{
        return getMasterWorkbook().getSheet("Sheet1");
    }

    public static Sheet getMasterListSheet() throws XLWrapException{
        return getMasterWorkbook().getSheet(LIST_SHEET);
    }

    private void addMetaDataSheet(CYAB_Workbook dataWorkbook, String dataFile, String doi){
        CYAB_Sheet sheet = dataWorkbook.getSheet("MetaData");
        sheet.setValue("A", 1, "File");
        sheet.setValue("B", 1, dataFile);
        sheet.setValue("A", 2, "Doi");
        sheet.setValue("B", 2, doi);
    }

    private void writeMeta(CYAB_Workbook metaWorkbook, String dataName) throws FileNotFoundException, IOException{
        String fileFront = dataName.substring(0, dataName.lastIndexOf("."));
        metaWorkbook.write(metaRoot + fileFront + "MetaData.xls");
    }

    private void createNamedRanges (CYAB_Workbook metaWorkbook)
            throws XLWrapException, XLWrapEOFException {
        Sheet masterSheet =  getMasterListSheet();
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
            range.setRefersToFormula(rangeDef);
            zeroColumn++;
            rangeName = getTextZeroBased(masterSheet, zeroColumn, 0);
            if (rangeName.isEmpty() && categories == -1){
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

    public static String getTextZeroBased(Sheet sheet, int column, int row) throws XLWrapException, XLWrapEOFException{
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
        } while (!value.isEmpty());
        metaSheet.autoSizeColumn(0);
        int maxRow = dataSheet.getRows();
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
            throws XLWrapException, XLWrapEOFException{
        for (int zeroRow = 0; zeroRow < lastMetaRow; zeroRow ++) {
            String list = getTextZeroBased(masterSheet, 2, zeroRow);
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
            metaSheet.addValidation("B", column, zeroRow + 1, list, popupTitle, popupMessage,
                    errorStyle, errorTitle, errorMessage);
        }
    }

    private void copyData(int letterRow, CYAB_Sheet metaSheet, Sheet dataSheet, String metaColumn)
            throws XLWrapException, XLWrapEOFException{
        int zeroColumn = POI_Utils.alphaToIndex(metaColumn) - 1; //-1 as Column A of Data goes in B of Meta
        String dataColumn = POI_Utils.indexToAlpha(zeroColumn);
        metaSheet.setValue(metaColumn, letterRow, dataColumn);
        int maxRow = dataSheet.getRows();
        if (maxRow > 100){
            maxRow = 100;
        }
        for (int zeroRow = 0; zeroRow < maxRow; zeroRow++){
            Cell cell = dataSheet.getCell(zeroColumn, zeroRow);
            TypeAnnotation typeAnnotation = cell.getType();
            switch (typeAnnotation){
                case BOOLEAN:
                    boolean booleanValue = cell.getBoolean();
                    metaSheet.setValue(metaColumn, letterRow + zeroRow + 1, booleanValue);
                    break;
                case NUMBER:
                    double doubleValue = cell.getNumber();
                    metaSheet.setValue(metaColumn, letterRow + zeroRow + 1, doubleValue);
                    break;
                case TEXT:
                    String textValue = cell.getText();
                    metaSheet.setValue(metaColumn, letterRow + zeroRow + 1, textValue);
                    break;
                case DATE:
                    Date dateValue = cell.getDate();
                    String format = cell.getDateFormat();
                    metaSheet.setValue(metaColumn, letterRow + zeroRow + 1, dateValue, format);
                    break;
                case NULL:
                    break;
                default:
                    throw new XLWrapException("Unexpected Cell Type");
            }
        }
    }

    private void prepareSheet(Sheet masterSheet, CYAB_Sheet metaSheet, Sheet dataSheet) 
            throws XLWrapException, XLWrapEOFException{
        int lastMetaRow = prepareColumnA(masterSheet, metaSheet, dataSheet);
        int lastColumn = dataSheet.getColumns();
        for ( int zeroDataColumn = 0;  zeroDataColumn < lastColumn; zeroDataColumn++){
            String metaColumn = POI_Utils.indexToAlpha(zeroDataColumn + 1); //Plus one as metaColumn on over from DetaColumn
            prepareDropDowns(masterSheet, lastMetaRow, metaSheet, metaColumn);
            copyData(lastMetaRow + 2, metaSheet, dataSheet, metaColumn);
        }
    }

    private boolean containsData (Sheet dataSheet){
        if (dataSheet.getColumns() < 1) return false;
        if (dataSheet.getRows() < 1) return false;
        return true;
    }

    private void prepareSheets(CYAB_Workbook metaWorkbook, Workbook dataWorkbook) 
            throws XLWrapException, XLWrapEOFException{
        Sheet masterSheet = getMasterDropdownSheet();
        String[] dataSheets = dataWorkbook.getSheetNames();
        for (int i = 0; i  < dataSheets.length; i++){
            Sheet dataSheet = dataWorkbook.getSheet(dataSheets[i]);
            if (containsData(dataSheet)){
                CYAB_Sheet metaSheet = metaWorkbook.getSheet(dataSheets[i]);
                prepareSheet(masterSheet, metaSheet, dataSheet);
            } else  {
                System.out.println("Skipping empty " + dataSheet.getSheetInfo());
            }
        }
    }

    public void prepareMetaDataOnDoi(String dataRoot, String dataFile, String doi) throws FileNotFoundException,
            IOException, InvalidFormatException, XLWrapException, XLWrapEOFException{
        System.out.println("Preparing meta data collector for " + dataFile);
        Workbook dataWorkbook = getExecutionContext().getWorkbook("file:" + dataRoot +dataFile);
        CYAB_Workbook metaWorkbook = new CYAB_Workbook();
        addMetaDataSheet(metaWorkbook, dataFile, doi);
        createNamedRanges(metaWorkbook);
        prepareSheets(metaWorkbook,dataWorkbook);
        writeMeta(metaWorkbook, dataFile);
    }

    public void prepareMetaDataOnTarget(String dataPath, String targetPath) throws FileNotFoundException,
            IOException, InvalidFormatException, XLWrapException, XLWrapEOFException{
        Workbook dataWorkbook;
        try {
            dataWorkbook = getExecutionContext().getWorkbook(dataPath);
        } catch (Exception e){
            //assume the "file:" bit is missing
            dataWorkbook = getExecutionContext().getWorkbook("file:" + dataPath);
        }
        CYAB_Workbook metaWorkbook = new CYAB_Workbook();
        createNamedRanges(metaWorkbook);
        prepareSheets(metaWorkbook,dataWorkbook);
        metaWorkbook.write(targetPath);
    }

    public static void main(String[] args) throws IOException, FileNotFoundException, InvalidFormatException, 
           XLWrapException, XLWrapEOFException{
        MetaDataCreator creator = new MetaDataCreator(MAIN_ROOT + "Meta Data/");

        //creator.prepareMetaDataOnDoi (MAIN_ROOT + "Raw Data/", "FBA_Tarns.xls", "FBA345");
        //creator.prepareMetaDataOnDoi (MAIN_ROOT + "Raw Data/", "Species.xls", "spec564");
        //creator.prepareMetaDataOnDoi (MAIN_ROOT + "Raw Data/", "Stokoe.xls", "stokoe32433232");
        creator.prepareMetaDataOnDoi (MAIN_ROOT + "Raw Data/", "WillbyGroups.xls", "wbgROUPS8734");
    }
}
