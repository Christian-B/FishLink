package org.freshwaterlife.fishlink.metadatacreator;

import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.exec.ExecutionContext;
import at.jku.xlwrap.spreadsheet.Cell;
import at.jku.xlwrap.spreadsheet.Sheet;
import at.jku.xlwrap.spreadsheet.TypeAnnotation;
import at.jku.xlwrap.spreadsheet.Workbook;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import java.io.File;
import java.util.Date;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Name;
import org.freshwaterlife.fishlink.FishLinkPaths;
import org.freshwaterlife.fishlink.FishLinkUtils;
import org.freshwaterlife.fishlink.FishLinkConstants;
import org.freshwaterlife.fishlink.FishLinkException;
import org.freshwaterlife.fishlink.xlwrap.expr.func.FishLinkToXlWrapRegister;

/**
 *
 * @author Christian
 */
public class MetaDataCreator {
   
    //Sheet masterSheet;
    
    private ExecutionContext context;

    /**
     * Constructs a new MetaDataCreator class.
     * 
     * Creates an XLWrap ExecutionContext
     */
    public MetaDataCreator(){
        context = new ExecutionContext();
    }
        
    private String createNamedRange (Sheet masterListSheet, FishLinkWorkbook metaWorkbook, int zeroColumn) throws FishLinkException {
        FishLinkSheet metaSheet = metaWorkbook.getSheet(FishLinkConstants.LIST_SHEET);
        String rangeName =  FishLinkUtils.getTextZeroBased(masterListSheet, zeroColumn, 0);
        if (rangeName == null || rangeName.isEmpty()){
            return "";
        }
        String columnName = FishLinkUtils.indexToAlpha(zeroColumn);
        metaSheet.setValueZeroBased(zeroColumn, 0, rangeName);
        int zeroRow = 1;
        String fieldName = FishLinkUtils.getTextZeroBased(masterListSheet, zeroColumn, zeroRow);
        while (!fieldName.isEmpty()) {
            metaSheet.setValueZeroBased(zeroColumn, zeroRow, fieldName);
            zeroRow++;
            fieldName = FishLinkUtils.getTextZeroBased(masterListSheet, zeroColumn, zeroRow);
        } while (!fieldName.isEmpty());
        Name range = metaWorkbook.createName();
        range.setNameName(rangeName);
        String rangeDef = "'" + FishLinkConstants.LIST_SHEET + "'!$" + columnName + "$2:$" + columnName + "$" + zeroRow;
        range.setRefersToFormula(rangeDef);
        return rangeName;
    }

    private void createNamedRanges (Sheet masterListSheet, FishLinkWorkbook metaWorkbook) throws FishLinkException {
        FishLinkSheet metaSheet = metaWorkbook.getSheet(FishLinkConstants.LIST_SHEET);
        //Get the categories
        int zeroColumn = 0;
        String rangeName =  createNamedRange(masterListSheet, metaWorkbook, zeroColumn);
        while (!rangeName.isEmpty()) {
            zeroColumn++;
            rangeName = createNamedRange(masterListSheet, metaWorkbook, zeroColumn);
        }
        //first space splits the categories from the rest
        Name range = metaWorkbook.createName();
        range.setNameName("Category");
        String columnName = FishLinkUtils.indexToAlpha(zeroColumn -1);
        range.setRefersToFormula("'" + FishLinkConstants.LIST_SHEET + "'!$A1:$" + columnName + "$1");
        //Now get the subTypes
        do {
            zeroColumn++;
            rangeName = createNamedRange(masterListSheet, metaWorkbook, zeroColumn);
        } while (!rangeName.isEmpty());
        //Now the renaming ranges
        do {
            zeroColumn++;
            rangeName = createNamedRange(masterListSheet, metaWorkbook, zeroColumn);
        } while (!rangeName.isEmpty());
        //Hack to avoid last column getting lost
        metaSheet.setValueZeroBased(zeroColumn + 1, 0, "");
    }

    private int prepareColumnA(Sheet masterSheet, FishLinkSheet metaSheet) throws FishLinkException {
        int zeroRow = 0;
        String value;
        do {
            value = FishLinkUtils.getTextZeroBased(masterSheet, 0, zeroRow);
            metaSheet.setValue("A",zeroRow + 1, value);
            zeroRow++;
        } while (!value.isEmpty());
        metaSheet.autoSizeColumn(0);
        zeroRow++; //leave a blank row
        metaSheet.setValue("A",zeroRow, "Header");
        return zeroRow;
    }

    private void prepareDropDowns(Sheet masterSheet, int lastMetaRow, FishLinkSheet metaSheet, String column) 
            throws FishLinkException {
        for (int zeroRow = 0; zeroRow < lastMetaRow; zeroRow ++) {
            String list = FishLinkUtils.getTextZeroBased(masterSheet, 2, zeroRow);
            String popupTitle = FishLinkUtils.getTextZeroBased(masterSheet, 3, zeroRow);
            String popupMessage =  FishLinkUtils.getTextZeroBased(masterSheet, 4, zeroRow);
            String errorStyleString =  FishLinkUtils.getTextZeroBased(masterSheet, 5, zeroRow);
            int errorStyle = DataValidation.ErrorStyle.STOP;
            if (errorStyleString.startsWith("W")){
                errorStyle = DataValidation.ErrorStyle.WARNING;
            }
            if (errorStyleString.startsWith("I")){
                errorStyle = DataValidation.ErrorStyle.INFO;
            }
            String errorTitle =  FishLinkUtils.getTextZeroBased(masterSheet, 6, zeroRow);
            String errorMessage =  FishLinkUtils.getTextZeroBased(masterSheet, 7, zeroRow);
            metaSheet.addValidation("B", column, zeroRow + 1, list, popupTitle, popupMessage,
                    errorStyle, errorTitle, errorMessage);
        }
    }

    private void copyMetaData(Sheet masterSheet, FishLinkSheet metaSheet, Sheet copySheet, int lastMetaRow, int lastColumn) 
            throws FishLinkException{
        //ystem.out.println(lastMetaRow + " " + copySheet.getSheetInfo());
        for (int zeroRow = 0; zeroRow < lastMetaRow; zeroRow ++) {
            String masterColumnA = FishLinkUtils.getTextZeroBased(masterSheet, 0, zeroRow);
            int copyRow = -1;
            for (int tryRow = 0; tryRow <= lastMetaRow; tryRow ++) {
                 String tryColumnA = FishLinkUtils.getTextZeroBased(copySheet, 0, tryRow);
                 if (tryColumnA.equalsIgnoreCase(masterColumnA)){
                     copyRow = tryRow;
                 }
            }
            //ystem.out.println(masterColumnA + ": " + copyRow);
            if (copyRow >= 0){
                for ( int zeroColumn = 0;  zeroColumn <= lastColumn; zeroColumn++){
                    String copyMeta = FishLinkUtils.getTextZeroBased(copySheet, zeroColumn, copyRow);
                    metaSheet.setValueZeroBased(zeroColumn, zeroRow, copyMeta);
                }
            }
        }
    }
            
    private void copyData(int headerRow, FishLinkSheet metaSheet, Sheet dataSheet, String metaColumn, 
            int zeroDataColumn, int ignoreRows) throws FishLinkException {
       // String dataColumn = FishLinkUtils.indexToAlpha(zeroDataColumn);
       // metaSheet.setValue(metaColumn, letterRow, dataColumn);
        for (int zeroRow = 0; zeroRow < dataSheet.getRows() - ignoreRows; zeroRow++){
            Cell cell;
            try {
                cell = dataSheet.getCell(zeroDataColumn, zeroRow + ignoreRows);
            } catch (XLWrapException ex) {
                throw new FishLinkException("Error getting Cell", ex);
            } catch (XLWrapEOFException ex) {
                throw new FishLinkException("Error getting Cell", ex);
            }
            TypeAnnotation typeAnnotation;
            try {
                typeAnnotation = cell.getType();
            } catch (XLWrapException ex) {
                throw new FishLinkException("Error getting annotation type", ex);
            }
            switch (typeAnnotation){
                case BOOLEAN:
                    boolean booleanValue;
                    try {
                        booleanValue = cell.getBoolean();
                    } catch (XLWrapException ex) {
                        throw new FishLinkException("Error getting boolean value", ex);
                    }
                    metaSheet.setValue(metaColumn, headerRow + zeroRow, booleanValue);
                    break;
                case NUMBER:
                    double doubleValue;
                    try {
                        doubleValue = cell.getNumber();
                    } catch (XLWrapException ex) {
                        throw new FishLinkException("Error getting double value", ex);
                    }
                    metaSheet.setValue(metaColumn, headerRow + zeroRow, doubleValue);
                    break;
                case TEXT:
                    String textValue;
                    try {
                        textValue = cell.getText();
                    } catch (XLWrapException ex) {
                        throw new FishLinkException("Error getting text value", ex);
                    }
                    metaSheet.setValue(metaColumn, headerRow + zeroRow, textValue);
                    break;
                case DATE:
                    Date dateValue;
                    try {
                        dateValue = cell.getDate();
                    } catch (XLWrapException ex) {
                        throw new FishLinkException("Error getting date text value", ex);
                    }
                    String format;
                    try {
                        format = cell.getDateFormat();
                    } catch (XLWrapException ex) {
                        throw new FishLinkException("Error getting date format", ex);
                    }
                    metaSheet.setValue(metaColumn, headerRow + zeroRow, dateValue, format);
                    break;
                case NULL:
                    break;
                default:
                    throw new FishLinkException("Unexpected Cell Type");
            }
        }
        metaSheet.setForegroundAqua(metaColumn, headerRow);
        metaSheet.autoSizeColumn(zeroDataColumn);
    }
    
    private int isOldMetaData(Sheet dataSheet) throws FishLinkException{
        String columnA = FishLinkUtils.getTextZeroBased(dataSheet, 0, 0);
        if (!columnA.equalsIgnoreCase(FishLinkConstants.CATEGORY_LABEL)){
            return 0;
        }
        columnA = FishLinkUtils.getTextZeroBased(dataSheet, 0, 1);
        if (!columnA.equalsIgnoreCase(FishLinkConstants.FIELD_LABEL)){
            return 0;
        }
        //Ok assume it is old metaData
        int zeroRow = 2;
        String value;
        do {
            value = FishLinkUtils.getTextZeroBased(dataSheet, 0, zeroRow);
            zeroRow++;
        } while (!value.isEmpty());
        return zeroRow;  
    }

    private void prepareSheet(Sheet masterSheet, FishLinkSheet metaSheet, Sheet dataSheet) throws FishLinkException {
        int ignoreRows = isOldMetaData(dataSheet);
        int columnShift;
        if (ignoreRows == 0){
            //Plus one as metaColumn on over from DetaColumn
            columnShift = 1;
        } else {
            //dataSheet is a metaSheet so already has a none data column A
            columnShift = 0;
        }
        int headerRow = prepareColumnA(masterSheet, metaSheet);
        int lastColumn = dataSheet.getColumns();
        metaSheet.createFreezePane("B", headerRow);
        for ( int zeroDataColumn = 0;  zeroDataColumn < lastColumn; zeroDataColumn++){
            String metaColumn = FishLinkUtils.indexToAlpha(zeroDataColumn + columnShift); 
            prepareDropDowns(masterSheet, headerRow -2, metaSheet, metaColumn);
            copyData(headerRow, metaSheet, dataSheet, metaColumn, zeroDataColumn, ignoreRows);
        }
        //Hack tp avoid last column getting lost
        metaSheet.setValueZeroBased(lastColumn + 1, headerRow, "");
        if (columnShift == 0){
            copyMetaData(masterSheet, metaSheet, dataSheet, headerRow -2, lastColumn); 
        }
    }

    private boolean containsData (Sheet dataSheet){
        if (dataSheet.getColumns() < 1) return false;
        if (dataSheet.getRows() < 1) return false;
        return true;
    }

    private void prepareSheets(Sheet masterDropdownSheet, FishLinkWorkbook metaWorkbook, Workbook dataWorkbook) 
            throws FishLinkException{
        String[] dataSheets = dataWorkbook.getSheetNames();
        for (int i = 0; i  < dataSheets.length; i++){
            if (!dataSheets[i].equalsIgnoreCase(FishLinkConstants.LIST_SHEET)){
                Sheet  dataSheet;
                try {
                    dataSheet = dataWorkbook.getSheet(dataSheets[i]);
                } catch (XLWrapException ex) {
                    throw new FishLinkException("Unable to get sheet " + dataSheets[i], ex);
                }
                if (containsData(dataSheet)){
                    FishLinkSheet metaSheet = metaWorkbook.getSheet(dataSheets[i]);
                    prepareSheet(masterDropdownSheet, metaSheet, dataSheet);
                } else  {
                    FishLinkUtils.report("Skipping empty " + dataSheet.getSheetInfo());
                }
            }
        }
    }

    private void writeMetaData(String dataUrl, String masterUrl, File output) throws FishLinkException{
        FishLinkUtils.report("Creating metaData sheet for " + dataUrl);
        Sheet masterListSheet;
        try {
            masterListSheet = context.getSheet(masterUrl, FishLinkConstants.LIST_SHEET);
        } catch (XLWrapException ex) {
            throw new FishLinkException("Error opening the vocabulary sheet " + FishLinkConstants.LIST_SHEET + 
                    " in ExcelSheet " + masterUrl, ex);
        }
        Sheet masterDropdownSheet;
        try {
            masterDropdownSheet = context.getSheet(masterUrl, FishLinkConstants.DROP_DOWN_SHEET);
        } catch (XLWrapException ex) {
            throw new FishLinkException("Error opening the dropdown sheet " + FishLinkConstants.DROP_DOWN_SHEET+ 
                    " in ExcelSheet " + masterUrl, ex);
        }
        Workbook dataWorkbook;
        try {
            dataWorkbook = context.getWorkbook(dataUrl);
        } catch (XLWrapException ex) {
            throw new FishLinkException("Error opening the dataset " + dataUrl, ex);
        }
        FishLinkWorkbook metaWorkbook = new FishLinkWorkbook();
        createNamedRanges(masterListSheet, metaWorkbook);
        //TODO copy
        prepareSheets(masterDropdownSheet, metaWorkbook, dataWorkbook);
        metaWorkbook.write(output);
        FishLinkUtils.report("Wrote to  " + output.getAbsolutePath());
    }

    /**
     * Calls {@link #createMetaData(java.lang.String, java.lang.String)} with the default MetaMaster
     * @param dataUrl
     * @return
     * @throws FishLinkException 
     */
    public File createMetaData(String dataUrl) throws FishLinkException{
        return createMetaData(dataUrl, FishLinkPaths.MASTER_FILE);
    }
    
    /**
     * Creates a annotated workbook based on the dataFile and the masterFile.
     * 
     * Opens an XLWrap Workbook based on dataUrl, and an XLWrap Workbook based on masterUrl.
     * Creates a new Excel workbook which contains the vocabulary lists from the master and all the sheets from data,
     *     annotated with the dropDowns found in the Master.
     * <p>Both the data pointed to by dataUrl and the masterFile pointed to by masterUrl 
     *     must be in a format XLWrap can handle.
     * <p>File formats XLWrap can handle include
     * <ul>
     *     <li>Excel xls files
     *     <li>Excel 2001 xlsx files (Assuming the FishLink extended XlWrap is used.
     *     <li>csv Files
     *     <li>Open Office files. (untested)
     * </ul>
     * <p>Know protocols it can handle included
     * <ul>
     *     <li>File: 
     *     <li>http:
     * </ul>
     * <p>If the data already contains annotated header rows, this tool will copy over the notations are much as possible.
     *     Even if they are no longer correct. The vocabulary lists and dropdown rules are taken from the new MetaMaster.
     * @param dataUrl Url to the raw data or a previous annotated data workbook to be updated. 
     * @param masterUrl URl to the MetaMaster file
     * @return File Where the annotated sheet is saved.
     * @throws FishLinkException Any Exceptions that might have occurred. Including that where thrown by XLWrap. 
     */
    public File createMetaData(String dataUrl, String masterUrl) throws FishLinkException{
        String[] parts = dataUrl.split("[\\\\/.]");// Split on a forawrd slash, back slash and a full stop
        System.out.println("=======");
        String fileName;
        if (parts[parts.length-2].endsWith("MetaData")){
            fileName = parts[parts.length-2] + ".xls";
        } else {
            fileName = parts[parts.length-2] + "MetaData.xls";
        }
        File output =  new File (FishLinkPaths.META_FILE_ROOT, fileName);
        writeMetaData(dataUrl, masterUrl, output);
        return output;
    }
   
    private static void usage(){
        FishLinkUtils.report("Creates a Meta Data collection file based on input.");
        FishLinkUtils.report("Requires two paramters.");
        FishLinkUtils.report("First is the raw data to be described.");
        FishLinkUtils.report("Second if the location (path and file name) of the MetaData file.");
    }

    /**
     * Main method for creating MetaData.
     * 
     * Converts the first argument into DataUrl, the second into MasterUrl and 
     * then calls {@link #createMetaData(java.lang.String, java.lang.String)}
     * @param args DataUrl and MasterUrl
     * @throws FishLinkException 
     */
    public static void main(String[] args) throws FishLinkException{
        if (args.length != 2){
            usage();
            System.exit(2);
        }
        FishLinkToXlWrapRegister.register();
        MetaDataCreator metaDataCreator = new MetaDataCreator();
        File output = metaDataCreator.createMetaData(args[0], args[1]);
        FishLinkUtils.report("Mapping file can be found at: "+output.getAbsolutePath());
    }

}
