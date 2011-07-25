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
import org.freshwaterlife.fishlink.xlwrap.Constants;
import org.freshwaterlife.fishlink.xlwrap.XLWrapMapException;

/**
 *
 * @author Christian
 */
public class MetaDataCreator {
   
    //Sheet masterSheet;
    
    private ExecutionContext context;

    public MetaDataCreator(){
        context = new ExecutionContext();
    }
        
    private String createNamedRange (Sheet masterListSheet, FishLinkWorkbook metaWorkbook, int zeroColumn) throws XLWrapMapException {
        FishLinkSheet metaSheet = metaWorkbook.getSheet(Constants.LIST_SHEET);
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
        String rangeDef = "'" + Constants.LIST_SHEET + "'!$" + columnName + "$2:$" + columnName + "$" + zeroRow;
        range.setRefersToFormula(rangeDef);
        return rangeName;
    }

    private void createNamedRanges (Sheet masterListSheet, FishLinkWorkbook metaWorkbook) throws XLWrapMapException {
        FishLinkSheet metaSheet = metaWorkbook.getSheet(Constants.LIST_SHEET);
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
        range.setRefersToFormula("'" + Constants.LIST_SHEET + "'!$A1:$" + columnName + "$1");
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

    private int prepareColumnA(Sheet masterSheet, FishLinkSheet metaSheet) throws XLWrapMapException {
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
            throws XLWrapMapException {
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
            throws XLWrapMapException{
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
            int zeroDataColumn, int ignoreRows) throws XLWrapMapException {
       // String dataColumn = FishLinkUtils.indexToAlpha(zeroDataColumn);
       // metaSheet.setValue(metaColumn, letterRow, dataColumn);
        for (int zeroRow = 0; zeroRow < dataSheet.getRows() - ignoreRows; zeroRow++){
            Cell cell;
            try {
                cell = dataSheet.getCell(zeroDataColumn, zeroRow + ignoreRows);
            } catch (XLWrapException ex) {
                throw new XLWrapMapException("Error getting Cell", ex);
            } catch (XLWrapEOFException ex) {
                throw new XLWrapMapException("Error getting Cell", ex);
            }
            TypeAnnotation typeAnnotation;
            try {
                typeAnnotation = cell.getType();
            } catch (XLWrapException ex) {
                throw new XLWrapMapException("Error getting annotation type", ex);
            }
            switch (typeAnnotation){
                case BOOLEAN:
                    boolean booleanValue;
                    try {
                        booleanValue = cell.getBoolean();
                    } catch (XLWrapException ex) {
                        throw new XLWrapMapException("Error getting boolean value", ex);
                    }
                    metaSheet.setValue(metaColumn, headerRow + zeroRow, booleanValue);
                    break;
                case NUMBER:
                    double doubleValue;
                    try {
                        doubleValue = cell.getNumber();
                    } catch (XLWrapException ex) {
                        throw new XLWrapMapException("Error getting double value", ex);
                    }
                    metaSheet.setValue(metaColumn, headerRow + zeroRow, doubleValue);
                    break;
                case TEXT:
                    String textValue;
                    try {
                        textValue = cell.getText();
                    } catch (XLWrapException ex) {
                        throw new XLWrapMapException("Error getting text value", ex);
                    }
                    metaSheet.setValue(metaColumn, headerRow + zeroRow, textValue);
                    break;
                case DATE:
                    Date dateValue;
                    try {
                        dateValue = cell.getDate();
                    } catch (XLWrapException ex) {
                        throw new XLWrapMapException("Error getting date text value", ex);
                    }
                    String format;
                    try {
                        format = cell.getDateFormat();
                    } catch (XLWrapException ex) {
                        throw new XLWrapMapException("Error getting date format", ex);
                    }
                    metaSheet.setValue(metaColumn, headerRow + zeroRow, dateValue, format);
                    break;
                case NULL:
                    break;
                default:
                    throw new XLWrapMapException("Unexpected Cell Type");
            }
        }
        metaSheet.setForegroundAqua(metaColumn, headerRow);
        metaSheet.autoSizeColumn(zeroDataColumn);
    }
    
    private int isOldMetaData(Sheet dataSheet) throws XLWrapMapException{
        String columnA = FishLinkUtils.getTextZeroBased(dataSheet, 0, 0);
        if (!columnA.equalsIgnoreCase(Constants.CATEGORY_LABEL)){
            return 0;
        }
        columnA = FishLinkUtils.getTextZeroBased(dataSheet, 0, 1);
        if (!columnA.equalsIgnoreCase(Constants.FIELD_LABEL)){
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

    private void prepareSheet(Sheet masterSheet, FishLinkSheet metaSheet, Sheet dataSheet) throws XLWrapMapException {
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
            throws XLWrapMapException{
        String[] dataSheets = dataWorkbook.getSheetNames();
        for (int i = 0; i  < dataSheets.length; i++){
            if (!dataSheets[i].equalsIgnoreCase(Constants.LIST_SHEET)){
                Sheet  dataSheet;
                try {
                    dataSheet = dataWorkbook.getSheet(dataSheets[i]);
                } catch (XLWrapException ex) {
                    throw new XLWrapMapException("Unable to get sheet " + dataSheets[i], ex);
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

    public void writeMetaData(String dataUrl, String masterUrl, File output) throws XLWrapMapException{
        FishLinkUtils.report("Creating metaData sheet for " + dataUrl);
        Sheet masterListSheet;
        try {
            masterListSheet = context.getSheet(masterUrl, Constants.LIST_SHEET);
        } catch (XLWrapException ex) {
            throw new XLWrapMapException("Error opening the vocabulary sheet " + Constants.LIST_SHEET + 
                    " in ExcelSheet " + masterUrl, ex);
        }
        Sheet masterDropdownSheet;
        try {
            masterDropdownSheet = context.getSheet(masterUrl, Constants.DROP_DOWN_SHEET);
        } catch (XLWrapException ex) {
            throw new XLWrapMapException("Error opening the dropdown sheet " + Constants.DROP_DOWN_SHEET+ 
                    " in ExcelSheet " + masterUrl, ex);
        }
        Workbook dataWorkbook;
        try {
            dataWorkbook = context.getWorkbook(dataUrl);
        } catch (XLWrapException ex) {
            throw new XLWrapMapException("Error opening the dataset " + dataUrl, ex);
        }
        FishLinkWorkbook metaWorkbook = new FishLinkWorkbook();
        createNamedRanges(masterListSheet, metaWorkbook);
        //TODO copy
        prepareSheets(masterDropdownSheet, metaWorkbook, dataWorkbook);
        metaWorkbook.write(output);
        FishLinkUtils.report("Wrote to  " + output.getAbsolutePath());
    }

    public File createMetaData(String dataUrl) throws XLWrapMapException{
        return createMetaData(dataUrl, FishLinkPaths.MASTER_FILE);
    }
    
    public File createMetaData(String dataUrl, String masterUrl) throws XLWrapMapException{
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
   
    public static void main(String[] args) throws XLWrapMapException{
        MetaDataCreator creator = new MetaDataCreator();
        /*
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Raw Data\\CumbriaTarnsPart1.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Raw Data\\FBA_Tarns.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Raw Data\\Records.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Raw Data\\Species.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Raw Data\\Stokoe.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Raw Data\\Tarns.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Raw Data\\TarnschemFinal.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Raw Data\\WillbyGroups.xls");
         */
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Old Meta Data\\CumbriaTarnsPart1MetaData.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Old Meta Data\\FBA_TarnsMetaData.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Old Meta Data\\RecordsMetaData.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Old Meta Data\\SpeciesMetaData.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Old Meta Data\\StokoeMetaData.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Old Meta Data\\TarnsMetaData.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Old Meta Data\\TarnschemFinalMetaData.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Old Meta Data\\WillbyGroupsMetaData.xls");
    }}
