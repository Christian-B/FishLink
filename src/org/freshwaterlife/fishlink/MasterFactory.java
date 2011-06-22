package org.freshwaterlife.fishlink;

import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.exec.ExecutionContext;
import at.jku.xlwrap.spreadsheet.Cell;
import at.jku.xlwrap.spreadsheet.Sheet;
import at.jku.xlwrap.spreadsheet.Workbook;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;

/**
 *
 * @author Christian
 */
public class MasterFactory {

    static private String MASTER_FILE = "data/MetaMaster.xlsx";
    public static final String LIST_SHEET = "Lists";

    private static ExecutionContext context1;

    private static Workbook masterWorkbook;

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

    public static String getTextZeroBased(Sheet sheet, int column, int row)
            throws XLWrapException, XLWrapEOFException{
        if (column >= sheet.getColumns()){
            return "";
        }
        if (row >= sheet.getRows()){
            return "";
        }
        Cell cell = sheet.getCell(column, row);
        return cell.getText();
    }


}
