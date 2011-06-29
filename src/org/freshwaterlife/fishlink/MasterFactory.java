package org.freshwaterlife.fishlink;

import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.exec.ExecutionContext;
import at.jku.xlwrap.spreadsheet.Cell;
import at.jku.xlwrap.spreadsheet.Sheet;
import at.jku.xlwrap.spreadsheet.Workbook;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import java.io.IOException;
import org.freshwaterlife.fishlink.xlwrap.XLWrapMapException;

/**
 *
 * @author Christian
 */
public class MasterFactory {

    public static final String LIST_SHEET = "Lists";

    static final String masterPid = "ff19868d-5b12-846f-81ef-e47468c85068";

    private static ExecutionContext context1;

    private static Workbook masterWorkbook;

    private static Workbook getMasterWorkbook() throws XLWrapMapException{
        if (masterWorkbook == null){
            PidRegister pidRegister;
            try {
                pidRegister = PidStore.padStoreFactory();
                String path = pidRegister.retreiveFileOrNull(masterPid);
                if (path == null){
                    path = "file:" + FishLinkPaths.MASTER_FILE;
                }
                masterWorkbook = getExecutionContext().getWorkbook(path);
            } catch (XLWrapException ex) {
                throw new XLWrapMapException("Unable to find the MetaMaster file", ex);
            }
        }
        return masterWorkbook;
    }

    public static ExecutionContext getExecutionContext(){
        if (context1 == null){
            context1 = new ExecutionContext();
        }
        return context1;
    }

    public static Sheet getMasterDropdownSheet() throws XLWrapMapException{
        try {
            return getMasterWorkbook().getSheet("Sheet1");
        } catch (XLWrapException ex) {
            throw new XLWrapMapException("Unable to find dropdown Sheet in MetaMaster file", ex);
        }
    }

    public static Sheet getMasterListSheet() throws XLWrapMapException{
        try {
            return getMasterWorkbook().getSheet(LIST_SHEET);
        } catch (XLWrapException ex) {
            throw new XLWrapMapException("Unable to find list Sheet in MetaMaster file", ex);
        }
    }

    public static String getTextZeroBased(Sheet sheet, int column, int row) throws XLWrapMapException {
        if (column >= sheet.getColumns()){
            return "";
        }
        if (row >= sheet.getRows()){
            return "";
        }
        try {
            Cell cell = sheet.getCell(column, row);
            return cell.getText();
        } catch (XLWrapException ex) {
            throw new XLWrapMapException("Unable to cell column: " + column + " row:" + row + " from " + 
                    sheet.getSheetInfo(), ex);
        } catch (XLWrapEOFException ex) {
            throw new XLWrapMapException("Unable to cell column: " + column + " row:" + row + " from " + 
                    sheet.getSheetInfo(), ex);
        }
    }


}
