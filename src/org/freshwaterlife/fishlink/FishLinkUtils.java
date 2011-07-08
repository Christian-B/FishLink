package org.freshwaterlife.fishlink;

import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.spreadsheet.Cell;
import at.jku.xlwrap.spreadsheet.Sheet;
import at.jku.xlwrap.spreadsheet.Workbook;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import org.freshwaterlife.fishlink.xlwrap.XLWrapMapException;

/**
 *
 * @author Christian
 */
public class FishLinkUtils {

    /**
     * @param alpha index
     * @return zero-based numerical index
     */
    public static final int alphaToIndex(String alpha) {
	char[] letters = alpha.toUpperCase().toCharArray();
	int index = 0;
	for (int i = 0; i < letters.length; i++)
            index += ((letters[letters.length-i-1]) - 64) * (Math.pow(26, i)); // A is 64
	return --index;
    }

    /**
     *
     * @param index zero-based numerical index
     * @return alpha index
     */
    public static final String indexToAlpha(int index){
        String reply = "";
        if (index < 26){
           char first = (char)( index + 65);
           reply = first + "";
        } else {
           char last = (char)(index - ((index / 26) * 26) + 65);
           String rest = indexToAlpha((index / 26)-1);
           reply = rest + last;
        }
        if (index != alphaToIndex(reply)){
            System.err.println("Error converting " + index);
            System.err.println("Answer was "+ reply);
            System.err.println("Which comes out as "+ alphaToIndex(reply));
            throw new AssertionError("Error in indexToAlpha");
        }
        return reply;
    }

   /* public static final Workbook getWorkbookOnPid (String pid) throws XLWrapMapException{
        PidRegister pidRegister = PidStore.padStoreFactory();
        String dataFile = pidRegister.retreiveFile(pid);
        return getWorkbook(dataFile);
    }*/

    /*public static Workbook getWorkbook (String path) throws XLWrapMapException{
        try {
            return MasterFactory.getExecutionContext().getWorkbook(path);
        } catch (XLWrapException ex) {
            //try {
            //    return MasterFactory.getExecutionContext().getWorkbook("file:" + path);
            //} catch (XLWrapException ex2) {
                throw new XLWrapMapException ("Unable to create workbook", ex);
            //}        
        }        
    }
*/
    public static Sheet getSheet(Workbook workbook, String name) throws XLWrapMapException{
        try {
            return workbook.getSheet(name);
        } catch (XLWrapException ex) {
            throw new XLWrapMapException ("Unable to create metaData in workbook " + workbook.getWorkbookInfo(), ex);
        }     
    }
    
    public static Cell getCell(Sheet metaData, int col, int row) throws XLWrapMapException{
        try{
            return metaData.getCell(col, row);
        } catch (NullPointerException ex) {
            throw new XLWrapMapException ("Unable to find cell " + col + "  " + row + " in sheet " + 
                    metaData.getSheetInfo(), ex);
        } catch (XLWrapException ex) {
            throw new XLWrapMapException ("Unable to find cell " + col + "  " + row + " in sheet " + 
                    metaData.getSheetInfo(), ex);
        } catch (XLWrapEOFException ex) {
            throw new XLWrapMapException ("Unable to find cell " + col + "  " + row + " in sheet " + 
                    metaData.getSheetInfo(), ex);
        }        
    }
    
    public static String getText(Cell cell) throws XLWrapMapException{
        try{
            return cell.getText();
        } catch (XLWrapException ex) {
            throw new XLWrapMapException ("Unable to get text from cell " + cell.getCellInfo(), ex);
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

    /**
     * Reports any message in a standard way.
     * 
     * Current implementation is just to System.out but logging can be added here!
     * @param messages 
     */
    public static void report (String message){
        System.out.println(message);
    }

    public static void main(String[] args) {
        String test = indexToAlpha(48);
    }
}