/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package uk.co.brenn.xlwrap;

import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import uk.co.brenn.metadata.MetaDataCreator;

/**
 *
 * @author Christian
 */
public class MasterReader extends AbstractSheet {

    public MasterReader ()
            throws XLWrapException, XLWrapEOFException {
       super(MetaDataCreator.getMasterDropdownSheet());
    }

    private void checkName (SheetWrite sheet, String otherName, int firstRow, int lastRow)
            throws XLWrapException, XLWrapEOFException, XLWrapMapException {
        for (int i = firstRow; i<= lastRow; i++){
            if (this.getCellValue("A", i).equalsIgnoreCase(otherName)){
                return;
            }
        }
        throw new XLWrapMapException ("In " + sheet.getSheetInfo() + " Map contains " + otherName + " in the master does not in the same place");
    }

    public void check (SheetWrite other) throws XLWrapMapException, XLWrapException, XLWrapEOFException{
        System.out.println("checking " + other.toString());
       //categoryRow checked by findAndCheckMetaSplits
       //fieldRow checked by findAndCheckMetaSplits
       //idTypeRow checked by findAndCheckMetaSplits
       if (other.externalSheetRow > 0){
           if (this.externalSheetRow == -1){
                throw new XLWrapMapException("Map contains \"external sheet\" but master does not.");
           }
       }
       if (other.ignoreZerosRow > 0){
           if (this.ignoreZerosRow == -1){
                throw new XLWrapMapException("Map contains \"ignore zero\" but master does not.");
           }
       }
       if (other.firstLink > 0){
           for (int i = other.firstLink; i <= other.lastLink; i++){
               checkName (other, other.getCellValue("A", i), firstLink, lastLink);
           }
       }
       if (other.firstConstant > 0){
           for (int i = other.firstConstant; i <= other.lastConstant; i++){
               checkName (other, other.getCellValue("A", i), firstConstant, lastConstant);
           }
       }
       System.out.println("checked" + other.toString());
    }
}
