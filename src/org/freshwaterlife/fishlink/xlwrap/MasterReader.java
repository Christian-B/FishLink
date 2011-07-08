package org.freshwaterlife.fishlink.xlwrap;

import at.jku.xlwrap.spreadsheet.Sheet;

/**
 *
 * @author Christian
 */
public class MasterReader extends AbstractSheet {

    public MasterReader (Sheet masterDropdownSheet) throws XLWrapMapException {
       super(masterDropdownSheet);
    }

    private void checkName (SheetWrite sheet, String otherName, int firstRow, int lastRow) throws XLWrapMapException {
        for (int i = firstRow; i<= lastRow; i++){
            if (this.getCellValue("A", i).equalsIgnoreCase(otherName)){
                return;
            }
        }
        throw new XLWrapMapException ("In " + sheet.getSheetInfo() + " Map contains " + otherName + " in the master does not in the same place");
    }

    void check (SheetWrite other) throws XLWrapMapException {
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
       if (other.firstConstant > 0){
           for (int i = other.firstConstant; i <= other.lastConstant; i++){
               checkName (other, other.getCellValue("A", i), firstConstant, lastConstant);
           }
       }
    }
}
