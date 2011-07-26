package org.freshwaterlife.fishlink.xlwrap;

import org.freshwaterlife.fishlink.FishLinkException;
import at.jku.xlwrap.spreadsheet.Sheet;

/**
 *
 * @author Christian
 */
public class MasterReader extends AbstractSheet {

    MasterReader (Sheet masterDropdownSheet) throws FishLinkException {
       super(masterDropdownSheet);
    }

    private void checkName (SheetWrite sheet, String otherName, int firstRow, int lastRow) throws FishLinkException {
        for (int i = firstRow; i<= lastRow; i++){
            if (this.getCellValue("A", i).equalsIgnoreCase(otherName)){
                return;
            }
        }
        throw new FishLinkException ("In " + sheet.getSheetInfo() + " Map contains " + otherName + " in the master does not in the same place");
    }

    void check (SheetWrite other) throws FishLinkException {
       if (other.externalSheetRow > 0){
           if (this.externalSheetRow == -1){
                throw new FishLinkException("Map contains \"external sheet\" but master does not.");
           }
       }
       if (other.ZeroNullRow > 0){
           if (this.ZeroNullRow == -1){
                throw new FishLinkException("Map contains \"ignore zero\" but master does not.");
           }
       }
       if (other.firstConstant > 0){
           for (int i = other.firstConstant; i <= other.lastConstant; i++){
               checkName (other, other.getCellValue("A", i), firstConstant, lastConstant);
           }
       }
    }
}
