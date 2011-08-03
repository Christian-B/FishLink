package org.freshwaterlife.fishlink.xlwrap;

import org.freshwaterlife.fishlink.FishLinkException;
import at.jku.xlwrap.spreadsheet.Sheet;

/**
 * Wrapper around MetaMaster drop down sheet to check Annotation Sheet does not contain Annotation rows not the MetaMaster.
 * @author Christian
 */
public class MasterReader extends AbstractSheet {

    /**
     * Wraps the passed in Sheet.
     *
     * @param masterDropdownSheet Dropdown sheet from the MetaMaster
     * @throws FishLinkException If the sheet does not contain the required rows
     */
    MasterReader (Sheet masterDropdownSheet) throws FishLinkException {
       super(masterDropdownSheet);
    }

    /**
     * Check to see if the constant annotation rows in the sheet are also in the MetaMaster
     *
     * There is no requirement that the sheet must contain all constant annotation rows in the Master.
     * @param constant Constant to check
     * @throws FishLinkException If the Constant ignore case bot found in the MetaMaster
     */
    private void checkConstant (String constant) throws FishLinkException {
        for (int i = this.firstConstant; i<= this.lastConstant; i++){
            if (this.getCellValue("A", i).equalsIgnoreCase(constant)){
                return;
            }
        }
        throw new FishLinkException ("In " + sheet.getSheetInfo() + " Map contains " + constant +
                " in the master does not in the same place");
    }

    /**
     * Checks all the information rows in the Sheet vs the MetaMaster.
     *
     * There is no requirement that the sheet must all the annotation rows in the Master.
     * <p>
     * Does not check the required rows as the SheetWrite COnstructor does that.

     * @param other Annotated Sheet to Check
     * @throws FishLinkException If the Annotation Sheet contains an Annotation Row not in the MetaMaster
     */
    void check (SheetWrite other) throws FishLinkException {
       if (other.externalSheetRow != NOT_FOUND){
           if (this.externalSheetRow == NOT_FOUND){
                throw new FishLinkException("Map contains \"external sheet\" but master does not.");
           }
       }
       if (other.ZeroNullRow != NOT_FOUND){
           if (this.ZeroNullRow == NOT_FOUND){
                throw new FishLinkException("Map contains \"ignore zero\" but master does not.");
           }
       }
       if (other.firstConstant != NOT_FOUND){
           for (int i = other.firstConstant; i <= other.lastConstant; i++){
               checkConstant (other.getCellValue("A", i));
           }
       }
    }
}
