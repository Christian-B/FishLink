/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package uk.co.brenn.test;

import org.apache.poi.ss.usermodel.Name;
import uk.co.brenn.metadata.CYAB_Sheet;
import uk.co.brenn.metadata.CYAB_Workbook;

/**
 *
 * @author Christian
 */
public class ListWriter {

    public static String CATEGORY_NAME = "Category";
    public static String ASSAY_NAME = "Assay";
    public static String DESIGNATION_NAME = "Designation";

    private static void createRange(CYAB_Workbook workbook, String name, String[] values){
        CYAB_Sheet listSheet = workbook.getSheet("Lists");
        Name rangeName = workbook.createName();
        rangeName.setNameName(name);
        String column = listSheet.findFreeColumn();
        listSheet.setValue(column, 1, name);
        for (int i = 0; i < values.length; i++){
            listSheet.setValue(column, i + 2, values[i]);
        }
        //ystem.out.println (column);
        rangeName.setRefersToFormula("'Lists'!$" + column + "$2:$" + column + "$" + (values.length + 2));
    }

    public static void writeLists (CYAB_Workbook workbook){
        String[] categoryValues = {ASSAY_NAME, DESIGNATION_NAME,"Location","Observation","Person","Site","Species","Survey"};
        createRange(workbook, CATEGORY_NAME, categoryValues);
        String[] assayValues = {"Id", "Name", "AlternativeName", "Abbreviation", "Remark", "Contributor", "Date"};
        createRange(workbook, ASSAY_NAME, assayValues);
        String[] designationValues = {"Id", "Name", "AlternativeName", "Abbreviation", "Remark"};
        createRange(workbook, "Designation", designationValues);
    }
}
