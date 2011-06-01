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
public class NamedRange {

    public NamedRange(CYAB_Workbook workbook, String name, String[] values){
        CYAB_Sheet listSheet = workbook.getSheet("Lists");
        Name rangeName = workbook.createName();
        rangeName.setNameName(name);
        String column = listSheet.findFreeColumn();
        listSheet.setValue(column, 1, name);
        for (int i = 0; i < values.length; i++){
            listSheet.setValue(column, i + 2, values[i]);
        }
        System.out.println (column);
        rangeName.setRefersToFormula("'Lists'!$" + column + "$2:$" + column + "$" + (values.length + 2));
    }

}
