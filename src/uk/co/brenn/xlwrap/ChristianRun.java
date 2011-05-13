/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package uk.co.brenn.xlwrap;

import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import java.io.IOException;
import uk.co.brenn.xlwrap.expr.func.BrennRegister;

/**
 *
 * @author christian
 */
public class ChristianRun {

    private static void loadXLS(String name) throws XLWrapException, XLWrapEOFException, XLWrapMapException, IOException{
        MapWrite mapWrite = new MapWrite("cyab/"+name+".xls", name+"doi");
        mapWrite.writeMap(name + ".trig");
        mapWrite.runMap(name + ".trig", name + ".rdf");
    }
    public static void main(String[] args) throws XLWrapException, XLWrapEOFException, IOException, XLWrapMapException {
        BrennRegister.register();
        loadXLS("surveys");
    }

}
