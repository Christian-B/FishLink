/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package uk.co.brenn.xlwrap;

import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.exec.ExecutionContext;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import java.io.IOException;
import uk.co.brenn.xlwrap.expr.func.BrennRegister;

/**
 *
 * @author christian
 */
public class ChristianRun {

    private static void loadXLS(ExecutionContext context, String name) throws XLWrapException, XLWrapEOFException, XLWrapMapException, IOException{
        //Adjust this file to the your local path
        WorkbookWrite.setRoots("file:c:Dropbox/FishLink XLWrap data/Meta Data/","file:c:Dropbox/FishLink XLWrap data/Raw Data/");
        WorkbookWrite mapWrite = new WorkbookWrite(context, name);
        mapWrite.writeMap();
        mapWrite.runMap();
    }
    public static void main(String[] args) throws XLWrapException, XLWrapEOFException, IOException, XLWrapMapException {
        ExecutionContext context = new ExecutionContext();
        BrennRegister.register();
        loadXLS(context, "CumbriaTarnsPart1MetaData.xls");
    }

}
