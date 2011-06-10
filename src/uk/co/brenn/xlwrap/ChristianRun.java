/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package uk.co.brenn.xlwrap;

import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.exec.ExecutionContext;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import java.io.IOException;
import uk.co.brenn.xlwrap.WorkbookWrite;
import uk.co.brenn.xlwrap.XLWrapMapException;
import uk.co.brenn.xlwrap.expr.func.BrennRegister;

/**
 *
 * @author christian
 */
public class ChristianRun {

    //The file: bit is required by xlwrap whiuch can alos handle http urls.
    static private String  DROPBOX = "file:c:Dropbox/FishLink XLWrap data/";

    private static void loadXLS(ExecutionContext context, String name) throws XLWrapException, XLWrapEOFException, XLWrapMapException, IOException{
        //Adjust this file to the your local path
        WorkbookWrite.setRoots(DROPBOX + "Meta Data/",DROPBOX + "Raw Data/");
        WorkbookWrite mapWrite = new WorkbookWrite(context, name);
        mapWrite.writeMap();
        mapWrite.runMap();
    }
    public static void main(String[] args) throws XLWrapException, XLWrapEOFException, IOException, XLWrapMapException {
        ExecutionContext context = new ExecutionContext();
        BrennRegister.register();
        //loadXLS(context, "CumbriaTarnsPart1MetaData.xls");
        //loadXLS(context, "FBA_Tarn_MetaData.xls");
        //loadXLS(context, "RecordsMetaData1.xlsx");
        loadXLS(context, "TarnschemFinalMetaData1.xls");

    }

}