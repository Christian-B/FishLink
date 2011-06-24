package org.freshwaterlife.fishlink;



import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import java.io.IOException;
import org.freshwaterlife.fishlink.xlwrap.WorkbookWrite;
import org.freshwaterlife.fishlink.xlwrap.XLWrapMapException;
import org.freshwaterlife.fishlink.xlwrap.expr.func.BrennRegister;
import org.freshwaterlife.fishlink.xlwrap.run.MapRun;

/**
 *
 * @author christian
 */
public class ChristianRun {

    private static void loadXLS(String name) throws XLWrapException, XLWrapEOFException, XLWrapMapException, IOException{
        //Adjust this file to the your local path
        WorkbookWrite.setRoots(FishLinkPaths.DROPBOX + "xMeta Data/",FishLinkPaths.DROPBOX + "Raw Data/");
        WorkbookWrite mapWrite = new WorkbookWrite(name);
        String doi = mapWrite.writeMap();
        MapRun.runMap(doi);
    }
    
    //public static void main(String[] args) throws XLWrapException, XLWrapEOFException, IOException, XLWrapMapException {
   //     BrennRegister.register();
   //     loadXLS("TarnschemFinalMetaData.xls");
   //     loadXLS("CumbriaTarnsPart1MetaData.xls");
   //     loadXLS("FBA_TarnsMetaData.xls");
   //     loadXLS("RecordsMetaData.xls");
   //     loadXLS("SpeciesMetaData.xls");
   //     loadXLS("StokoeMetaData.xls");
   //     loadXLS("TarnsMetaData.xls");
   //     loadXLS("WillbyGroupsMetaData.xls");

    //}

    public static void main(String[] args) throws XLWrapException, XLWrapEOFException, IOException, XLWrapMapException {
        BrennRegister.register();
        loadXLS("xTarnschemFinalMetaData.xls");
        loadXLS("xCumbriaTarnsPart1MetaData.xls");
        loadXLS("xFBA_TarnsMetaData.xls");
        loadXLS("xRecordsMetaData.xls");
        loadXLS("xSpeciesMetaData.xls");
        loadXLS("xStokoeMetaData.xls");
        loadXLS("xTarnsMetaData.xls");
        loadXLS("xWillbyGroupsMetaData.xls");

    }
}