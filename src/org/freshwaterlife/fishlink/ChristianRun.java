package org.freshwaterlife.fishlink;



import org.freshwaterlife.fishlink.xlwrap.WorkbookWrite;
import org.freshwaterlife.fishlink.xlwrap.XLWrapMapException;
import org.freshwaterlife.fishlink.xlwrap.expr.func.BrennRegister;
import org.freshwaterlife.fishlink.xlwrap.run.MapRun;

/**
 *
 * @author christian
 */
public class ChristianRun {

    private static void loadXLS(String name) throws XLWrapMapException{
        //Adjust this file to the your local path
        WorkbookWrite.setRoots(FishLinkPaths.META_DIR,FishLinkPaths.RAW_DIR);
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

    public static void main(String[] args) throws XLWrapMapException {
        BrennRegister.register();
        loadXLS("TarnschemFinalMetaData.xls");
        loadXLS("CumbriaTarnsPart1MetaData.xls");
        loadXLS("FBA_TarnsMetaData.xls");
        loadXLS("RecordsMetaData.xls");
        loadXLS("SpeciesMetaData.xls");
        loadXLS("StokoeMetaData.xls");
        loadXLS("TarnsMetaData.xls");
        loadXLS("WillbyGroupsMetaData.xls");

    }
}