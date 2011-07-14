package org.freshwaterlife.fishlink;



import org.freshwaterlife.fishlink.xlwrap.NameChecker;
import org.freshwaterlife.fishlink.xlwrap.WorkbookWrite;
import org.freshwaterlife.fishlink.xlwrap.XLWrapMapException;
import org.freshwaterlife.fishlink.xlwrap.expr.func.BrennRegister;
import org.freshwaterlife.fishlink.xlwrap.run.MapRun;

/**
 *
 * @author christian
 */
public class ChristianRun {

    private static void mapAndRdf(String dataUrl, String pid) throws XLWrapMapException{
        //Adjust this file to the your local path
        WorkbookWrite mapWrite = new WorkbookWrite(dataUrl, pid, FishLinkPaths.MASTER_FILE); 
        mapWrite.writeMap();
        MapRun.runMap(pid);
    }
    
    public static void main(String[] args) throws XLWrapMapException {
        BrennRegister.register();
        mapAndRdf("file:c:\\Dropbox\\FishLink XLWrap data\\Meta Data\\CumbriaTarnsPart1MetaData.xls", "CTP1");
        mapAndRdf("file:c:\\Dropbox\\FishLink XLWrap data\\Meta Data\\FBA_TarnsMetaData.xls", "FBA345");
        mapAndRdf("file:c:\\Dropbox\\FishLink XLWrap data\\Meta Data\\RecordsMetaData.xls", "rec12564");
        mapAndRdf("file:c:\\Dropbox\\FishLink XLWrap data\\Meta Data\\SpeciesMetaData.xls", "spec564");
        mapAndRdf("file:c:\\Dropbox\\FishLink XLWrap data\\Meta Data\\StokoeMetaData.xls", "stokoe32433232");
        mapAndRdf("file:c:\\Dropbox\\FishLink XLWrap data\\Meta Data\\TarnsMetaData.xls", "tarns33exdw2");
        mapAndRdf("file:c:\\Dropbox\\FishLink XLWrap data\\Meta Data\\TarnschemFinalMetaData.xls", "TSF1234");
        mapAndRdf("file:c:\\Dropbox\\FishLink XLWrap data\\Meta Data\\WillbyGroupsMetaData.xls", "wbgROUPS8734");
        /* */
    }
}