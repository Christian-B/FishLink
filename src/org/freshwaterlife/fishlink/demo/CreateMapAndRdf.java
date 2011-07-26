package org.freshwaterlife.fishlink.demo;

import org.freshwaterlife.fishlink.FishLinkException;
import org.freshwaterlife.fishlink.FishLinkPaths;
import org.freshwaterlife.fishlink.xlwrap.WorkbookWrite;
import org.freshwaterlife.fishlink.xlwrap.expr.func.BrennRegister;
import org.freshwaterlife.fishlink.xlwrap.run.MapRun;

/**
 *
 * @author christian
 */
public class CreateMapAndRdf {

    private static void mapAndRdf(String dataUrl, String pid) throws FishLinkException{
        //Adjust this file to the your local path
        WorkbookWrite mapWrite = new WorkbookWrite(dataUrl, pid, FishLinkPaths.MASTER_FILE); 
        mapWrite.writeMap();
        MapRun.runMap(pid);
    }
    
    public static void main(String[] args) throws FishLinkException {
        BrennRegister.register();
        mapAndRdf(FishLinkPaths.META_FILE_ROOT + "CumbriaTarnsPart1MetaData.xls", "CTP1");
        mapAndRdf(FishLinkPaths.META_FILE_ROOT + "FBA_TarnsMetaData.xls", "FBA345");
        mapAndRdf(FishLinkPaths.META_FILE_ROOT + "RecordsMetaData.xls", "rec12564");
        mapAndRdf(FishLinkPaths.META_FILE_ROOT + "SpeciesMetaData.xls", "spec564");
        mapAndRdf(FishLinkPaths.META_FILE_ROOT + "StokoeMetaData.xls", "stokoe32433232");
        mapAndRdf(FishLinkPaths.META_FILE_ROOT + "TarnsMetaData.xls", "tarns33exdw2");
        mapAndRdf(FishLinkPaths.META_FILE_ROOT + "TarnschemFinalMetaData.xls", "TSF1234");
        mapAndRdf(FishLinkPaths.META_FILE_ROOT + "WillbyGroupsMetaData.xls", "wbgROUPS8734");
        /* */
    }
}