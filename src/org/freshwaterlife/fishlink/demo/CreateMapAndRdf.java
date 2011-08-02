package org.freshwaterlife.fishlink.demo;

import java.io.File;
import org.freshwaterlife.fishlink.FishLinkException;
import org.freshwaterlife.fishlink.FishLinkPaths;
import org.freshwaterlife.fishlink.xlwrap.WorkbookWrite;
import org.freshwaterlife.fishlink.xlwrap.expr.func.FishLinkToXlWrapRegister;
import org.freshwaterlife.fishlink.xlwrap.run.MapRun;

/**
 * Class Used by Christian to the creation of mapping files and rdf.
 * @author christian
 */
public class CreateMapAndRdf {

    /**
     * Creates the mapping file and rdf for a single file.
     * 
     * @param dataUrl URL (in xlwrap format) to the Workbook that holds the Annotated Data
     * @param pid Unique identifier to this data. Used in URIs and file naming.
     * @throws FishLinkException 
     */
    private static void mapAndRdf(String dataUrl, String pid) throws FishLinkException{
        //Adjust this file to the your local path
        WorkbookWrite mapWrite = new WorkbookWrite(dataUrl, pid, FishLinkPaths.MASTER_FILE); 
        File mapping = mapWrite.writeMap();
        MapRun.runMap(mapping, pid);
    }
    
    /**
     * Runs through all the test files.
     * 
     * @param args NONE
     * @throws FishLinkException 
     */
    public static void main(String[] args) throws FishLinkException {
        FishLinkToXlWrapRegister.register();
        mapAndRdf(FishLinkPaths.META_FILE_ROOT + "MiniMetaData.xls", "Mini");
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