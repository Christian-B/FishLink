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

    private static void mapAndRdf(String metaPid, String dataPid) throws XLWrapMapException{
        //Adjust this file to the your local path
        WorkbookWrite mapWrite = new WorkbookWrite(metaPid, dataPid);
        mapWrite.writeMap();
        MapRun.runMap(dataPid);
    }
    
    public static void main(String[] args) throws XLWrapMapException {
        BrennRegister.register();
        mapAndRdf("META_CTP1", "CTP1");
        mapAndRdf("META_FBA345", "FBA345");
        mapAndRdf("META_rec12564", "rec12564");
        mapAndRdf("META_spec564", "spec564");
        mapAndRdf("META_stokoe32433232", "stokoe32433232");
        mapAndRdf("META_tarns33exdw2", "tarns33exdw2");
        mapAndRdf("META_TSF1234", "TSF1234");
        mapAndRdf("META_wbgROUPS8734", "wbgROUPS8734");
    }
}