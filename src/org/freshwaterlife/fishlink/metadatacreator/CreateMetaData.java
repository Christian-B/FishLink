package org.freshwaterlife.fishlink.metadatacreator;

import org.freshwaterlife.fishlink.POI_Utils;
import org.freshwaterlife.fishlink.xlwrap.XLWrapMapException;

/**
 *
 * @author Christian
 */
public class CreateMetaData {

    public static void usage(){
        POI_Utils.report("Creates a Meta Data collection file based on input.");
        POI_Utils.report("Requires two paramters.");
        POI_Utils.report("First is the raw data to be described.");
        POI_Utils.report("Second if the location (path and file name) of the MetaData file.");
    }

    public static void main(String[] args) throws XLWrapMapException{
        if (args.length != 2){
            usage();
            System.exit(2);
        }
        MetaDataCreator metaDataCreator = new MetaDataCreator();
        metaDataCreator.prepareMetaDataOnTarget(args[0], args[1]);
    }

}
