package org.freshwaterlife.fishlink.metadatacreator;

import org.freshwaterlife.fishlink.FishLinkUtils;
import org.freshwaterlife.fishlink.xlwrap.XLWrapMapException;

/**
 *
 * @author Christian
 */
public class CreateMetaData {

    public static void usage(){
        FishLinkUtils.report("Creates a Meta Data collection file based on input.");
        FishLinkUtils.report("Requires two paramters.");
        FishLinkUtils.report("First is the raw data to be described.");
        FishLinkUtils.report("Second if the location (path and file name) of the MetaData file.");
    }

    public static void main(String[] args) throws XLWrapMapException{
        if (args.length != 2){
            usage();
            System.exit(2);
        }
        MetaDataCreator metaDataCreator = new MetaDataCreator();
        metaDataCreator.createMetaData(args[0], args[1]);
    }

}
