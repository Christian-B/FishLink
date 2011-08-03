package org.freshwaterlife.fishlink.demo;

import org.freshwaterlife.fishlink.FishLinkException;
import org.freshwaterlife.fishlink.FishLinkPaths;
import org.freshwaterlife.fishlink.metadatacreator.MetaDataCreator;

/**
 *  Class Used by Christian to test the creation of annotated sheets.
 * @author Christian
 */
public class CreateMetaData {
 
    /**
     * Used by Christian to test the creation of annotated sheets.
     * @param args NONE
     * @throws FishLinkException 
     */
    public static void main(String[] args) throws FishLinkException{
        MetaDataCreator creator = new MetaDataCreator(FishLinkPaths.MASTER_FILE);
        creator.createMetaData(FishLinkPaths.META_FILE_ROOT + "Mini.xlsx");
        /* 
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Raw Data\\CumbriaTarnsPart1.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Raw Data\\FBA_Tarns.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Raw Data\\Records.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Raw Data\\Species.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Raw Data\\Stokoe.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Raw Data\\Tarns.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Raw Data\\TarnschemFinal.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Raw Data\\WillbyGroups.xls");
        /* 
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Old Meta Data\\CumbriaTarnsPart1MetaData.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Old Meta Data\\FBA_TarnsMetaData.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Old Meta Data\\RecordsMetaData.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Old Meta Data\\SpeciesMetaData.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Old Meta Data\\StokoeMetaData.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Old Meta Data\\TarnsMetaData.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Old Meta Data\\TarnschemFinalMetaData.xls");
        creator.createMetaData("file:c:\\Dropbox\\FishLink XLWrap data\\Old Meta Data\\WillbyGroupsMetaData.xls");
        /* */
    }
}
