/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package org.freshwaterlife.fishlink;

/**
 *
 * @author christian
 */
public class FishLinkPaths {

   public static final String MAP_FILE_ROOT = "output/mappings/";

   public static final String RDF_FILE_ROOT = "output/rdf/";

   public static final String RDF_BASE_URL = "http://rdf.freshwaterlife.org/";

   public static final String MAIN_ROOT = "c:/Dropbox/FishLink XLWrap data/";

   public static final String META_DIR = MAIN_ROOT + "Meta Data/";
   
   //Old Meta data is optional - will not throw an error if missing
   public static final String OLD_META_DIR = MAIN_ROOT + "Old Meta Data/";
   
   public static final String RAW_DIR = MAIN_ROOT + "Raw Data/";

   //The file: bit is required by xlwrap whiuch can alos handle http urls.
   static public final String  DROPBOX = "file:c:/Dropbox/FishLink XLWrap data/";

   static public final String MASTER_FILE = "data/MetaMaster.xlsx";

}
