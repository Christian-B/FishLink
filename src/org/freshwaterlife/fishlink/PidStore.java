/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package org.freshwaterlife.fishlink;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map.Entry;
import java.util.Random;
import java.util.UUID;
import org.freshwaterlife.fishlink.xlwrap.XLWrapMapException;

/**
 *
 * @author Christian
 */
public class PidStore implements PidRegister{

    static PidStore pidstore;
    
    static final String pidPath =  "data/padRegisrty.csv";
    
    HashMap<String,String> filesMap;
    
    public static synchronized PidStore padStoreFactory() throws XLWrapMapException{
        if (pidstore == null){
            return pidstore = new PidStore();
        }
        return pidstore;
    }
    
    private PidStore() throws XLWrapMapException{
        filesMap = new HashMap<String,String>();
        File file = new File(pidPath);
        if (file.exists()){
            try {
                FileReader reader = new FileReader(file);
                BufferedReader buffer = new BufferedReader(reader);
                String line = buffer.readLine();
                while (line != null) {
                    String[] tokens = line.split (",");
                    if (tokens.length != 2){
                        throw new XLWrapMapException("Pid register error! Found line " + line);
                    }
                    filesMap.put(tokens[0], tokens[1]);
                    line = buffer.readLine();
                }
            } catch (IOException ex) {
                throw new XLWrapMapException("Pid register error!", ex);
            }
        } 
    }
    
    public String registerFile(File file) throws XLWrapMapException{
        String path = file.getAbsolutePath();
        Random random = new Random();
        UUID uid = new UUID(random.nextLong(), random.nextLong());
        String pid = uid.toString();
        registerFile(path, pid);
        return pid;
     }
       
    public synchronized void registerFile(String path, String pid) throws XLWrapMapException{
        String check = filesMap.get(pid);
        if (check != null) {
            if (check.equals(path)){
                return; //Already there so do nothing
            } else {
                throw new XLWrapMapException ("Pid " + pid + " is already assigned to a different file");
            }
        }
        filesMap.put(pid, path);
        File padFile = new File(pidPath);
        try {
            FileWriter writer = new FileWriter(padFile);
            BufferedWriter buffer = new BufferedWriter(writer);
            for(Entry<String, String> e: filesMap.entrySet()){
                buffer.write(e.getKey());
                buffer.write(",");
                buffer.write(e.getValue());
                buffer.newLine();
            }
            buffer.close();
        } catch (IOException ex) {
            throw new XLWrapMapException("Pid register error!", ex);
        }
     }
       
   public String retreiveFile(String pid) throws XLWrapMapException {
        String path = retreiveFileOrNull(pid);
        if (path == null){
            throw new XLWrapMapException("Pid " + pid + "does not mapp to any know path");
        }
        return path;
    }
    
    public String retreiveFileOrNull(String pid) throws XLWrapMapException {
        return filesMap.get(pid);
    }
    
    public static void main(String[] args) throws XLWrapMapException{
        PidStore thePidstore = PidStore.padStoreFactory();
        thePidstore.registerFile("file:" + FishLinkPaths.OLD_META_DIR + "CumbriaTarnsPart1MetaData.xls", "OLDMETA_CTP1");
        thePidstore.registerFile("file:" + FishLinkPaths.OLD_META_DIR + "FBA_Tarns.xlsMetaData", "OLDMETA_FBA345");
        thePidstore.registerFile("file:" + FishLinkPaths.OLD_META_DIR + "Records.xlsMetaData", "OLDMETA_rec12564");
        thePidstore.registerFile("file:" + FishLinkPaths.OLD_META_DIR + "Species.xlsMetaData", "OLDMETA_spec564");
        thePidstore.registerFile("file:" + FishLinkPaths.OLD_META_DIR + "Stokoe.xlsMetaData", "OLDMETA_stokoe32433232");
        thePidstore.registerFile("file:" + FishLinkPaths.OLD_META_DIR + "Tarns.xlsMetaData", "OLDMETA_tarns33exdw2");
        thePidstore.registerFile("file:" + FishLinkPaths.OLD_META_DIR + "TarnschemFinalMetaData.xls", "OLDMETA_TSF1234");
        thePidstore.registerFile("file:" + FishLinkPaths.OLD_META_DIR + "WillbyGroupsMetaData.xls", "OLDMETA_wbgROUPS8734");
    }

}
