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
    
    public static synchronized PidStore padStoreFactory() throws IOException, XLWrapMapException{
        if (pidstore == null){
            return pidstore = new PidStore();
        }
        return pidstore;
    }
    
    private PidStore() throws IOException, XLWrapMapException{
        filesMap = new HashMap<String,String>();
        File file = new File(pidPath);
        if (file.exists()){
            FileReader reader = new FileReader(file);
            BufferedReader buffer = new BufferedReader(reader);
            String line = buffer.readLine();
            while (line != null) {
                String[] tokens = line.split (",");
                if (tokens.length != 2){
                    throw new XLWrapMapException("Pad register error! Found line " + line);
                }
                filesMap.put(tokens[0], tokens[1]);
                line = buffer.readLine();
            }
        } 
    }
    
    public synchronized String registerFile(File file) throws IOException{
        String path = file.getAbsolutePath();
        //UID uid = new UID();
        Random random = new Random();
        UUID uid = new UUID(random.nextLong(), random.nextLong());
        String pid = uid.toString();
        filesMap.put(pid, file.getAbsolutePath());
        File padFile = new File(pidPath);
        FileWriter writer = new FileWriter(padFile);
        BufferedWriter buffer = new BufferedWriter(writer);
        for(Entry<String, String> e: filesMap.entrySet()){
            buffer.write(e.getKey());
            buffer.write(",");
            buffer.write(e.getValue());
            buffer.newLine();
        }
        buffer.close();
        return pid;
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

    public static void main(String[] args) throws IOException, XLWrapMapException{
        //File file = new File(FishLinkPaths.MASTER_FILE);
        PidStore pidstore = padStoreFactory();
        //String pid = pidstore.registerFile(file);
        //System.out.println(pidstore.registerFile(file));
        System.out.println(pidstore.retreiveFile(MasterFactory.masterPid));
    }

}
