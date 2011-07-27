package org.freshwaterlife.fishlink.demo;

import java.util.Date;

/**
 * This is just an example of how to manage a single instance of a class.
 * @author Christian
 */
public class FactoryDemo {
    
    static FactoryDemo onlyInstance;
    
    private Date creation;
    
    private FactoryDemo(){
       creation = new Date(); 
    }
    public static FactoryDemo getDemo(){
        if (onlyInstance == null){
            onlyInstance = new FactoryDemo();
        }
        return onlyInstance;
    }
    
    public Date getCreation(){
        return creation;
    }
    
    public static void main(String[] args) {
        FactoryDemo d1 = FactoryDemo.getDemo();
        System.out.println (d1.getCreation().getTime());
        FactoryDemo d2 = FactoryDemo.getDemo();
        System.out.println (d2.getCreation().getTime());
    }
}
