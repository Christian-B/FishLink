package org.freshwaterlife.fishlink.demo;

import java.util.Date;

/**
 * This is just an example of how to manage a single instance of a class.
 * @author Christian
 */
public class SingletonDemo {
    
    static SingletonDemo onlyInstance;
    
    private Date creation;
    
    private SingletonDemo(){
       creation = new Date(); 
    }
    public static SingletonDemo getDemo(){
        if (onlyInstance == null){
            onlyInstance = new SingletonDemo();
        }
        return onlyInstance;
    }
    
    public Date getCreation(){
        return creation;
    }
    
    public static void main(String[] args) {
        SingletonDemo d1 = SingletonDemo.getDemo();
        System.out.println (d1.getCreation().getTime());
        SingletonDemo d2 = SingletonDemo.getDemo();
        System.out.println (d2.getCreation().getTime());
    }
}
