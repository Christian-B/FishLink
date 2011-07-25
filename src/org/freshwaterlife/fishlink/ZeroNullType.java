package org.freshwaterlife.fishlink;

import org.freshwaterlife.fishlink.xlwrap.Constants;
import org.freshwaterlife.fishlink.xlwrap.XLWrapMapException;

/**
 *
 * @author Christian
 */
public enum ZeroNullType {
    KEEP(Constants.KEEP_LABEL), 
    ZEROS_AS_NULLS (Constants.ZERO_AS_NULLS_LABEL), 
    NULLS_AS_ZERO(Constants.NULLS_AS_ZERO_LABEL);
        
    String text;
    
    private ZeroNullType(String label){
        text = label;
    }
    
    @Override
    public String toString(){
        return text;
    }
    
    public static ZeroNullType parse(String label) throws XLWrapMapException{
        if (label == null){
            return KEEP;
        }
        if (label.isEmpty()){
            return KEEP;
        }
        for (ZeroNullType zeroNullType : ZeroNullType.values()){
            if (label.equalsIgnoreCase(zeroNullType.text)){
                return zeroNullType;
            }
        }
        throw new XLWrapMapException("Unexpected ZeroVSNull type " + label);
    }
  
}
