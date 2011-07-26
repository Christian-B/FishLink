package org.freshwaterlife.fishlink.xlwrap.expr.func.spreadsheet;

import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.map.expr.func.XLExprFunction;
import java.util.Date;

/**
 * Abstract XLExprFunction which provides the isZero method.
 * @author Christian
 *
 */
public abstract class E_Func_with_zero extends XLExprFunction {

    /**
     * Check to see if an Object could be considered Zero.
     * 
     * Currently supported are subclasses of 
     * <ul>
     *     <li>Number in which case it looks for 0.0, 
     *     <li>String in which case it looks for "0",
     *     <li>Boolean which are never zero
     *     <li>Date which is zero if the getTime() methods returns zero.
     * </ul>
     * <p>
     * Additional types can be added as required.
     * 
     * @param value an Object one of the Expected type(s) or thir subtypes.
     * @return true If and only if the value could be considered zero, otherwise false. Nulls return false.
     * @throws XLWrapException Thrown if value is an unexpected type.
     */
    boolean isZero(Object value) throws XLWrapException{
        if (value == null){
            return false;
        }
        if (value instanceof Number){
            double number = ((Number)value).doubleValue();
            return number == 0.0;
        }
        if (value instanceof String){
            return value.toString().equals("0");
        }
        if (value instanceof Boolean){
            return false;
        }
        if (value instanceof Date){
            return (((Date)value).getTime() == 0);
        }
        throw new XLWrapException("Expected type " + value.getClass());
    }
    
}
