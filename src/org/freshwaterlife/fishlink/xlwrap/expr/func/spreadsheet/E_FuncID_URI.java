/**
 * Copyright 2009 Andreas Langegger, andreas@langegger.at, Austria
 * Copyright 2011 Christian Brenninkmeijer, Brenninc@cs.man.ac.uk, UK
 * 
 * Licensed under the Apache License, Version 2.0 (the "License"); 
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.freshwaterlife.fishlink.xlwrap.expr.func.spreadsheet;

import at.jku.xlwrap.common.XLWrapException;
import at.jku.xlwrap.exec.ExecutionContext;
import at.jku.xlwrap.map.expr.val.E_String;
import at.jku.xlwrap.map.expr.val.XLExprValue;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;
import org.freshwaterlife.fishlink.ZeroNullType;
import org.freshwaterlife.fishlink.FishLinkException;

/**
 * Create a URI based on an ID/ value found in a Cell.
 * 
 * Arguments expected are:
 * <ul>
 *    <li>The main part of the URI which does not depend on the Cell
 *    <li>The location of the Cell which contains the id
 *    <li>The ZeroNull setting for the id Cell
 *    <li>The Cell of the data currently being written
 *    <li>The ZeroNull setting for the data being written
 * </ul>
 * If the data being written is the id, then the last two arguments can be omitted rather than repeated.
 * <p>
 * If either the Id cell or the data cell (taking its ZeroNull setting into account) evaluates to null,
 *     the whole URI will be null and this cell will be ignored by XlWrap.
 * @author Christian
 *
 */
public class E_FuncID_URI extends E_Func_with_zero {

    /**
     * default constructor
     */
    public E_FuncID_URI() {
    }

    /**
     * Test if an expression evaluates to null when taking into account the associated ZeroNull setting. 
     * @param expression The expression to consider which may be null.
     * @param zeroToNullExpr String representation which can be parsed to a ZeroNull expression.
     * @return False for any expression that are already not zero and not null, regardless of ZeroNull Expression
     *         True If expression is null and ZeroNull does not convert nulls to Zero.
     *         True if expression is zero and ZeroNull converts zeros to nulls.
     *         False for other Zero/Null Expressions and ZeroNull values. 
     * @throws XLWrapException 
     */
    private boolean evaluatesToNull(XLExprValue<?> expression, XLExprValue<?> zeroToNullExpr) throws XLWrapException{
        Object value;
        if (expression == null){
            value = null;
        } else {
            value = expression.getValue();
        }
        String zeroToNullString = zeroToNullExpr.getValue().toString();
        ZeroNullType zeroNullType;
        try {
            zeroNullType = ZeroNullType.parse(zeroToNullString);
        } catch (FishLinkException ex) {
            throw new XLWrapException(ex);
        }
        switch (zeroNullType){
            case KEEP: 
                return value == null;
            case NULLS_AS_ZERO:
                return true;
            case ZEROS_AS_NULLS:
                if (value == null){
                    return true;
                }
                return isZero(value);
            default:
                throw new XLWrapException("Unexpected ZeroNullType: "+ zeroNullType);
        }
     }
    
    /**
     * Evaluate the cells returning either null or a concatenation of the main URI part and the specific ending.
     *  
     * This will check the Id/value cell and if available also the data cell.
     * <p>
     * Arguments described at top of class {@link E_FuncID_URI}
     * 
     * @param context XlWrap content required to evaluate expressions in.
     * @param specific The ending of the URI which changes depending on the Cell value and function.
     * @return Null is any cell is null according to:
     *    {@link #evaluatesToNull(at.jku.xlwrap.map.expr.val.XLExprValue, at.jku.xlwrap.map.expr.val.XLExprValue)}
     *    Otherwise a URI ending with the value of the ID Cell.
     * @throws XLWrapException
     * @throws XLWrapEOFException 
     */
    final XLExprValue<String> doEval(ExecutionContext context, String specific) throws XLWrapException, XLWrapEOFException {
        // ignores actual cell value, just use the range reference to determine row
        String url = getArg(0).eval(context).getValue().toString();

        //Check ID/Value column is not evaluatesToNull
        if (evaluatesToNull (getArg(1).eval(context), getArg(2).eval(context))){
            return null;
        }
        if (args.size() == 5){
            //Check Data column is not evaluatesToNull
            if (evaluatesToNull (getArg(3).eval(context), getArg(4).eval(context))){
                return null;
            }
        }
        return new E_String(url + specific);
    }

    /**
     * Evaluate by checking for null calls and creating a URI ending with the Cells value.
     * 
     * Arguments described at top of class {@link E_FuncID_URI}
     * @param context XlWrap content required to evaluate expressions in.
     * @return Null is any cell is null according to:
     *    {@link #evaluatesToNull(at.jku.xlwrap.map.expr.val.XLExprValue, at.jku.xlwrap.map.expr.val.XLExprValue)}
     *    Otherwise a URI ending with the value found in the id cell (args(1))
     * @throws XLWrapException
     * @throws XLWrapEOFException 
     */
    @Override
    public XLExprValue<String> eval(ExecutionContext context) throws XLWrapException, XLWrapEOFException {
        XLExprValue<?> value1 = getArg(1).eval(context);
        String valueString;
        if (value1 == null) {
            //Only time this is needed is is Nulls are converted to zero so just convert it here.
            //In other case null will be returned before this value is considered.
            valueString = "0";
        } else {
            valueString = value1.toString();
        }
        return doEval(context, valueString);
    }
	
}
