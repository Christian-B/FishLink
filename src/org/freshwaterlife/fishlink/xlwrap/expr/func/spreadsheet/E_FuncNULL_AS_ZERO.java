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
import at.jku.xlwrap.map.expr.func.XLExprFunction;
import at.jku.xlwrap.map.expr.val.E_Long;
import at.jku.xlwrap.map.expr.val.XLExprValue;
import at.jku.xlwrap.spreadsheet.XLWrapEOFException;

/**
 * Function to convert Nulls to zero while leaving all other values as is.
 * @author Christian
 *
 */
public class E_FuncNULL_AS_ZERO extends XLExprFunction {

    /**
     * default constructor
     */
    public E_FuncNULL_AS_ZERO() {
    }

    /**
     * Function to convert Nulls to zero while leaving all other values as is.
     * @param context XlWrap content required to evaluate expressions in.
     * @return The expression in the first argument unchanged, 
     *         unless it evaluates to null in which case the long null is returned.
     * @throws XLWrapException
     * @throws XLWrapEOFException 
     */
    @Override
    public XLExprValue<?> eval(ExecutionContext context) throws XLWrapException, XLWrapEOFException {
        XLExprValue<?> value = getArg(0).eval(context);
        if (value.getValue() == null){
            return new E_Long(new Long(0));
        }
        return value;
    }
	
}
