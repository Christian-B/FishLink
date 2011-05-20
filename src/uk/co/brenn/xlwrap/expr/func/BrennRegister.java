/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package uk.co.brenn.xlwrap.expr.func;

import at.jku.xlwrap.map.expr.func.FunctionRegistry;
import uk.co.brenn.xlwrap.expr.func.spreadsheet.*;

/**
 *
 * @author Christian
 */
public class BrennRegister {

    public static void register(){
        Class functionClass = E_FuncROW_URI.class;
        FunctionRegistry.register(functionClass);
        functionClass = E_FuncCELL_URI.class;
        FunctionRegistry.register(functionClass);
        functionClass = E_FuncID_URI.class;
        FunctionRegistry.register(functionClass);
        functionClass = E_FuncOTHER_ID_URI.class;
        FunctionRegistry.register(functionClass);
        functionClass = E_FuncSHEET_REFERENCE.class;
        FunctionRegistry.register(functionClass);
    }
}
