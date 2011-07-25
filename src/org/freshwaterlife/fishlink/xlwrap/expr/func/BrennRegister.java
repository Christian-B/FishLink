package org.freshwaterlife.fishlink.xlwrap.expr.func;

import at.jku.xlwrap.map.expr.func.FunctionRegistry;
import org.freshwaterlife.fishlink.xlwrap.expr.func.spreadsheet.*;

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
        functionClass = E_FuncNULL_AS_ZERO.class;
        FunctionRegistry.register(functionClass);
        functionClass = E_FuncZERO_AS_NULL.class;
        FunctionRegistry.register(functionClass);
     }
}
