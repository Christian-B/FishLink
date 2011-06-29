package org.freshwaterlife.fishlink.xlwrap;

import at.jku.xlwrap.common.XLWrapException;

/**
 *
 * @author Christian
 */
public class XLWrapMapException extends Exception {

    /**
     * Constructs an instance of <code>XLWrapMapException</code> with the specified detail message.
     * @param msg the detail message.
     */
    public XLWrapMapException(String msg) {
        super(msg);
    }

    public XLWrapMapException(String msg, Exception ex) {
        super(msg, ex);
    }
}
