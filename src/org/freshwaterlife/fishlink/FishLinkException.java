package org.freshwaterlife.fishlink;

import at.jku.xlwrap.common.XLWrapException;

/**
 *
 * @author Christian
 */
public class FishLinkException extends Exception {

    /**
     * Constructs an instance of <code>FishLinkException</code> with the specified detail message.
     * @param msg the detail message.
     */
    public FishLinkException(String msg) {
        super(msg);
    }

    public FishLinkException(String msg, Exception ex) {
        super(msg, ex);
    }
}
