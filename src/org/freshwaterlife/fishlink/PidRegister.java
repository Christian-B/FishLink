package org.freshwaterlife.fishlink;

import java.io.File;
import org.freshwaterlife.fishlink.xlwrap.XLWrapMapException;

public interface PidRegister {
    
    public String registerFile(File file) throws XLWrapMapException;
    
    public String retreiveFile(String pid) throws XLWrapMapException;

    public String retreiveFileOrNull(String pid) throws XLWrapMapException;
}
