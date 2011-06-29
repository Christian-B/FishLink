package org.freshwaterlife.fishlink;

import java.io.File;
import java.io.IOException;
import org.freshwaterlife.fishlink.xlwrap.XLWrapMapException;

public interface PidRegister {
    
    public String registerFile(File file) throws IOException;
    
    public String retreiveFile(String pid) throws XLWrapMapException;
}
