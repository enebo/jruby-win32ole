package win32ole;

import java.io.IOException;import org.jruby.Ruby;
import org.jruby.RubyClass;
import org.jruby.ext.win32ole.RubyWIN32OLE;
import org.jruby.runtime.load.BasicLibraryService;

public class Win32oleService implements BasicLibraryService {
    public boolean basicLoad(Ruby runtime) throws IOException {
        System.out.println("LOADING WIN32OLE");
        RubyClass object = runtime.getObject();
        RubyClass win32ole = runtime.defineClass("WIN32OLE", object,
                RubyWIN32OLE.WIN32OLE_ALLOCATOR);

        win32ole.defineAnnotatedMethods(RubyWIN32OLE.class);

        return true;
    }
}
