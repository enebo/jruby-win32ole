package org.jruby.ext.win32ole;


import com.jacob.com.Dispatch;
import org.jruby.RubyModule;
import org.jruby.internal.runtime.methods.CallConfiguration;
import org.jruby.internal.runtime.methods.DynamicMethod;
import org.jruby.runtime.Block;
import org.jruby.runtime.ThreadContext;
import org.jruby.runtime.Visibility;
import org.jruby.runtime.builtin.IRubyObject;

public class PutMethod extends DynamicMethod {
    private final Dispatch dispatch;
    private final int dispatchId;

    public PutMethod(RubyModule impl, Dispatch dispatch, int dispatchId) {
        super(impl, Visibility.PUBLIC, CallConfiguration.FrameNoneScopeNone);
        this.dispatch = dispatch;
        this.dispatchId = dispatchId;
    }

    @Override
    public IRubyObject call(ThreadContext tc, IRubyObject iro, RubyModule rm, String string, IRubyObject[] args, Block block) {
        Dispatch.put(dispatch, dispatchId, RubyWIN32OLE.toObject(args[0]));
        return tc.getRuntime().getNil();
    }

    @Override
    public DynamicMethod dup() {
        throw new UnsupportedOperationException("Not supported yet.");
    }
}
