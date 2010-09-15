package org.jruby.ext.win32ole;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import org.jruby.RubyClass;
import org.jruby.RubyModule;
import org.jruby.internal.runtime.methods.CallConfiguration;
import org.jruby.internal.runtime.methods.DynamicMethod;
import org.jruby.runtime.Block;
import org.jruby.runtime.ThreadContext;
import org.jruby.runtime.Visibility;
import org.jruby.runtime.builtin.IRubyObject;

/**
 */
public class MethodOrGetMethod extends DynamicMethod {
    private Object[] EMPTY_ARGS = new Object[0];
    private final Dispatch dispatch;
    private final int dispatchId;
    
    public MethodOrGetMethod(RubyModule impl, Dispatch dispatch, int dispatchId) {
        super(impl, Visibility.PUBLIC, CallConfiguration.FrameNoneScopeNone);
        this.dispatch = dispatch;
        this.dispatchId = dispatchId;
    }

    @Override
    public IRubyObject call(ThreadContext context, IRubyObject iro, RubyModule rm, String string) {
        return RubyWIN32OLE.fromVariant(context.getRuntime(),
                (RubyClass) getImplementationClass(),
                Dispatch.call(dispatch, dispatchId));
    }

    @Override
    public IRubyObject call(ThreadContext context, IRubyObject self, RubyModule klazz, String name, Block block) {
        return RubyWIN32OLE.fromVariant(context.getRuntime(),
                (RubyClass) getImplementationClass(),
                Dispatch.call(dispatch, dispatchId));
    }


    @Override
    public IRubyObject call(ThreadContext context, IRubyObject iro, RubyModule rm, String string, IRubyObject[] args, Block block) {
        // N-arg path <- Jacob has no real unboxed paths currently.
        return RubyWIN32OLE.fromVariant(context.getRuntime(),
                (RubyClass) getImplementationClass(),
                Dispatch.callN(dispatch, dispatchId, createObjectArgs(args)));
    }

    private Object[] createObjectArgs(IRubyObject[] args) {
        if (args.length == 0) return EMPTY_ARGS;

        Object[] objectArgs = new Object[args.length];
        for (int i = 0; i < args.length; i++) {
            objectArgs[i] = RubyWIN32OLE.toObject(args[i]);
        }

        return objectArgs;
    }

    @Override
    public DynamicMethod dup() {
        throw new UnsupportedOperationException("Not supported yet.");
    }

}
