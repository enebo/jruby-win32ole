package org.jruby.ext.win32ole;

import com.jacob.com.Dispatch;
import com.jacob.com.DispatchEvents;
import com.jacob.com.InvocationProxy;
import org.jruby.runtime.builtin.IRubyObject;

/**
 */
public class RubyDispatchEvents extends DispatchEvents {
    public RubyDispatchEvents(Dispatch dispatch, InvocationProxy proxy, String progId) {
        super(dispatch, proxy, progId);
    }

    public static RubyDispatchEvents setupDispatchEventHandler(Dispatch dispatch,
            IRubyObject eventHandler) {
        return new RubyDispatchEvents(dispatch, 
                new RubyInvocationProxy(eventHandler), dispatch.getProgramId());
    }
}
