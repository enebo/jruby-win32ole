package org.jruby.ext.win32ole;

import com.jacob.com.Dispatch;
import com.jacob.com.EnumVariant;
import com.jacob.com.Variant;
import org.jruby.Ruby;
import org.jruby.RubyArray;
import org.jruby.RubyClass;
import org.jruby.RubyInteger;
import org.jruby.RubyObject;
import org.jruby.anno.JRubyMethod;
import org.jruby.internal.runtime.methods.DynamicMethod;
import org.jruby.javasupport.Java;
import org.jruby.javasupport.JavaObject;
import org.jruby.javasupport.JavaUtil;
import org.jruby.runtime.Block;
import org.jruby.runtime.ObjectAllocator;
import org.jruby.runtime.ThreadContext;
import org.jruby.runtime.builtin.IRubyObject;

/**
 */
public class RubyWIN32OLE extends RubyObject {
    public static ObjectAllocator WIN32OLE_ALLOCATOR = new ObjectAllocator() {
        public IRubyObject allocate(Ruby runtime, RubyClass klass) {
            return new RubyWIN32OLE(runtime, klass);
        }
    };
    
    public Dispatch dispatch = null;

    public RubyWIN32OLE(Ruby runtime, RubyClass metaClass) {
        super(runtime, metaClass);
    }

    private RubyWIN32OLE(Ruby runtime, RubyClass metaClass, Dispatch dispatch) {
        this(runtime, metaClass);

        this.dispatch = dispatch;
    }

    @JRubyMethod()
    public IRubyObject each(ThreadContext context, Block block) {
        Ruby runtime = context.getRuntime();
        EnumVariant enumVariant = dispatch.toEnumVariant();

        // FIXME: when no block is passed handling
        
        while (enumVariant.hasMoreElements()) {
            Variant value = enumVariant.nextElement();
            block.yield(context, fromVariant(runtime, getMetaClass(), value));
                    
        }

        return runtime.getNil();
    }

    @JRubyMethod(required = 3)
    public IRubyObject _getproperty(ThreadContext context, IRubyObject dispid,
            IRubyObject args, IRubyObject argTypes) {
        RubyArray argsArray = args.convertToArray();
        int argsArraySize = argsArray.size();
        int dispatchId = (int) RubyInteger.num2long(dispid);
        Variant returnValue;
        
        if (argsArraySize == 0) {
            returnValue = Dispatch.call(dispatch, dispatchId);
        } else {
            Object[] objectArgs = new Object[argsArraySize];
            for (int i = 0; i < argsArraySize; i++) {
                objectArgs[i] = toObject(argsArray.eltInternal(i));
            }
            returnValue = Dispatch.call(dispatch, dispatchId, objectArgs);
        }
        return fromVariant(context.getRuntime(), getMetaClass(), returnValue);
    }

    @JRubyMethod(required = 1, rest = true)
    public IRubyObject initialize(ThreadContext context, IRubyObject[] args) {
        String id = args[0].convertToString().asJavaString();

        dispatch = new Dispatch(toProgID(id));

        return this;
    }

    @JRubyMethod(required = 1, rest = true)
    public IRubyObject invoke(ThreadContext context, IRubyObject[] args) {
        return method_missing(context, args);
    }

    @JRubyMethod(required = 1, rest = true)
    public IRubyObject method_missing(ThreadContext context, IRubyObject[] args) {
        String methodName = args[0].asJavaString();
        
        if (methodName.endsWith("=")) return setupInvokeSet(context, methodName, args);

        return setupInvokeMethodOrGet(context, methodName, args);
    }

    @JRubyMethod(name = "[]", required = 1)
    public IRubyObject op_aref(ThreadContext context, IRubyObject property) {
        String propertyName = property.asJavaString();
        
        return fromVariant(context.getRuntime(), getSingletonClass(), Dispatch.get(dispatch, propertyName));
    }

    @JRubyMethod(name = "[]=", required = 2)
    public IRubyObject op_aset(ThreadContext context, IRubyObject property, IRubyObject value) {
        Ruby runtime = context.getRuntime();
        String propertyName = property.asJavaString();

        Dispatch.put(dispatch, propertyName, toObject(value));
        return runtime.getNil();
    }

    private IRubyObject setupInvokeSet(ThreadContext context, String methodName, IRubyObject[] args) {
        int id = Dispatch.getIDOfName(dispatch, methodName.substring(0, methodName.length() - 1));
        DynamicMethod method = new PutMethod(getSingletonClass(), dispatch, id);

        getSingletonClass().addMethod(methodName, method);

        IRubyObject[] newArgs = new IRubyObject[args.length - 1];
        System.arraycopy(args, 1, newArgs, 0, args.length - 1);
        return method.call(context, NEVER, metaClass, methodName, newArgs);
    }

    private IRubyObject setupInvokeMethodOrGet(ThreadContext context, String methodName, IRubyObject[] args) {
        int id = Dispatch.getIDOfName(dispatch, methodName);
        DynamicMethod method = new MethodOrGetMethod(getSingletonClass(), dispatch, id);

        getSingletonClass().addMethod(methodName, method);

        IRubyObject[] newArgs = new IRubyObject[args.length - 1];
        System.arraycopy(args, 1, newArgs, 0, args.length - 1);
        return method.call(context, NEVER, metaClass, methodName, newArgs);
    }

    @Override
    public Object toJava(Class klass) {
        return dispatch;
    }

    public static Object toObject(IRubyObject rubyObject) {
        return rubyObject.toJava(Object.class);
    }

    public static IRubyObject fromVariant(Ruby runtime, RubyClass metaClass, Variant variant) {
        switch (variant.getType()) {
            case Variant.VariantBoolean:
                return runtime.newBoolean(variant.getBoolean());
            case Variant.VariantDispatch:
                return new RubyWIN32OLE(runtime, metaClass, variant.getDispatch());
        }

        IRubyObject rubyObject = JavaUtil.convertJavaToUsableRubyObject(runtime, variant.toJavaObject());

        return rubyObject instanceof JavaObject ? Java.wrap(runtime, rubyObject) : rubyObject;
    }

    public static String toProgID(String id) {
        if (id != null && id.startsWith("{{") && id.endsWith("}}")) {
            return id.substring(2, id.length() - 2);
        }

        return id;
    }
}

