package org.jruby.ext.win32ole;

import com.jacob.com.Dispatch;
import com.jacob.com.EnumVariant;
import com.jacob.com.Variant;
import java.util.Calendar;
import java.util.Date;
import org.jruby.Ruby;
import org.jruby.RubyArray;
import org.jruby.RubyClass;
import org.jruby.RubyInteger;
import org.jruby.RubyObject;
import org.jruby.RubyTime;
import org.jruby.anno.JRubyMethod;
import org.jruby.javasupport.Java;
import org.jruby.javasupport.JavaObject;
import org.jruby.javasupport.JavaUtil;
import org.jruby.runtime.Block;
import org.jruby.runtime.ObjectAllocator;
import org.jruby.runtime.ThreadContext;
import org.jruby.runtime.builtin.IRubyObject;
import win32ole.Win32oleService;

/**
 */
public class RubyWIN32OLE extends RubyObject {
    private static final Object[] EMPTY_OBJECT_ARGS = new Object[0];
    
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

    public Dispatch getDispatch() {
        return dispatch;
    }

    // Accessor for Ruby side of win32ole to get ahold of this object
    @JRubyMethod()
    public IRubyObject dispatch(ThreadContext context) {
        return JavaUtil.convertJavaToUsableRubyObject(context.getRuntime(), dispatch);
    }

    @JRubyMethod()
    public IRubyObject each(ThreadContext context, Block block) {
        Ruby runtime = context.getRuntime();
        EnumVariant enumVariant = dispatch.toEnumVariant();

        // FIXME: when no block is passed handling
        
        while (enumVariant.hasMoreElements()) {
            Variant value = enumVariant.nextElement();
            block.yield(context, fromVariant(runtime, value));
                    
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
        return fromVariant(context.getRuntime(), returnValue);
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
        
        if (methodName.endsWith("=")) return invokeSet(context, methodName, args);

        return invokeMethodOrGet(context, methodName, args);
    }

    @JRubyMethod()
    public IRubyObject ole_free(ThreadContext context) {
        dispatch.safeRelease();
        
        return context.getRuntime().getNil();
    }

    @JRubyMethod(name = "[]", required = 1)
    public IRubyObject op_aref(ThreadContext context, IRubyObject property) {
        String propertyName = property.asJavaString();
        
        return fromVariant(context.getRuntime(), Dispatch.get(dispatch, propertyName));
    }

    @JRubyMethod(name = "[]=", required = 2)
    public IRubyObject op_aset(ThreadContext context, IRubyObject property, IRubyObject value) {
        Ruby runtime = context.getRuntime();
        String propertyName = property.asJavaString();

        Dispatch.put(dispatch, propertyName, toObject(value));
        return runtime.getNil();
    }

    @JRubyMethod()
    public IRubyObject _setproperty(ThreadContext context, IRubyObject dispid,
            IRubyObject args, IRubyObject argTypes) {
        RubyArray argsArray = args.convertToArray();
        int argsArraySize = argsArray.size();
        int dispatchId = (int) RubyInteger.num2long(dispid);

        Object[] objectArgs = new Object[argsArraySize];
        for (int i = 0; i < argsArraySize; i++) {
            Object object = toObject(argsArray.eltInternal(i));
            System.out.println("OBJ: " + object.getClass().getSimpleName());
            objectArgs[i] = object;
        }
        // TODO: Maybe share these since we don't actually use them
        int[] errorArgs = new int[argsArraySize];
        
        Variant returnValue = Dispatch.invoke(dispatch, dispatchId, Dispatch.Put,
                objectArgs, errorArgs);

        return fromVariant(context.getRuntime(), returnValue);
    }

    private IRubyObject invokeSet(ThreadContext context, String methodName, IRubyObject[] args) {
        String comName = methodName.substring(0, methodName.length() - 1);
        
        // TODO set/put can have array of elements too....
        Dispatch.put(dispatch, comName, RubyWIN32OLE.toObject(args[1]));
        
        return context.getRuntime().getNil();
    }

    private Object[] makeObjectArgs(IRubyObject[] rubyArgs, int startIndex) {
        int length = rubyArgs.length;
        if (length - startIndex <= 0) return EMPTY_OBJECT_ARGS;

        Object[] objectArgs = new Object[length - startIndex];
        for (int i = startIndex; i < length; i++) {
            objectArgs[i - startIndex] = RubyWIN32OLE.toObject(rubyArgs[i]);
        }

        return objectArgs;
    }

    private IRubyObject invokeMethodOrGet(ThreadContext context, String methodName, IRubyObject[] args) {
        return fromVariant(context.getRuntime(),
                Dispatch.callN(dispatch, methodName, makeObjectArgs(args, 1)));
    }

    @Override
    public Object toJava(Class klass) {
        return dispatch;
    }

    public static Object toObject(IRubyObject rubyObject) {
        return rubyObject.toJava(Object.class);
    }

    public static IRubyObject fromVariant(Ruby runtime, Variant variant) {
        if (variant == null) return runtime.getNil();

        switch (variant.getType()) {
            case Variant.VariantBoolean:
                return runtime.newBoolean(variant.getBoolean());
            case Variant.VariantDispatch:
                return new RubyWIN32OLE(runtime, Win32oleService.getMetaClass(), variant.getDispatch());
            case Variant.VariantDate:
                return date2ruby(runtime, variant.getDate());
        }

        IRubyObject rubyObject = JavaUtil.convertJavaToUsableRubyObject(runtime, variant.toJavaObject());

        return rubyObject instanceof JavaObject ? Java.wrap(runtime, rubyObject) : rubyObject;
    }

    public static IRubyObject date2ruby(Ruby runtime, Date date) {
        Calendar cal = Calendar.getInstance();

        cal.setTime(date);

        return runtime.newTime(cal.getTimeInMillis());
    }

    public static String toProgID(String id) {
        if (id != null && id.startsWith("{{") && id.endsWith("}}")) {
            return id.substring(2, id.length() - 2);
        }

        return id;
    }
}

