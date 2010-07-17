
class RubyInvocationProxy < com.jacob.com.InvocationProxy
  def initialize(target)
    super()
    @target = target
  end

  def invoke(name, parameters) # parameters is Variant[] always
    parms = parameters[1..-1].map {|p| VariantUtilities.variant_to_object(p) }
    @target.__send__ name, *parms
    nil # TODO I am guessing we need to return actual variant here
  end
end

class RubyDispatchEvents < com.jacob.com.DispatchEvents
  def initialize(source, event_sink, prog_id=nil)
    super(source, event_sink, prog_id)
  end
end

class WIN32OLE_EVENT
  def initialize(ole, event_name)
    @event_handlers = {}
    # TODO: Argument errors + specs

    if event_name.nil? # Default event name
      # TODO: get default event
    end
    
    dispatch = ole.dispatch
    proxy = RubyInvocationProxy.new(self)
    RubyDispatchEvents.new(dispatch, proxy, dispatch.program_id)
  end

  def on_event(name=nil, &block)
    if name
      @event_handlers[name.to_sym] = block
    else
      @default_handler = block
    end
  end

  def method_missing(name, *args)
    handler = @event_handlers[name]
    if handler
      handler.call *args
    elsif @default_handler
      @default_handler.call name, *args
    end
  end

  # Although this has an FFI version of MRI's message_loop it is very unlikely
  # to work since Jacob also has an event thread doing this same thing.  So
  # these extra calls do not hurt, but they also probably will not help either.
  def self.message_loop
    # TODO: Explicitly free message vs wait for GC?
    message = ::Win::User32::MSG.new
    while ::Win::User32::peek_message(message, nil, 0, 0, ::Win::User32::PM_REMOVE)
      p value, message
      ::Win::User32::translate_message(message)
      ::Win::User32::dispatch_message(message)
    end
  end
end
