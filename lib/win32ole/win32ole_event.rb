
class RubyInvocationProxy < com.jacob.com.InvocationProxy
  include WIN32OLE::Utils

  def initialize(target)
    super()
    @target = target
  end

  def invoke(name, parameters) # parameters is Variant[] always
    @target.__send__ name, *parameters.map {|p| from_variant(p) }
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

  # Almost noop this.  We don't because it get CPU hot when people put this
  # in a hot loop!
  def self.message_loop
    sleep 0.1
  end
end
