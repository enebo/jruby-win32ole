
class RubyInvocationProxy < com.jacob.com.InvocationProxy
  def initialize(target)
    super()
    @target = target
  end

  def invoke(name, parameters) # parameters is Variant[] always
    @target.__send__ name, *parameters
  end
end

class RubyDispatchEvents < com.jacob.com.DispatchEvents
  def initialize(source, event_sink)
    super(source, event_sink)
  end
end

class WIN32OLE_EVENT
  def initialize(ole, event_name)
    @event_handlers = {}
    # TODO: Argument errors + specs

    if event_name.nil? # Default event name
      # TODO: get default event
    end
    
    RubyDispatchEvents.new(ole.dispatch, RubyInvocationProxy.new(self))
  end

  def on_event(name, &block)
    @event_handlers[name.to_sym] = block
  end

  def method_missing(name, *args)
    puts "Called '#{name}' #{args.join(',' )}"
    handler = @event_handlers[name]
    handler.call(args) if handler
  end
end
