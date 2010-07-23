class WIN32OLE_TYPE
  module TypeHelper 
    def ole_methods_from_typeinfo(info, mask)
      ole_methods_sub(nil, info, methods, mask)
      type_info.impl_types_count.times do |i|
        href = type_info.get_ref_type_of_impl_type(i)
        ref_type_info = type_info.get_ref_type_info(href) # TODO: Failure mode?
        ole_methods_sub(type_info, ref_type_info, methods, mask)
      end
      methods
    end

    def ole_methods_sub(owner_typeinfo, typeinfo, methods, mask)
      type_info.funcs_count.times do |i|
        begin
          desc = type_info.get_func_desc(i)
          docs = type_info.get_documentation(desc.memid)
          # TODO: MASK CHECK
          methods << WIN32OLE_METHOD.new(nil, typeinfo, owner_typeinfo, docs.name, desc.memid)
        rescue ComFailException => e
          puts "ole_methods_sub: #{e}"
        end
      end
    end
  end

  def initialize(*args)
    if args.length == 3
      @typelib, @info, @docs = *args
    else
      @typelib_name, ole_name = *args
      tl = WIN32OLE_TYPELIB.new(@typelib, olename) # Internal call
      @info = nil
    end
  end

  def name
    @docs.name
  end

  def major_version
    @info.major_version
  end

  def minor_version
    @info.minor_version
  end

  def ole_methods
    ole_methods_from_typeinfo(info, nil)
  end

  def typekind
    @info.typekind
  end

  class << self
    # This is obsolete, but easy to emulate
    def typelibs
      WIN32OLE_TYPELIB.typelibs.collect {|t| t.name }
    end

    def ole_classes(tlib)
      WIN32OLE_TYPELIB.ole_classes(tlib)
    end
  end

  include TypeHelper
end
