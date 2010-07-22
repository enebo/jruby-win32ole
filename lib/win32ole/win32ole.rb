class WIN32OLE
  include WIN32OLE::Utils
  attr_reader :dispatch

  # TODO: server_name, host, others are missing
  def initialize(id, *rest)
    @dispatch = Dispatch.new WIN32OLE.to_progid(id)
  end

  # Needs to support property gets and sets as well as methods
  def method_missing(method_name, *args)
    name = method_name.to_s # TODO: MRI uses symbol (rb_to_id to be specific)

    if name.end_with? '='
      define_set(name)
    else
      define_method_or_get(name)
    end

    __send__(name, *args)
  end

  # Returns the named property from the OLE automation object. 
  def [](property_name)
    VariantUtilities.variant_to_object(Dispatch.get(@dispatch, property_name))
  end

  # Sets the named property in the OLE automation object. 
  def []=(property_name, value)
    Dispatch.put(@dispatch, property_name, value)
  end

  # Iterates over each item of this OLE server that supports IEnumVARIANT 
  def each
    # TODO: Make EnumVariant have builtin each
    enum_variant = EnumVariant.new @dispatch

    while enum_variant.has_more_elements
      yield from_variant(enum_variant.next_element)
    end
  end

  def invoke(name, *args)
    method_missing(name, *args)
  end

  def ole_free
    @dispatch.safe_release
  end

  # TODO: Maybe make yielding impl so ole_method and ole_methods_sub can 
  # be combined
  def ole_method(name)
    type_info.funcs_count.times do |i|
      func_desc = type_info.get_func_desc(i)
      documentation = type_info.get_documentation(func_desc.memid)
      if name == documentation.name
        puts "FOUND #{name}1!!!"
        return WIN32OLE_METHOD.new nil, name, func_desc.memid, type_info
      end
    end
    nil
  end

  def ole_methods
    ole_methods_from_typeinfo(typeinfo_from_ole)
  end

  def ole_methods_sub(owner_type_info, type_info, methods, mask)
    # TODO: add owner_type_info to WIN32OLE_METHOD
    type_info.funcs_count.times do |i|
      begin
        func_desc = type_info.get_func_desc(i)
        docs = type_info.get_documentation(func_desc.memid)
        # TODO: MASK CHECK
        methods << WIN32OLE_METHOD.new(nil, docs.name, func_desc.memid, type_info)
      rescue ComFailException => e
      end
    end
  end

  def ole_methods_from_typeinfo(type_info, mask=nil)
    methods = []
    puts "TOTAL IMPL_TYPES #{type_info.impl_types_count}"
    ole_methods_sub(nil, type_info, methods, mask)
    type_info.impl_types_count.times do |i|
      href = type_info.get_ref_type_of_impl_type(i)
      ref_type_info = type_info.get_ref_type_info(href) # TODO: Failure mode?
      ole_methods_sub(type_info, ref_type_info, methods, mask)
    end
    methods
  end

  # TODO: All these methods in MRI do many continues on error!!!

  # TODO: Implement typeinfo_from_ole
  def typeinfo_from_ole
    type_info = @dispatch.type_info
    docs = type_info.get_documentation(-1)
    type_lib = type_info.get_containing_type_lib
    type_lib.get_type_info_count.times do |i|
      begin
        ti = type_lib.get_type_info(i)
        tdocs = type_lib.get_documentation(i)
        puts "#{docs.name} == #{tdocs.name}"
        return typelib.get_type_info(i) if tdocs.name == docs.name
      rescue ComFailException => e
        # We continue on failure. 
      end
    end
    type_info # Actually MRI seems like it could fail in weird case
  end

  def type_info
    @dispatch.type_info
  end

  class << self
    def connect(id)
      WIN32OLE.new to_progid(id)
    end

    def const_load(ole, a_class=WIN32OLE)
      constants = {}
      ole.type_info.containing_type_lib.type_info.to_a.each do |info|
        info.vars_count.times do |i|
          var_desc = info.get_var_desc(i)
          # TODO: Missing some additional flag checks to limit no. of constants
          if var_desc.constant
            name = first_var_name(info, var_desc)
            name = name[0].chr.upcase + name[1..-1] if name
            if constant?(name)
              a_class.const_set name, var_desc.constant
            else # vars which don't start [A-Z]?
              constants[name] = var_desc.constant
            end
          end
        end
      end
      a_class.const_set 'CONSTANTS', constants
      nil
    end

    def to_progid(id)
      return "clsid:#{$1}" if id =~ /^{(.*)}/
      id
    end

    private

    def constant?(name)
      name =~ /^[A-Z]/
    end

    def first_var_name(type_info, var_desc)
      type_info.get_names(var_desc.memid)[0]
    rescue
      nil
    end
  end

  private

  def define_set(name)
    id = Dispatch.getIDOfName(@dispatch, name[0..-2])

    self.class.__send__(:define_method, name) do |*parms|
      Dispatch.put(@dispatch, id, *parms)
      nil # TODO: Should set's return nil?
    end
  end

  def define_method_or_get(name)
    id = Dispatch.getIDOfName(@dispatch, name)

    self.class.__send__(:define_method, name) do |*parms|
      from_variant(Dispatch.call(@dispatch, id, *parms))
    end
  end
end
