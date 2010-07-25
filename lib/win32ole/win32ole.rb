# Shorthand vocabulary:
#  ti, oti - typeinfo or owner_typeinfo

class WIN32OLE
  attr_reader :dispatch

  CP_ACP = 0
  CP_OEMCP = 1
  CP_MACCP = 2
  CP_THREAD_ACP = 3
  CP_SYMBOL = 42
  CP_UTF7 = 65000
  CP_UTF8 = 65001

  # TODO: server_name, host, others are missing
  def initialize(id, *rest)
    @dispatch = Dispatch.new WIN32OLE.to_progid(SafeStringValue(id))
  rescue ComFailException
    raise WIN32OLERuntimeError
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
    from_variant(Dispatch.get(@dispatch, property_name))
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

  def ole_method(name)
    all_methods(type_info) do |ti, oti, desc, docs|
      if name == docs.name
        return WIN32OLE_METHOD.new(nil, name, ti, oti, desc.memid)
      end
      nil
    end
  end
  alias :ole_method_help :ole_method

  def ole_methods
    members = []
    all_methods(type_info) do |ti, oti, desc, docs|
      members << WIN32OLE_METHOD.new(nil, docs.name, ti, oti, desc.memid)
      nil
    end
    members
  end

  def _getproperty(dispid, args, arg_types)
    # TODO: What verification needs to happen with arg_types?
    from_variant(Dispatch.get(@dispatch, dispid, *args))
  end

  def _setproperty(dispid, args, arg_types)
    # TODO: What verification needs to happen with arg_types?
    Dispatch.put(@dispatch, dispid, *args)
  end

  # TODO: All these methods in MRI do many continues on error!!!

  def type_info
    @dispatch.type_info
  end

  class << self
    def codepage
      @@codepage ||= CP_ACP
    end

    def codepage=(new_codepage)
      @@codepage = new_codepage
    end

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
      id =~ /^{(.*)}/ ? "clsid:#{$1}" : id
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

  include WIN32OLE::Utils

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
