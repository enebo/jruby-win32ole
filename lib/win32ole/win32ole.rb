class WIN32OLE
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

  def each
    # TODO: Make EnumVariant have builtin each
    enum_variant = EnumVariant.new @dispatch

    while enum_variant.has_more_elements
      yield variant_value(enum_variant.next_element)
    end
  end

  class << self
    def connect(id)
      WIN32OLE.new to_progid(id)
    end

    def to_progid(id)
      return "clsid:#{$1}" if id =~ /^{(.*)}/
      id
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
      variant_value(Dispatch.call(@dispatch, id, *parms))
    end
  end

  def variant_value(variant)
    # TODO: Consider having to_ruby on all Variant types instead of this
    case(variant.getvt)
    when Variant::VariantInt
      variant.getInt
    when Variant::VariantString
      variant.getString
    when Variant::VariantDouble
      variant.getDouble
    when Variant::VariantFloat
      variant.getFloat
    when Variant::VariantShort
      variant.getShort
    when Variant::VariantEmpty, Variant::VariantNull
      nil
    when Variant::VariantBoolean
      variant.getBoolean
    when Variant::VariantDispatch
      WIN32OLE.new(variant.getDispatch)
    else
      puts "other = #{variant.getvt}"
    end
  end
end
