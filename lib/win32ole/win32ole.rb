class WIN32OLE
  # TODO: server_name, host, others are missing
  def initialize(id, *rest)
    @obj = Dispatch.new(id)
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

  def define_set(name)
    id = Dispatch.getIDOfName(@obj, name[0..-2])

    self.class.__send__(:define_method, name) do |*parms|
      Dispatch.put(@obj, id, *parms)
      nil # TODO: Should set's return nil?
    end
  end

  def define_method_or_get(name)
    id = Dispatch.getIDOfName(@obj, name)

    self.class.__send__(:define_method, name) do |*parms|
      variant_value(Dispatch.call(@obj, id, *parms))
    end
  end

  def variant_value(variant)
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
    when Variant::VariantEmpty
      nil
    when Variant::VariantBoolean
      variant.getBoolean
    else
      puts "other = #{variant.getvt}"
    end
  end
end
